// Workbook model over a lossless package.
//
// Load: parse workbook, styles, shared strings and every worksheet into a cell model (values,
// formulas, styles, column widths, row heights, merges, frozen panes, selection). Edits change the
// model and remember what changed. Save: every untouched part keeps its bytes; edited worksheets are
// rewritten by splicing only the edited cells (and edited row/column/merge/pane markup) into the
// original XML; new cell formats are appended to styles.xml. A save without edits returns a package
// that is byte-identical to the input.

import { ZipPackage } from './zip.ts';
import { parseXml, kid, kids, attrNum, attrBool, textOf, escText, escAttr, applySplices, retag, type XDocument, type XNode, type Splice } from './xml.ts';
import { Styles, parseTheme, fontXml, fillXml, borderXml, xfXml, type Xf, type ResolvedStyle, type Font, type Border, type Align } from './styles.ts';
import { parseFormula, shiftFormula, Evaluator, isErr, formulaRefs, parseDateText, ERR, type Scalar, type Node, type XlError } from './formula.ts';
import { parseAddr, parseRange, addrName, colName, key, unkey, norm, inRange, rangeName, type Range, type Addr } from './ref.ts';
import { formatValue, isDateFormat, partsToSerial, type Formatted } from './numfmt.ts';
import { protectionXml, type SheetProtection } from './protection.ts';

export interface Cell {
  v: Scalar;
  /** formula text without '=' */
  f?: string;
  s: number;
  /** shared-formula master info (kept so untouched children stay valid) */
  shared?: { si: string; master: boolean; ref?: string };
  array?: string;
}

export interface ColInfo { min: number; max: number; width?: number; hidden?: boolean; customWidth?: boolean; style?: number; raw?: Record<string, string> }
export interface RowInfo { ht?: number; hidden?: boolean; customHeight?: boolean; s?: number }

export interface Sheet {
  name: string;
  part: string;
  rId: string;
  sheetId: number;
  state: string;
  cells: Map<number, Cell>;
  cols: ColInfo[];
  rows: Map<number, RowInfo>;
  merges: Range[];
  freeze: { rows: number; cols: number } | null;
  topLeft: Addr;
  active: Addr;
  sqref: string;
  defaultRowHt: number; // points
  defaultColWidth: number; // characters incl. padding
  showGrid: boolean;
  zoom: number;
  tabColor?: string;
  maxRow: number;
  maxCol: number;
  doc: XDocument | null; // original XML (null for new sheets; a copied sheet's own XML for Move or Copy copies)
  isNew?: boolean;
  /** <sheetProtection> attributes while the sheet is protected (sheet="1"), else null */
  protection: SheetProtection | null;
  /** Move or Copy copies: the relationships part to write next to the new sheet part (external links only) */
  copyRels?: string;
}

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WS_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const WS_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';

export interface Edits {
  cells: Map<number, Set<number>>; // sheet index -> keys
  cols: Set<number>;
  rows: Map<number, Set<number>>;
  merges: Set<number>;
  view: Set<number>;
  protection: Set<number>;
  workbook: boolean;
}

export class Workbook {
  zip: ZipPackage;
  wbPart = 'xl/workbook.xml';
  wbDoc!: XDocument;
  styles!: Styles;
  theme: string[] = [];
  sst: string[] = [];
  sheets: Sheet[] = [];
  date1904 = false;
  activeTab = 0;
  definedNames = new Map<string, string>();
  removedSheets: Sheet[] = [];
  edits: Edits = { cells: new Map(), cols: new Set(), rows: new Map(), merges: new Set(), view: new Set(), protection: new Set(), workbook: false };
  /** the sheet that was active on disk (it carries tabSelected); activeTab follows it when sheets move */
  private diskActive: Sheet | null = null;
  /** sheet order changed (Move or Copy): workbook.xml rewrites activeTab and localSheetId */
  private reordered = false;
  /** scenario clock for TODAY()/NOW() */
  nowMs = Date.parse('2026-09-30T21:05:00Z');
  private computed = new Map<string, Scalar>();
  private computing = new Set<string>();
  private parsed = new Map<string, Node | null>();

  static load(bytes: Uint8Array): Workbook {
    const wb = new Workbook(ZipPackage.open(bytes));
    wb.parse();
    return wb;
  }

  constructor(zip: ZipPackage) { this.zip = zip; }

  // ---------------------------------------------------------------- parsing

  private xml(name: string): XDocument | null {
    return this.zip.has(name) ? parseXml(this.zip.readText(name)) : null;
  }

  private rels(source: string): Map<string, { type: string; target: string; node: XNode }> {
    const i = source.lastIndexOf('/');
    const name = `${source.slice(0, i + 1)}_rels/${source.slice(i + 1)}.rels`;
    const doc = this.xml(name);
    const out = new Map<string, { type: string; target: string; node: XNode }>();
    if (!doc) return out;
    for (const r of kids(doc.root, 'Relationship')) {
      const t = r.attrs.Target ?? '';
      const target = t.startsWith('/') ? t.slice(1) : normPath(source.slice(0, i + 1) + t);
      out.set(r.attrs.Id, { type: r.attrs.Type ?? '', target, node: r });
    }
    return out;
  }

  private parse() {
    const root = this.xml('_rels/.rels');
    const off = root ? kids(root.root, 'Relationship').find(r => /officeDocument$/.test(r.attrs.Type ?? '')) : undefined;
    if (off) this.wbPart = (off.attrs.Target ?? 'xl/workbook.xml').replace(/^\//, '');
    this.wbDoc = this.xml(this.wbPart)!;
    if (!this.wbDoc) throw new Error('not a workbook');
    const rels = this.rels(this.wbPart);
    let themeDoc: XDocument | null = null, stylesDoc: XDocument | null = null;
    for (const r of rels.values()) {
      if (/\/theme$/.test(r.type)) themeDoc = this.xml(r.target);
      else if (/\/styles$/.test(r.type)) { stylesDoc = this.xml(r.target); this.stylesPart = r.target; }
      else if (/\/sharedStrings$/.test(r.type)) {
        const d = this.xml(r.target);
        if (d) this.sst = kids(d.root, 'si').map(si => siText(d, si));
      }
    }
    this.theme = parseTheme(themeDoc);
    this.styles = new Styles(stylesDoc, this.theme);
    const wb = this.wbDoc.root;
    this.date1904 = attrBool(kid(wb, 'workbookPr'), 'date1904');
    this.activeTab = attrNum(kid(kid(wb, 'bookViews'), 'workbookView'), 'activeTab', 0);
    for (const dn of kids(kid(wb, 'definedNames'), 'definedName')) {
      if (dn.attrs.localSheetId === undefined && dn.attrs.name) this.definedNames.set(dn.attrs.name.toUpperCase(), textOf(this.wbDoc, dn));
    }
    for (const s of kids(kid(wb, 'sheets'), 'sheet')) {
      const rId = s.attrs['r:id'] ?? Object.entries(s.attrs).find(([k]) => k.endsWith(':id'))?.[1] ?? '';
      const rel = rels.get(rId);
      if (!rel || !/\/worksheet$/.test(rel.type)) continue; // chartsheets are not shown
      this.sheets.push(this.parseSheet(s.attrs.name ?? 'Sheet', rel.target, rId, attrNum(s, 'sheetId', this.sheets.length + 1), s.attrs.state ?? 'visible'));
    }
    if (this.activeTab >= this.sheets.length) this.activeTab = 0;
    this.diskActive = this.sheets[this.activeTab] ?? null;
  }

  stylesPart = 'xl/styles.xml';

  private parseSheet(name: string, part: string, rId: string, sheetId: number, state: string): Sheet {
    const doc = this.xml(part)!;
    const ws = doc.root;
    const sh: Sheet = {
      name, part, rId, sheetId, state, cells: new Map(), cols: [], rows: new Map(), merges: [], freeze: null,
      topLeft: { r: 0, c: 0 }, active: { r: 0, c: 0 }, sqref: 'A1', defaultRowHt: 15, defaultColWidth: 8.43 + 0.71, showGrid: true, zoom: 100,
      maxRow: 0, maxCol: 0, doc, protection: null,
    };
    const prot = kid(ws, 'sheetProtection');
    if (prot && attrBool(prot, 'sheet')) sh.protection = { ...prot.attrs };
    const fmt = kid(ws, 'sheetFormatPr');
    if (fmt) {
      sh.defaultRowHt = attrNum(fmt, 'defaultRowHeight', 15);
      if (fmt.attrs.defaultColWidth) sh.defaultColWidth = attrNum(fmt, 'defaultColWidth', 9.14);
      else if (fmt.attrs.baseColWidth) sh.defaultColWidth = attrNum(fmt, 'baseColWidth', 8) + 1.14;
    }
    const pr = kid(ws, 'sheetPr');
    const tab = kid(pr, 'tabColor');
    if (tab) sh.tabColor = tab.attrs.rgb ? '#' + tab.attrs.rgb.slice(-6) : undefined;
    const view = kid(kid(ws, 'sheetViews'), 'sheetView');
    if (view) {
      sh.showGrid = view.attrs.showGridLines !== '0';
      sh.zoom = attrNum(view, 'zoomScale', 100);
      if (view.attrs.topLeftCell) sh.topLeft = parseAddr(view.attrs.topLeftCell) ?? sh.topLeft;
      const pane = kid(view, 'pane');
      if (pane && (pane.attrs.state === 'frozen' || pane.attrs.state === 'frozenSplit')) sh.freeze = { rows: attrNum(pane, 'ySplit', 0), cols: attrNum(pane, 'xSplit', 0) };
      const sels = kids(view, 'selection');
      const sel = sels.find(s => !pane || s.attrs.pane === pane.attrs.activePane) ?? sels[0];
      if (sel?.attrs.activeCell) { sh.active = parseAddr(sel.attrs.activeCell) ?? sh.active; sh.sqref = sel.attrs.sqref ?? sel.attrs.activeCell; }
    }
    for (const c of kids(kid(ws, 'cols'), 'col')) {
      sh.cols.push({ min: attrNum(c, 'min', 1) - 1, max: attrNum(c, 'max', 1) - 1, width: c.attrs.width ? attrNum(c, 'width', 9) : undefined, hidden: attrBool(c, 'hidden'), customWidth: attrBool(c, 'customWidth'), style: c.attrs.style ? attrNum(c, 'style', 0) : undefined, raw: c.attrs });
    }
    let rowIdx = -1;
    const shared = new Map<string, { f: string; r: number; c: number }>();
    const pendingShared: { k: number; si: string; r: number; c: number }[] = [];
    for (const row of kids(kid(ws, 'sheetData'), 'row')) {
      rowIdx = row.attrs.r ? Number(row.attrs.r) - 1 : rowIdx + 1;
      const info: RowInfo = {};
      if (row.attrs.ht) info.ht = Number(row.attrs.ht);
      if (attrBool(row, 'hidden')) info.hidden = true;
      if (attrBool(row, 'customHeight')) info.customHeight = true;
      if (row.attrs.s && attrBool(row, 'customFormat')) info.s = Number(row.attrs.s);
      if (Object.keys(info).length) sh.rows.set(rowIdx, info);
      let colIdx = -1;
      for (const c of kids(row, 'c')) {
        const a = c.attrs.r ? parseAddr(c.attrs.r) : null;
        colIdx = a ? a.c : colIdx + 1;
        const r = a ? a.r : rowIdx;
        const cell: Cell = { v: null, s: attrNum(c, 's', 0) };
        const t = c.attrs.t ?? 'n';
        const vEl = kid(c, 'v');
        const fEl = kid(c, 'f');
        const raw = vEl ? textOf(doc, vEl) : null;
        if (t === 'inlineStr') { const is = kid(c, 'is'); cell.v = is ? siText(doc, is) : ''; }
        else if (raw !== null) {
          if (t === 's') cell.v = this.sst[Number(raw)] ?? '';
          else if (t === 'b') cell.v = raw === '1' || raw === 'true';
          else if (t === 'e') cell.v = { error: raw };
          else if (t === 'str') cell.v = raw;
          else if (t === 'd') { const d = /^(\d{4})-(\d{2})-(\d{2})/.exec(raw); cell.v = d ? partsToSerial(Number(d[1]), Number(d[2]), Number(d[3]), this.date1904) : raw; }
          else cell.v = Number(raw);
        }
        if (fEl) {
          const ft = fEl.attrs.t;
          const text = textOf(doc, fEl);
          if (ft === 'shared') {
            const si = fEl.attrs.si ?? '';
            if (text) { shared.set(si, { f: text, r, c: colIdx }); cell.f = text; cell.shared = { si, master: true, ref: fEl.attrs.ref }; }
            else { cell.shared = { si, master: false }; pendingShared.push({ k: key(r, colIdx), si, r, c: colIdx }); }
          } else if (text) {
            cell.f = text;
            if (ft === 'array') cell.array = fEl.attrs.ref;
          }
        }
        if (cell.v === null && !cell.f && !c.attrs.s && !cell.shared) continue;
        sh.cells.set(key(r, colIdx), cell);
        if (cell.v !== null || cell.f) { sh.maxRow = Math.max(sh.maxRow, r + 1); sh.maxCol = Math.max(sh.maxCol, colIdx + 1); }
      }
    }
    for (const p of pendingShared) {
      const m = shared.get(p.si);
      const cell = sh.cells.get(p.k);
      if (m && cell) cell.f = shiftFormula(m.f, p.r - m.r, p.c - m.c);
    }
    for (const m of kids(kid(ws, 'mergeCells'), 'mergeCell')) { const g = parseRange(m.attrs.ref ?? ''); if (g) sh.merges.push(g); }
    return sh;
  }

  // ---------------------------------------------------------------- reading

  sheetIndex(name: string): number { return this.sheets.findIndex(s => s.name.toLowerCase() === name.toLowerCase()); }

  cell(si: number, r: number, c: number): Cell | undefined { return this.sheets[si]?.cells.get(key(r, c)); }

  /** displayed (computed) value */
  value(si: number, r: number, c: number): Scalar {
    const cell = this.cell(si, r, c);
    if (!cell) return null;
    if (!cell.f) return cell.v;
    const k = `${si}:${r}:${c}`;
    const hit = this.computed.get(k);
    if (hit !== undefined) return hit;
    if (this.computing.has(k)) return 0; // circular reference: Excel shows 0 and warns
    this.computing.add(k);
    let out: Scalar;
    try {
      let ast = this.parsed.get(cell.f);
      if (ast === undefined) { try { ast = parseFormula(cell.f); } catch { ast = null; } this.parsed.set(cell.f, ast); }
      if (!ast) out = ERR.name;
      else {
        const ev = new Evaluator(this.context(si, r, c));
        const v = ev.eval(ast);
        out = ev.scalar(Array.isArray(v) ? v : v as Scalar);
        if (Array.isArray(v)) out = v[0]?.[0] ?? null;
        if (out === null) out = 0;
      }
    } catch (e) {
      out = isErr(e) ? e as XlError : ERR.value;
    } finally {
      this.computing.delete(k);
    }
    this.computed.set(k, out);
    return out;
  }

  private context(si: number, r: number, c: number) {
    return {
      cell: (sheet: string, rr: number, cc: number) => { const i = this.sheetIndex(sheet); return i < 0 ? ERR.ref : this.value(i, rr, cc); },
      extent: (sheet: string) => { const s = this.sheets[this.sheetIndex(sheet)]; return { rows: s?.maxRow ?? 0, cols: s?.maxCol ?? 0 }; },
      sheet: this.sheets[si].name, r, c,
      hasSheet: (n: string) => this.sheetIndex(n) >= 0,
      name: (n: string) => { const d = this.definedNames.get(n.toUpperCase()); if (!d) return null; try { return parseFormula(d); } catch { return null; } },
      nowMs: this.nowMs, date1904: this.date1904,
    };
  }

  recalc() { this.computed.clear(); }

  style(xf: number): ResolvedStyle { return this.styles.resolve(xf); }

  /** cell format index used for (r, c): the cell's own, else the row's, else the column's */
  xfAt(si: number, r: number, c: number): number {
    const sh = this.sheets[si];
    const cell = sh.cells.get(key(r, c));
    if (cell) return cell.s;
    const row = sh.rows.get(r);
    if (row?.s !== undefined) return row.s;
    const col = sh.cols.find(x => c >= x.min && c <= x.max);
    return col?.style ?? 0;
  }

  /** text as Excel shows it in the grid */
  display(si: number, r: number, c: number): Formatted & { value: Scalar } {
    const v = this.value(si, r, c);
    const st = this.style(this.xfAt(si, r, c));
    if (v === null) return { text: '', value: v };
    if (isErr(v)) return { text: v.error, value: v, align: undefined };
    const f = formatValue(v, st.numFmt, this.date1904);
    return { ...f, value: v };
  }

  /** what the formula bar shows */
  editText(si: number, r: number, c: number): string {
    const cell = this.cell(si, r, c);
    if (!cell) return '';
    if (cell.f) return '=' + cell.f;
    const v = cell.v;
    if (v === null) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (isErr(v)) return v.error;
    const st = this.style(cell.s);
    if (typeof v === 'number') {
      if (isDateFormat(st.numFmt)) {
        const hasTime = /[hs]/i.test(st.numFmt.replace(/"[^"]*"/g, ''));
        const hasDate = /[dy]/i.test(st.numFmt.replace(/"[^"]*"/g, '')) || /m{3,}/i.test(st.numFmt);
        return formatValue(v, hasDate && hasTime ? 'm/d/yyyy h:mm:ss AM/PM' : hasTime ? 'h:mm:ss AM/PM' : 'm/d/yyyy', this.date1904).text;
      }
      if (/%/.test(st.numFmt)) return trimNum(v * 100) + '%';
      return trimNum(v);
    }
    return (st.quotePrefix ? "'" : '') + v;
  }

  usedRange(si: number): Range {
    const sh = this.sheets[si];
    return { r1: 0, c1: 0, r2: Math.max(0, sh.maxRow - 1), c2: Math.max(0, sh.maxCol - 1) };
  }

  colWidthChars(si: number, c: number): number {
    const sh = this.sheets[si];
    const col = sh.cols.find(x => c >= x.min && c <= x.max);
    return col?.width ?? sh.defaultColWidth;
  }
  colHidden(si: number, c: number) { return !!this.sheets[si].cols.find(x => c >= x.min && c <= x.max)?.hidden; }
  rowHeightPt(si: number, r: number): number { const sh = this.sheets[si]; return sh.rows.get(r)?.ht ?? sh.defaultRowHt; }
  rowHidden(si: number, r: number) { return !!this.sheets[si].rows.get(r)?.hidden; }

  mergeAt(si: number, r: number, c: number): Range | undefined { return this.sheets[si].merges.find(g => inRange(g, r, c)); }

  // ---------------------------------------------------------------- editing

  private touchCell(si: number, r: number, c: number) {
    let s = this.edits.cells.get(si);
    if (!s) this.edits.cells.set(si, (s = new Set()));
    s.add(key(r, c));
  }

  private unshare(si: number, cell: Cell) {
    // editing the master of a shared formula: give every child its own formula first
    if (!cell.shared?.master) return;
    const sh = this.sheets[si];
    for (const [k, other] of sh.cells) {
      if (other !== cell && other.shared && !other.shared.master && other.shared.si === cell.shared.si) {
        other.shared = undefined;
        const a = unkey(k);
        this.touchCell(si, a.r, a.c);
      }
    }
    cell.shared = undefined;
  }

  /** set a cell from what the user typed (number, percent, date, boolean, formula or text) */
  setInput(si: number, r: number, c: number, input: string) {
    const sh = this.sheets[si];
    const k = key(r, c);
    const old = sh.cells.get(k);
    const s = old?.s ?? this.xfAt(si, r, c);
    if (old) this.unshare(si, old);
    if (old?.shared) old.shared = undefined;
    const cell: Cell = { v: null, s };
    if (input === '') {
      if (!old) return;
      if (old.s === 0) sh.cells.delete(k); else sh.cells.set(k, cell);
    } else if (input.startsWith('=') && input.length > 1) {
      cell.f = input.slice(1);
      try { parseFormula(cell.f); } catch { throw new Error('There\'s a problem with this formula.'); }
      sh.cells.set(k, cell);
    } else {
      const st = this.style(s);
      if (input.startsWith("'")) cell.v = input.slice(1);
      else if (st.numFmt === '@') cell.v = input;
      else {
        const parsed = parseInput(input);
        cell.v = parsed.v;
        if (parsed.fmt && st.numFmtId === 0) {
          const xf = { ...this.styles.allXfs()[s] };
          xf.numFmtId = this.styles.numFmtFor(parsed.fmt);
          cell.s = this.styles.xfFor(xf);
        }
      }
      sh.cells.set(k, cell);
    }
    if (input !== '') { sh.maxRow = Math.max(sh.maxRow, r + 1); sh.maxCol = Math.max(sh.maxCol, c + 1); }
    this.touchCell(si, r, c);
    this.recalc();
  }

  /** set a raw value/formula (paste, replace, fill) */
  setCellRaw(si: number, r: number, c: number, v: Scalar, f?: string, s?: number) {
    const sh = this.sheets[si];
    const k = key(r, c);
    const old = sh.cells.get(k);
    if (old) this.unshare(si, old);
    const cell: Cell = { v: f ? null : v, f, s: s ?? old?.s ?? this.xfAt(si, r, c) };
    if (cell.v === null && !cell.f && cell.s === 0) { if (!old) return; sh.cells.delete(k); }
    else sh.cells.set(k, cell);
    if (cell.v !== null || cell.f) { sh.maxRow = Math.max(sh.maxRow, r + 1); sh.maxCol = Math.max(sh.maxCol, c + 1); }
    this.touchCell(si, r, c);
    this.recalc();
  }

  /** apply a format change to every cell of a range */
  formatRange(si: number, g: Range, patch: { font?: Partial<Font>; fill?: string | null; numFmt?: string; align?: Partial<Align>; border?: 'all' | 'outside' | 'bottom' | 'top' | 'left' | 'right' | 'none' | 'thickBottom' }) {
    const sh = this.sheets[si];
    const r2 = Math.min(g.r2, Math.max(sh.maxRow - 1, g.r1, 200)), c2 = Math.min(g.c2, Math.max(sh.maxCol - 1, g.c1, 50));
    for (let r = g.r1; r <= r2; r++) for (let c = g.c1; c <= c2; c++) {
      const k = key(r, c);
      const cell = sh.cells.get(k);
      const cur = cell?.s ?? this.xfAt(si, r, c);
      const xf: Xf = { ...this.styles.allXfs()[cur] ?? this.styles.allXfs()[0] };
      if (patch.font) xf.fontId = this.styles.fontFor({ ...this.styles.allFonts()[xf.fontId], ...patch.font });
      if (patch.fill !== undefined) xf.fillId = patch.fill === null ? 0 : this.styles.fillFor({ pattern: 'solid', fg: patch.fill });
      if (patch.numFmt !== undefined) xf.numFmtId = this.styles.numFmtFor(patch.numFmt);
      if (patch.align) xf.align = { ...(xf.align ?? {}), ...patch.align };
      if (patch.border) {
        const b: Border = { ...(this.styles.allBorders()[xf.borderId] ?? {}) };
        const thin = { style: 'thin' };
        if (patch.border === 'none') { delete b.left; delete b.right; delete b.top; delete b.bottom; }
        if (patch.border === 'all') Object.assign(b, { left: thin, right: thin, top: thin, bottom: thin });
        if (patch.border === 'outside') { if (r === g.r1) b.top = thin; if (r === g.r2) b.bottom = thin; if (c === g.c1) b.left = thin; if (c === g.c2) b.right = thin; }
        if (patch.border === 'bottom' && r === g.r2) b.bottom = thin;
        if (patch.border === 'thickBottom' && r === g.r2) b.bottom = { style: 'medium' };
        if (patch.border === 'top' && r === g.r1) b.top = thin;
        if (patch.border === 'left' && c === g.c1) b.left = thin;
        if (patch.border === 'right' && c === g.c2) b.right = thin;
        xf.borderId = this.styles.borderFor(b);
      }
      const ns = this.styles.xfFor(xf);
      if (ns === cur && cell) continue;
      if (cell) { if (cell.s === ns) continue; cell.s = ns; } else { if (ns === 0) continue; sh.cells.set(k, { v: null, s: ns }); }
      this.touchCell(si, r, c);
    }
    this.recalc();
  }

  setColWidth(si: number, c1: number, c2: number, chars: number | null, hidden?: boolean) {
    const sh = this.sheets[si];
    for (let c = c1; c <= c2; c++) {
      // split any spanning <col> so this column has its own entry
      const i = sh.cols.findIndex(x => c >= x.min && c <= x.max);
      let col: ColInfo;
      if (i >= 0) {
        const x = sh.cols[i];
        const parts: ColInfo[] = [];
        if (x.min < c) parts.push({ ...x, max: c - 1 });
        col = { ...x, min: c, max: c };
        parts.push(col);
        if (x.max > c) parts.push({ ...x, min: c + 1 });
        sh.cols.splice(i, 1, ...parts);
      } else {
        col = { min: c, max: c, width: sh.defaultColWidth };
        sh.cols.push(col);
        sh.cols.sort((a, b) => a.min - b.min);
      }
      if (chars !== null) { col.width = chars; col.customWidth = true; }
      if (hidden !== undefined) col.hidden = hidden;
    }
    this.edits.cols.add(si);
  }

  setRowHeight(si: number, r1: number, r2: number, pt: number | null, hidden?: boolean) {
    const sh = this.sheets[si];
    for (let r = r1; r <= r2; r++) {
      const info = { ...(sh.rows.get(r) ?? {}) };
      if (pt !== null) { info.ht = pt; info.customHeight = true; }
      if (hidden !== undefined) info.hidden = hidden || undefined;
      sh.rows.set(r, info);
      let s = this.edits.rows.get(si);
      if (!s) this.edits.rows.set(si, (s = new Set()));
      s.add(r);
    }
  }

  merge(si: number, g: Range, center = true) {
    const sh = this.sheets[si];
    sh.merges = sh.merges.filter(m => !(m.r1 <= g.r2 && m.r2 >= g.r1 && m.c1 <= g.c2 && m.c2 >= g.c1));
    sh.merges.push(norm(g));
    // Excel keeps only the upper-left value
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
      if (r === g.r1 && c === g.c1) continue;
      const cell = sh.cells.get(key(r, c));
      if (cell && (cell.v !== null || cell.f)) this.setCellRaw(si, r, c, null);
    }
    if (center) this.formatRange(si, { r1: g.r1, c1: g.c1, r2: g.r1, c2: g.c1 }, { align: { h: 'center' } });
    this.edits.merges.add(si);
  }
  unmerge(si: number, g: Range) {
    const sh = this.sheets[si];
    sh.merges = sh.merges.filter(m => !(m.r1 <= g.r2 && m.r2 >= g.r1 && m.c1 <= g.c2 && m.c2 >= g.c1));
    this.edits.merges.add(si);
  }

  setFreeze(si: number, rows: number, cols: number) {
    const sh = this.sheets[si];
    sh.freeze = rows || cols ? { rows, cols } : null;
    this.edits.view.add(si);
  }

  renameSheet(si: number, name: string) {
    const old = this.sheets[si].name;
    if (old === name) return;
    this.sheets[si].name = name;
    // formulas that name the sheet follow it
    const q = (n: string) => (/^[A-Za-z_][\w.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);
    this.sheets.forEach((sh, i) => {
      for (const [k, cell] of sh.cells) {
        if (!cell.f) continue;
        const re = new RegExp(`(^|[^\\w'])(${escRe(q(old))}|'${escRe(old.replace(/'/g, "''"))}')!`, 'g');
        const next = cell.f.replace(re, (m, p: string) => `${p}${q(name)}!`);
        if (next !== cell.f) { cell.f = next; const a = unkey(k); this.unshare(i, cell); this.touchCell(i, a.r, a.c); }
      }
    });
    this.edits.workbook = true;
    this.recalc();
  }

  addSheet(at: number, name?: string): number {
    let n = this.sheets.length + 1;
    while (this.sheetIndex(`Sheet${n}`) >= 0) n++;
    const nm = name ?? `Sheet${n}`;
    let partNo = 1;
    const parts = new Set([...this.zip.names(), ...this.sheets.map(s => s.part)]);
    while (parts.has(`xl/worksheets/sheet${partNo}.xml`)) partNo++;
    const sheetId = Math.max(0, ...this.sheets.map(s => s.sheetId), ...this.removedSheets.map(s => s.sheetId)) + 1;
    const sh: Sheet = {
      name: nm, part: `xl/worksheets/sheet${partNo}.xml`, rId: '', sheetId, state: 'visible', cells: new Map(), cols: [], rows: new Map(), merges: [],
      freeze: null, topLeft: { r: 0, c: 0 }, active: { r: 0, c: 0 }, sqref: 'A1', defaultRowHt: 15, defaultColWidth: 9.14, showGrid: true, zoom: 100,
      maxRow: 0, maxCol: 0, doc: null, isNew: true, protection: null,
    };
    const order = this.sheets.map((_, i) => i);
    this.sheets.splice(at, 0, sh);
    order.splice(at, 0, -1);
    this.remapEdits(order);
    this.edits.workbook = true;
    return at;
  }

  /** Move or Copy: move sheet `from` before the sheet now at `before` (sheets.length = move to end); returns its new index */
  moveSheet(from: number, before: number): number {
    if (before === from || before === from + 1) return from;
    const order = this.sheets.map((_, i) => i);
    const active = this.sheets[this.activeTab];
    const [sh] = this.sheets.splice(from, 1);
    const [o] = order.splice(from, 1);
    const at = before > from ? before - 1 : before;
    this.sheets.splice(at, 0, sh);
    order.splice(at, 0, o);
    this.remapEdits(order);
    this.activeTab = Math.max(0, this.sheets.indexOf(active));
    this.reordered = true;
    this.edits.workbook = true;
    return at;
  }

  /** Move or Copy with Create a copy: a new sheet with the same cells, formats, layout and markup, named like
   *  Excel names copies ("Q3 grants (2)"), inserted before the sheet now at `before`; returns its index */
  copySheet(from: number, before: number, name?: string): number {
    const src = this.sheets[from];
    const nm = name ?? this.copyName(src.name);
    let partNo = 1;
    const parts = new Set([...this.zip.names(), ...this.sheets.map(s => s.part), ...this.removedSheets.map(s => s.part)]);
    while (parts.has(`xl/worksheets/sheet${partNo}.xml`)) partNo++;
    const part = `xl/worksheets/sheet${partNo}.xml`;
    const sheetId = Math.max(0, ...this.sheets.map(s => s.sheetId), ...this.removedSheets.map(s => s.sheetId)) + 1;
    let doc: XDocument | null = null;
    let copyRels: string | undefined;
    if (src.doc) { const c = this.copiedSheetXml(src); doc = parseXml(c.xml); copyRels = c.rels; }
    else if (src.copyRels !== undefined) copyRels = src.copyRels;
    const cells = new Map<number, Cell>();
    for (const [k, cell] of src.cells) cells.set(k, { ...cell, shared: cell.shared ? { ...cell.shared } : undefined });
    const sh: Sheet = {
      ...src, name: nm, part, rId: '', sheetId, state: 'visible', cells, cols: src.cols.map(c => ({ ...c, raw: c.raw ? { ...c.raw } : undefined })),
      rows: new Map([...src.rows].map(([r, info]) => [r, { ...info }])), merges: src.merges.map(g => ({ ...g })),
      freeze: src.freeze ? { ...src.freeze } : null, topLeft: { ...src.topLeft }, active: { ...src.active },
      doc, isNew: true, protection: src.protection ? { ...src.protection } : null, copyRels,
    };
    const order = this.sheets.map((_, i) => i);
    this.sheets.splice(before, 0, sh);
    order.splice(before, 0, -1);
    this.remapEdits(order);
    // the copy carries the source's unsaved edits (they are relative to the same markup)
    const e = this.edits, si = before, srcNow = this.sheets.indexOf(src);
    const cs = e.cells.get(srcNow); if (cs?.size) e.cells.set(si, new Set(cs));
    const rs = e.rows.get(srcNow); if (rs?.size) e.rows.set(si, new Set(rs));
    if (e.cols.has(srcNow)) e.cols.add(si);
    if (e.merges.has(srcNow)) e.merges.add(si);
    if (e.view.has(srcNow)) e.view.add(si);
    if (e.protection.has(srcNow)) e.protection.add(si);
    this.reordered = true;
    this.edits.workbook = true;
    this.recalc();
    return si;
  }

  /** "Q3 grants" -> "Q3 grants (2)" (then (3), ...), kept within Excel's 31 characters */
  copyName(base: string): string {
    const stem = base.replace(/ \(\d+\)$/, '');
    for (let n = 2; ; n++) {
      const suffix = ` (${n})`;
      const nm = stem.slice(0, 31 - suffix.length) + suffix;
      if (this.sheetIndex(nm) < 0) return nm;
    }
  }

  /** Protect Sheet / Unprotect Sheet */
  setProtection(si: number, p: SheetProtection | null) {
    this.sheets[si].protection = p ? { ...p } : null;
    this.edits.protection.add(si);
  }

  /** a sheet's XML for a copy: tabSelected dropped (the copy is not grouped with the source), and markup that
   *  points at parts the copy cannot share (drawings, comments, tables, controls, printer settings) left out;
   *  external hyperlinks keep their relationships */
  private copiedSheetXml(src: Sheet): { xml: string; rels?: string } {
    const doc = src.doc!;
    const ws = doc.root;
    const splices: Splice[] = [];
    for (const v of kids(kid(ws, 'sheetViews'), 'sheetView')) if (v.attrs.tabSelected !== undefined) {
      splices.push({ start: v.start, end: v.openEnd, text: `${retag(doc, v, { tabSelected: null })}${v.selfClosing ? '/>' : '>'}` });
    }
    const i = src.part.lastIndexOf('/');
    const relsName = `${src.part.slice(0, i + 1)}_rels/${src.part.slice(i + 1)}.rels`;
    const relsDoc = this.xml(relsName);
    const external = new Set<string>();
    let rels: string | undefined;
    if (relsDoc) {
      const keep = kids(relsDoc.root, 'Relationship').filter(r => r.attrs.TargetMode === 'External');
      for (const r of keep) external.add(r.attrs.Id);
      rels = applySplices(relsDoc.text, [{ start: relsDoc.root.openEnd, end: relsDoc.root.closeStart, text: keep.map(r => relsDoc.text.slice(r.start, r.end)).join('') }]);
      if (relsDoc.root.selfClosing) rels = relsDoc.text;
    }
    const rid = (n: XNode) => Object.entries(n.attrs).find(([k]) => k === 'r:id' || k.endsWith(':id'))?.[1];
    for (const n of ws.children) {
      if (['drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'tableParts'].includes(n.local)) splices.push({ start: n.start, end: n.end, text: '' });
      else if (n.local === 'hyperlinks') {
        for (const h of kids(n, 'hyperlink')) { const id = rid(h); if (id && !external.has(id)) splices.push({ start: h.start, end: h.end, text: '' }); }
      } else if (n.local === 'pageSetup' && rid(n)) {
        const k = Object.keys(n.attrs).find(a => a === 'r:id' || a.endsWith(':id'))!;
        splices.push({ start: n.start, end: n.openEnd, text: `${retag(doc, n, { [k]: null })}${n.selfClosing ? '/>' : '>'}` });
      }
    }
    return { xml: applySplices(doc.text, splices), rels: external.size ? rels : undefined };
  }

  /** edits are kept per sheet index: follow a change of sheet order (order[newIndex] = old index, -1 for new) */
  private remapEdits(order: number[]) {
    const e = this.edits;
    const mapM = <T>(m: Map<number, T>) => { const out = new Map<number, T>(); order.forEach((o, n) => { if (o >= 0 && m.has(o)) out.set(n, m.get(o)!); }); return out; };
    const mapS = (s: Set<number>) => { const out = new Set<number>(); order.forEach((o, n) => { if (o >= 0 && s.has(o)) out.add(n); }); return out; };
    e.cells = mapM(e.cells); e.rows = mapM(e.rows);
    e.cols = mapS(e.cols); e.merges = mapS(e.merges); e.view = mapS(e.view); e.protection = mapS(e.protection);
  }

  deleteSheet(si: number) {
    const order = this.sheets.map((_, i) => i);
    order.splice(si, 1);
    this.remapEdits(order);
    const [sh] = this.sheets.splice(si, 1);
    if (!sh.isNew) this.removedSheets.push(sh);
    if (this.activeTab >= this.sheets.length) this.activeTab = this.sheets.length - 1;
    this.edits.workbook = true;
    this.recalc();
  }

  setActiveTab(i: number) { this.activeTab = i; }

  isDirty(): boolean {
    const e = this.edits;
    return [...e.cells.values()].some(s => s.size > 0) || e.cols.size > 0 || [...e.rows.values()].some(s => s.size > 0) || e.merges.size > 0 || e.view.size > 0 || e.protection.size > 0 || e.workbook || this.styles.dirty();
  }

  // ---------------------------------------------------------------- saving

  save(): Uint8Array {
    if (!this.isDirty()) return this.zip.save();
    // formula cells that depend on edited cells get fresh cached values, as Excel writes them
    const affected = this.affectedFormulas();
    const sheetNames = new Set(this.sheets.map(s => s.part));
    this.sheets.forEach((sh, si) => {
      const cellKeys = new Set([...(this.edits.cells.get(si) ?? []), ...(affected.get(si) ?? [])]);
      const rows = this.edits.rows.get(si) ?? new Set<number>();
      const any = cellKeys.size || rows.size || this.edits.cols.has(si) || this.edits.merges.has(si) || this.edits.view.has(si) || this.edits.protection.has(si);
      if (sh.isNew) {
        this.zip.write(sh.part, sh.doc ? this.rewriteSheet(si, cellKeys, rows, affected.get(si) ?? new Set()) : this.newSheetXml(si));
        if (sh.copyRels) { const i = sh.part.lastIndexOf('/'); this.zip.write(`${sh.part.slice(0, i + 1)}_rels/${sh.part.slice(i + 1)}.rels`, sh.copyRels); }
        return;
      }
      if (!any || !sh.doc) return;
      this.zip.write(sh.part, this.rewriteSheet(si, cellKeys, rows, affected.get(si) ?? new Set()));
    });
    if (this.styles.dirty()) this.writeStyles();
    if (this.edits.workbook) this.writeWorkbook();
    for (const sh of this.removedSheets) if (!sheetNames.has(sh.part)) this.zip.remove(sh.part);
    // a calcChain naming cells whose formulas were removed makes Excel repair the file: drop it
    if (this.formulaRemoved) this.dropCalcChain();
    return this.zip.save();
  }

  private formulaRemoved = false;

  private affectedFormulas(): Map<number, Set<number>> {
    const out = new Map<number, Set<number>>();
    const dirty: { si: number; r: number; c: number }[] = [];
    this.edits.cells.forEach((set, si) => set.forEach(k => { const a = unkey(k); dirty.push({ si, ...a }); }));
    if (!dirty.length) return out;
    const formulas: { si: number; k: number; refs: { si: number; g: Range }[] }[] = [];
    this.sheets.forEach((sh, si) => {
      for (const [k, cell] of sh.cells) {
        if (!cell.f) continue;
        let ast: Node | null = null;
        try { ast = parseFormula(cell.f); } catch { continue; }
        const refs = formulaRefs(ast, sh.name).map(x => ({ si: this.sheetIndex(x.sheet), g: parseRange(x.b ? `${x.a}:${x.b}` : x.a) })).filter(x => x.si >= 0 && x.g) as { si: number; g: Range }[];
        formulas.push({ si, k, refs });
      }
    });
    const hit = (si: number, r: number, c: number) => formulas.filter(f => f.refs.some(x => x.si === si && inRange(x.g, r, c)));
    const queue = [...dirty];
    const seen = new Set<string>();
    while (queue.length) {
      const d = queue.pop()!;
      for (const f of hit(d.si, d.r, d.c)) {
        const id = `${f.si}:${f.k}`;
        if (seen.has(id)) continue;
        seen.add(id);
        let s = out.get(f.si);
        if (!s) out.set(f.si, (s = new Set()));
        s.add(f.k);
        queue.push({ si: f.si, ...unkey(f.k) });
      }
    }
    return out;
  }

  private cellXml(si: number, r: number, c: number, cell: Cell | undefined, orig: XNode | undefined, doc: XDocument): string {
    const ref = addrName(r, c);
    if (!cell) return '';
    // keep attributes we do not manage (cm, vm, ph) from the original element
    const extra = orig ? Object.entries(orig.attrs).filter(([k2]) => !['r', 's', 't'].includes(k2)).map(([k2, v]) => ` ${k2}="${escAttr(v)}"`).join('') : '';
    const sAttr = cell.s ? ` s="${cell.s}"` : '';
    if (cell.f) {
      const v = this.value(si, r, c);
      let t = '', vx = '';
      if (typeof v === 'string') { t = ' t="str"'; vx = `<v>${escText(v)}</v>`; }
      else if (typeof v === 'boolean') { t = ' t="b"'; vx = `<v>${v ? 1 : 0}</v>`; }
      else if (isErr(v)) { t = ' t="e"'; vx = `<v>${escText(v.error)}</v>`; }
      else if (typeof v === 'number') vx = `<v>${numText(v)}</v>`;
      let fx: string;
      if (cell.shared && orig) {
        // untouched shared formula (only its cached value changes): keep the original <f> element
        const fEl = kid(orig, 'f');
        fx = fEl ? doc.text.slice(fEl.start, fEl.end) : `<f>${escText(cell.f)}</f>`;
      } else if (cell.array) fx = `<f t="array" ref="${cell.array}">${escText(cell.f)}</f>`;
      else fx = `<f>${escText(cell.f)}</f>`;
      return `<c r="${ref}"${sAttr}${t}${extra}>${fx}${vx}</c>`;
    }
    const v = cell.v;
    if (v === null) return `<c r="${ref}"${sAttr}${extra}/>`;
    if (typeof v === 'string') return `<c r="${ref}"${sAttr} t="inlineStr"${extra}><is><t${/^\s|\s$|\n/.test(v) ? ' xml:space="preserve"' : ''}>${escText(v)}</t></is></c>`;
    if (typeof v === 'boolean') return `<c r="${ref}"${sAttr} t="b"${extra}><v>${v ? 1 : 0}</v></c>`;
    if (isErr(v)) return `<c r="${ref}"${sAttr} t="e"${extra}><v>${escText(v.error)}</v></c>`;
    return `<c r="${ref}"${sAttr}${extra}><v>${numText(v)}</v></c>`;
  }

  private rewriteSheet(si: number, cellKeys: Set<number>, rowEdits: Set<number>, affected: Set<number>): string {
    const sh = this.sheets[si];
    const doc = sh.doc!;
    const ws = doc.root;
    const splices: Splice[] = [];
    const sheetData = kid(ws, 'sheetData');
    const rowNodes = new Map<number, XNode>();
    const cellNodes = new Map<number, XNode>();
    let ri = -1;
    for (const row of kids(sheetData, 'row')) {
      ri = row.attrs.r ? Number(row.attrs.r) - 1 : ri + 1;
      rowNodes.set(ri, row);
      let ci = -1;
      for (const c of kids(row, 'c')) {
        const a = c.attrs.r ? parseAddr(c.attrs.r) : null;
        ci = a ? a.c : ci + 1;
        cellNodes.set(key(a ? a.r : ri, ci), c);
      }
    }
    // group edited cells by row
    const byRow = new Map<number, number[]>();
    for (const k of cellKeys) { const { r } = unkey(k); (byRow.get(r) ?? byRow.set(r, []).get(r)!).push(k); }
    for (const r of rowEdits) if (!byRow.has(r)) byRow.set(r, []);
    const newRows: { r: number; xml: string }[] = [];
    for (const [r, ks] of byRow) {
      ks.sort((a, b) => a - b);
      const rowNode = rowNodes.get(r);
      const info = sh.rows.get(r);
      if (rowNode) {
        // row attributes (height/hidden) when edited
        let rowTag: string | null = null;
        const rowChanges: Record<string, string | null> = {};
        if (rowEdits.has(r)) {
          rowChanges.ht = info?.ht !== undefined ? String(info.ht) : null;
          rowChanges.customHeight = info?.customHeight ? '1' : null;
          rowChanges.hidden = info?.hidden ? '1' : null;
        }
        // spans hint must cover new cells
        if (rowNode.attrs.spans && ks.length) {
          const [a, b] = rowNode.attrs.spans.split(':').map(Number);
          const cs = ks.map(k => unkey(k).c + 1);
          const lo = Math.min(a, ...cs), hi = Math.max(b, ...cs);
          if (lo !== a || hi !== b) rowChanges.spans = `${lo}:${hi}`;
        }
        if (Object.keys(rowChanges).length) rowTag = retag(doc, rowNode, rowChanges);
        const inserts: string[] = [];
        for (const k of ks) {
          const { c } = unkey(k);
          const node = cellNodes.get(k);
          const cell = sh.cells.get(k);
          if (node && kid(node, 'f') && !cell?.f) this.formulaRemoved = true;
          if (node) {
            if (affected.has(k) && !cellKeys.has(k)) {
              // dependent formula: rewrite only its cached value
              splices.push({ start: node.start, end: node.end, text: this.cellXml(si, r, c, cell, node, doc) });
            } else splices.push({ start: node.start, end: node.end, text: this.cellXml(si, r, c, cell, node, doc) });
          } else if (cell) {
            // insert before the first existing cell to the right, else before </row>
            const after = kids(rowNode, 'c').find(n => { const a = n.attrs.r ? parseAddr(n.attrs.r) : null; return a ? a.c > c : false; });
            const xml = this.cellXml(si, r, c, cell, undefined, doc);
            if (after) splices.push({ start: after.start, end: after.start, text: xml });
            else inserts.push(xml);
          }
        }
        if (rowNode.selfClosing) {
          const tag = rowTag ?? doc.text.slice(rowNode.start, rowNode.openEnd - 2).replace(/\s+$/, '');
          if (inserts.length) splices.push({ start: rowNode.start, end: rowNode.end, text: `${tag}>${inserts.join('')}</row>` });
          else if (rowTag) splices.push({ start: rowNode.start, end: rowNode.openEnd, text: `${rowTag}/>` });
        } else {
          if (rowTag) splices.push({ start: rowNode.start, end: rowNode.openEnd, text: `${rowTag}>` });
          if (inserts.length) splices.push({ start: rowNode.closeStart, end: rowNode.closeStart, text: inserts.join('') });
        }
      } else {
        const cellsXml = ks.map(k => { const { c } = unkey(k); return this.cellXml(si, r, c, sh.cells.get(k), undefined, doc); }).join('');
        if (!cellsXml && !rowEdits.has(r)) continue;
        const attrs = [`r="${r + 1}"`];
        if (info?.ht !== undefined) attrs.push(`ht="${info.ht}"`);
        if (info?.customHeight) attrs.push('customHeight="1"');
        if (info?.hidden) attrs.push('hidden="1"');
        newRows.push({ r, xml: cellsXml ? `<row ${attrs.join(' ')}>${cellsXml}</row>` : `<row ${attrs.join(' ')}/>` });
      }
    }
    if (newRows.length && sheetData) {
      newRows.sort((a, b) => a.r - b.r);
      if (sheetData.selfClosing) {
        splices.push({ start: sheetData.start, end: sheetData.end, text: `<sheetData>${newRows.map(n => n.xml).join('')}</sheetData>` });
      } else {
        const existing = [...rowNodes.entries()].sort((a, b) => a[0] - b[0]);
        for (const nr of newRows) {
          const after = existing.find(([r]) => r > nr.r);
          const at = after ? after[1].start : sheetData.closeStart;
          splices.push({ start: at, end: at, text: nr.xml });
        }
      }
    }
    // dimension
    const dim = kid(ws, 'dimension');
    if (dim && (cellKeys.size)) {
      const cur = parseRange(dim.attrs.ref ?? 'A1') ?? { r1: 0, c1: 0, r2: 0, c2: 0 };
      const want: Range = { ...cur };
      for (const k of cellKeys) { const a = unkey(k); if (sh.cells.has(k)) { want.r1 = Math.min(want.r1, a.r); want.c1 = Math.min(want.c1, a.c); want.r2 = Math.max(want.r2, a.r); want.c2 = Math.max(want.c2, a.c); } }
      if (rangeName(want) !== rangeName(cur)) splices.push({ start: dim.start, end: dim.end, text: `${retag(doc, dim, { ref: rangeName(want) })}/>` });
    }
    if (this.edits.cols.has(si)) {
      const colsNode = kid(ws, 'cols');
      const xml = sh.cols.length ? `<cols>${sh.cols.map(colXml).join('')}</cols>` : '';
      if (colsNode) splices.push({ start: colsNode.start, end: colsNode.end, text: xml });
      else if (xml && sheetData) splices.push({ start: sheetData.start, end: sheetData.start, text: xml });
    }
    if (this.edits.protection.has(si)) {
      const node = kid(ws, 'sheetProtection');
      const xml = sh.protection ? protectionXml(sh.protection) : '';
      if (node) splices.push({ start: node.start, end: node.end, text: xml });
      else if (xml) {
        // CT_Worksheet order: sheetData, sheetCalcPr, sheetProtection
        const after = kid(ws, 'sheetCalcPr') ?? sheetData;
        const at = after ? after.end : ws.closeStart;
        splices.push({ start: at, end: at, text: xml });
      }
    }
    if (this.edits.merges.has(si)) {
      const mc = kid(ws, 'mergeCells');
      const xml = sh.merges.length ? `<mergeCells count="${sh.merges.length}">${sh.merges.map(g => `<mergeCell ref="${rangeName(g)}"/>`).join('')}</mergeCells>` : '';
      if (mc) splices.push({ start: mc.start, end: mc.end, text: xml });
      else if (xml) {
        const order = ['sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews'];
        const last = [...ws.children].reverse().find(n => order.includes(n.local));
        const at = last ? last.end : ws.closeStart;
        splices.push({ start: at, end: at, text: xml });
      }
    }
    if (this.edits.view.has(si)) {
      const views = kid(ws, 'sheetViews');
      const view = kid(views, 'sheetView');
      const fz = sh.freeze;
      let inner = '';
      const sel = `<selection${fz ? ` pane="${fz.rows && fz.cols ? 'bottomRight' : fz.rows ? 'bottomLeft' : 'topRight'}"` : ''} activeCell="${addrName(sh.active.r, sh.active.c)}" sqref="${addrName(sh.active.r, sh.active.c)}"/>`;
      if (fz) {
        const pane = fz.rows && fz.cols ? 'bottomRight' : fz.rows ? 'bottomLeft' : 'topRight';
        inner = `<pane${fz.cols ? ` xSplit="${fz.cols}"` : ''}${fz.rows ? ` ySplit="${fz.rows}"` : ''} topLeftCell="${addrName(fz.rows, fz.cols)}" activePane="${pane}" state="frozen"/>${sel}`;
      } else inner = sel;
      if (view) {
        const tag = doc.text.slice(view.start, view.selfClosing ? view.openEnd - 2 : view.openEnd - 1).replace(/\s+$/, '');
        splices.push({ start: view.start, end: view.end, text: `${tag}>${inner}</sheetView>` });
      } else if (views) {
        splices.push({ start: views.start, end: views.end, text: `<sheetViews><sheetView workbookViewId="0">${inner}</sheetView></sheetViews>` });
      }
    }
    return applySplices(doc.text, splices);
  }

  private newSheetXml(si: number): string {
    const sh = this.sheets[si];
    const fake = parseXml(`<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`);
    sh.doc = fake;
    const keys = new Set(sh.cells.keys());
    const rows = new Set(sh.rows.keys());
    if (sh.cols.length) this.edits.cols.add(si);
    if (sh.merges.length) this.edits.merges.add(si);
    if (sh.freeze) this.edits.view.add(si);
    if (sh.protection) this.edits.protection.add(si);
    const out = this.rewriteSheet(si, keys, rows, new Set());
    sh.doc = null;
    return out;
  }

  private writeStyles() {
    const part = this.stylesPart;
    const doc = this.styles.doc;
    const a = this.styles.added;
    if (!doc) {
      const xml = `<styleSheet xmlns="${MAIN_NS}">${a.numFmts.length ? `<numFmts count="${a.numFmts.length}">${a.numFmts.map(n => `<numFmt numFmtId="${n.id}" formatCode="${escAttr(n.code)}"/>`).join('')}</numFmts>` : ''}<fonts count="${this.styles.allFonts().length}">${this.styles.allFonts().map(fontXml).join('')}</fonts><fills count="${this.styles.allFills().length}">${this.styles.allFills().map(fillXml).join('')}</fills><borders count="${this.styles.allBorders().length}">${this.styles.allBorders().map(borderXml).join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${this.styles.allXfs().length}">${this.styles.allXfs().map(xfXml).join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
      this.zip.write(part, xml);
      return;
    }
    const root = doc.root;
    const splices: Splice[] = [];
    const append = (local: string, items: string[], total: number, before?: string) => {
      if (!items.length) return;
      const node = kid(root, local);
      if (!node) {
        const anchor = before ? kid(root, before) : undefined;
        const at = anchor ? anchor.start : root.openEnd;
        splices.push({ start: at, end: at, text: `<${local} count="${total}">${items.join('')}</${local}>` });
        return;
      }
      if (node.selfClosing) {
        splices.push({ start: node.start, end: node.end, text: `${retag(doc, node, { count: String(total) })}>${items.join('')}</${node.name}>` });
      } else {
        splices.push({ start: node.start, end: node.openEnd, text: `${retag(doc, node, { count: String(total) })}>` });
        splices.push({ start: node.closeStart, end: node.closeStart, text: items.join('') });
      }
    };
    const numFmtsNode = kid(root, 'numFmts');
    append('numFmts', a.numFmts.map(n => `<numFmt numFmtId="${n.id}" formatCode="${escAttr(n.code)}"/>`), (numFmtsNode ? kids(numFmtsNode, 'numFmt').length : 0) + a.numFmts.length, 'fonts');
    append('fonts', a.fonts.map(fontXml), this.styles.allFonts().length);
    append('fills', a.fills.map(fillXml), this.styles.allFills().length);
    append('borders', a.borders.map(borderXml), this.styles.allBorders().length);
    append('cellXfs', a.xfs.map(xfXml), this.styles.allXfs().length);
    this.zip.write(part, applySplices(doc.text, splices));
  }

  private writeWorkbook() {
    const doc = this.wbDoc;
    const wb = doc.root;
    const sheetsNode = kid(wb, 'sheets')!;
    const relsName = this.wbPart.replace(/([^/]+)$/, '_rels/$1.rels');
    const relsDoc = this.xml(relsName)!;
    const relSplices: Splice[] = [];
    const ctDoc = this.xml('[Content_Types].xml')!;
    const ctSplices: Splice[] = [];
    // new sheets: relationship + content type
    let maxRel = 0;
    for (const r of kids(relsDoc.root, 'Relationship')) { const m = /^rId(\d+)$/.exec(r.attrs.Id ?? ''); if (m) maxRel = Math.max(maxRel, Number(m[1])); }
    const base = this.wbPart.slice(0, this.wbPart.lastIndexOf('/') + 1);
    for (const sh of this.sheets) {
      if (!sh.isNew || sh.rId) continue;
      sh.rId = `rId${++maxRel}`;
      const target = sh.part.startsWith(base) ? sh.part.slice(base.length) : '/' + sh.part;
      relSplices.push({ start: relsDoc.root.closeStart, end: relsDoc.root.closeStart, text: `<Relationship Id="${sh.rId}" Type="${WS_TYPE}" Target="${target}"/>` });
      ctSplices.push({ start: ctDoc.root.closeStart, end: ctDoc.root.closeStart, text: `<Override PartName="/${sh.part}" ContentType="${WS_CT}"/>` });
    }
    for (const sh of this.removedSheets) {
      const rel = kids(relsDoc.root, 'Relationship').find(r => r.attrs.Id === sh.rId);
      if (rel) relSplices.push({ start: rel.start, end: rel.end, text: '' });
      const ov = kids(ctDoc.root, 'Override').find(o => o.attrs.PartName === '/' + sh.part);
      if (ov) ctSplices.push({ start: ov.start, end: ov.end, text: '' });
    }
    if (relSplices.length) this.zip.write(relsName, applySplices(relsDoc.text, relSplices));
    if (ctSplices.length) this.zip.write('[Content_Types].xml', applySplices(ctDoc.text, ctSplices));
    // <sheets>: keep each original <sheet> element (renamed in place), add new ones in order
    const orig = new Map(kids(sheetsNode, 'sheet').map(s => [s.attrs['r:id'] ?? Object.entries(s.attrs).find(([k]) => k.endsWith(':id'))?.[1], s]));
    const items = this.sheets.map(sh => {
      const node = orig.get(sh.rId);
      if (node && !sh.isNew) {
        if (node.attrs.name === sh.name) return doc.text.slice(node.start, node.end);
        return `${retag(doc, node, { name: sh.name })}${node.selfClosing ? '/>' : '>' + doc.text.slice(node.openEnd, node.end)}`;
      }
      return `<sheet name="${escAttr(sh.name)}" sheetId="${sh.sheetId}" r:id="${sh.rId}"${/xmlns:r=/.test(doc.text.slice(wb.start, wb.openEnd)) ? '' : ` xmlns:r="${REL_NS}"`}/>`;
    });
    const splices: Splice[] = [{ start: sheetsNode.openEnd, end: sheetsNode.closeStart, text: items.join('') }];
    const view = kid(kid(wb, 'bookViews'), 'workbookView');
    if (this.reordered) {
      // sheet-scoped names (print areas, filter databases) index the <sheets> list: they follow their sheet
      const origIdx = new Map(kids(sheetsNode, 'sheet').map((s, i) => [s.attrs['r:id'] ?? Object.entries(s.attrs).find(([k]) => k.endsWith(':id'))?.[1], i]));
      for (const dn of kids(kid(wb, 'definedNames'), 'definedName')) {
        if (dn.attrs.localSheetId === undefined) continue;
        const sh = this.sheets.find(x => !x.isNew && origIdx.get(x.rId) === Number(dn.attrs.localSheetId));
        const ni = sh ? this.sheets.indexOf(sh) : -1;
        if (ni >= 0 && ni !== Number(dn.attrs.localSheetId)) splices.push({ start: dn.start, end: dn.openEnd, text: `${retag(doc, dn, { localSheetId: String(ni) })}>` });
      }
      // the sheet saved as selected (tabSelected) stays the active one, so the file opens ungrouped
      const ai = this.diskActive ? this.sheets.indexOf(this.diskActive) : -1;
      const want = ai >= 0 ? ai : 0;
      if (view && attrNum(view, 'activeTab', 0) !== want) splices.push({ start: view.start, end: view.selfClosing ? view.end : view.openEnd, text: `${retag(doc, view, { activeTab: want ? String(want) : null })}${view.selfClosing ? '/>' : '>'}` });
    } else if (view && attrNum(view, 'activeTab', 0) >= this.sheets.length) splices.push({ start: view.start, end: view.selfClosing ? view.end : view.openEnd, text: `${retag(doc, view, { activeTab: '0' })}${view.selfClosing ? '/>' : '>'}` });
    this.zip.write(this.wbPart, applySplices(doc.text, splices));
  }

  private dropCalcChain() {
    const relsName = this.wbPart.replace(/([^/]+)$/, '_rels/$1.rels');
    const relsDoc = this.xml(relsName);
    if (!relsDoc) return;
    const rel = kids(relsDoc.root, 'Relationship').find(r => /\/calcChain$/.test(r.attrs.Type ?? ''));
    if (!rel) return;
    const base = this.wbPart.slice(0, this.wbPart.lastIndexOf('/') + 1);
    const part = normPath(base + (rel.attrs.Target ?? ''));
    this.zip.write(relsName, applySplices(relsDoc.text, [{ start: rel.start, end: rel.end, text: '' }]));
    const ct = this.xml('[Content_Types].xml')!;
    const ov = kids(ct.root, 'Override').find(o => o.attrs.PartName === '/' + part);
    if (ov) this.zip.write('[Content_Types].xml', applySplices(ct.text, [{ start: ov.start, end: ov.end, text: '' }]));
    this.zip.remove(part);
  }
}

// ------------------------------------------------------------------ helpers

function siText(doc: XDocument, si: XNode): string {
  const t = kid(si, 't');
  if (t && si.children.length === 1) return textOf(doc, t);
  return kids(si, 'r').map(r => { const x = kid(r, 't'); return x ? textOf(doc, x) : ''; }).join('') || (t ? textOf(doc, t) : '');
}

function colXml(c: ColInfo): string {
  const attrs = [`min="${c.min + 1}"`, `max="${c.max + 1}"`];
  if (c.width !== undefined) attrs.push(`width="${c.width}"`);
  if (c.style !== undefined) attrs.push(`style="${c.style}"`);
  if (c.hidden) attrs.push('hidden="1"');
  if (c.customWidth) attrs.push('customWidth="1"');
  // keep other original attributes (bestFit, outlineLevel, collapsed)
  for (const [k, v] of Object.entries(c.raw ?? {})) if (!['min', 'max', 'width', 'style', 'hidden', 'customWidth'].includes(k)) attrs.push(`${k}="${escAttr(v)}"`);
  return `<col ${attrs.join(' ')}/>`;
}

function numText(v: number): string {
  if (Number.isInteger(v)) return String(v);
  const s = String(v);
  return /e/.test(s) ? v.toPrecision(17).replace(/0+e/, 'e') : s;
}
function trimNum(v: number): string {
  const s = String(Number(v.toPrecision(15)));
  return s;
}
function escRe(s: string) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function normPath(p: string): string {
  const out: string[] = [];
  for (const seg of p.split('/')) { if (seg === '..') out.pop(); else if (seg !== '.' && seg !== '') out.push(seg); }
  return out.join('/');
}

/** how Excel reads typed input: numbers, currency, percent, dates, booleans, else text */
export function parseInput(s: string): { v: Scalar; fmt?: string } {
  const t = s.trim();
  if (/^(true|false)$/i.test(t)) return { v: t.toUpperCase() === 'TRUE' };
  if (/^#(N\/A|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|NULL!)$/i.test(t)) return { v: { error: t.toUpperCase() } };
  let m = /^(-?)\$\s?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$/.exec(t);
  if (m) { const n = Number((m[1] + m[2].replace(/,/g, '') + (m[4] ?? ''))); return { v: n, fmt: m[4] ? '"$"#,##0.00' : '"$"#,##0' }; }
  m = /^(-?)(\d{1,3}(,\d{3})+)(\.\d+)?$/.exec(t);
  if (m) return { v: Number(m[1] + m[2].replace(/,/g, '') + (m[4] ?? '')), fmt: m[4] ? '#,##0.00' : '#,##0' };
  m = /^(-?\d*\.?\d+)\s?%$/.exec(t);
  if (m) return { v: Number(m[1]) / 100, fmt: /\./.test(m[1]) ? '0.00%' : '0%' };
  if (/^-?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t)) return { v: Number(t) };
  if (/^\(\d+(\.\d+)?\)$/.test(t)) return { v: -Number(t.slice(1, -1)) };
  const d = parseDateText(t);
  if (d !== null) return { v: d, fmt: 'm/d/yyyy' };
  m = /^(\d{1,2})-([A-Za-z]{3})(?:-(\d{2,4}))?$/.exec(t);
  if (m) {
    const mi = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'].indexOf(m[2].toLowerCase());
    if (mi >= 0) { let y = m[3] ? Number(m[3]) : 2026; if (y < 100) y += 2000; return { v: partsToSerial(y, mi + 1, Number(m[1])), fmt: m[3] ? 'd-mmm-yy' : 'd-mmm' }; }
  }
  m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i.exec(t);
  if (m) { let h = Number(m[1]); if (m[4]) { h = h % 12 + (m[4].toUpperCase() === 'PM' ? 12 : 0); } return { v: (h * 3600 + Number(m[2]) * 60 + Number(m[3] ?? 0)) / 86400, fmt: m[4] ? 'h:mm AM/PM' : 'h:mm' }; }
  return { v: s };
}

export { colName, addrName, parseAddr, parseRange, rangeName, key, unkey };
export type { Range, Addr };
