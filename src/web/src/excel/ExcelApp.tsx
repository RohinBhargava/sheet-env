// Spreadsheet workbook view: workbook window with ribbon, name box and formula bar, grid,
// sheet tabs, status bar, Protected View, Find and Replace, Go To, copy/paste through the shared
// clipboard, editing with recalculated formulas and lossless Save (.xlsx via the files API).
import { useCallback, useEffect, useLayoutEffect, useMemo, useReducer, useRef, useState } from 'react';
import type { AppProps, HostHandle } from '../shell/types';
import { Workbook, parseRange, parseAddr, addrName, colName, isErr, formatValue, shiftFormula, FUNCTION_NAMES, MAX_ROW, MAX_COL, PROTECTION_FLAGS, protectionAllows, hasPassword, buildProtection, verifyPassword, type Range, type Addr, type Cell, type Scalar, type Sheet } from '../engine/index.ts';
import { Grid, type Sel, type EditorState, type ScrollThumbs } from './Grid';
import { Ribbon, TABS, type TabName, type RibbonState } from './Ribbon';
import { FindReplace, GoTo, NumberPrompt, ZoomDialog, FormatCells, InsertFunction, MoveCopyDialog, ProtectSheetDialog, UnprotectSheetDialog, MacroDialog, NUMBER_CATEGORIES, type FindOptions, type FindHit } from './dialogs';
import { Icon, ExcelAppIcon } from './icons';
import { Geometry, charsToPx, textWidth, cssFamily } from './geometry';
import { menus, type MenuItem } from '../shell/Menu';
import { XlMenuHost, xlMenu, type XlMenuItem } from './menu';
import { clipboard } from '../shell/clipboard';
import { session, logAppEvent } from '../shell/events';
import { useStore } from '../shell/store';
import { now } from '../util/clock';
import { blankWorkbook } from '../engine/blank.ts';
import { api, MAX_WORKBOOK_BYTES } from '../api';
import './excel.css';
import { BRAND, MS_EXCEL } from '../util/brand';

const ID = {
  pv: 'excel.protected_view_bar', tabs: 'excel.sheet_tabs', name: 'excel.name_box', fbar: 'excel.formula_bar', status: 'excel.status_bar_sum_count_average',
  save: 'excel.save_save_as_xlsx', copy: 'excel.copy_ctrl_c_marching_ants', find: 'excel.find_and_replace_ctrl_f_h', goto: 'excel.go_to_ctrl_g_f5',
};

/** Excel's message for a change a protected sheet does not allow */
const LOCKED_MSG = "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be requested to enter a password.";
const LAST_SHEET_MSG = 'A workbook must contain at least one visible worksheet.\n\nTo hide, delete, or move the selected sheet(s), you must first insert a new sheet or unhide a sheet that is already hidden.';
/** commands a protected sheet gates: 'edit' needs unlocked cells, the rest name a Protect Sheet allow-list flag */
const PROTECT_GATE: Record<string, string> = {
  bold: 'formatCells', italic: 'formatCells', underline: 'formatCells', fontName: 'formatCells', fontSize: 'formatCells', fontGrow: 'formatCells', border: 'formatCells',
  fontColor: 'formatCells', vAlign: 'formatCells', hAlign: 'formatCells', wrap: 'formatCells', indent: 'formatCells', rotate: 'formatCells', mergeCenter: 'formatCells',
  mergeCells: 'formatCells', mergeAcross: 'formatCells', unmerge: 'formatCells', numFmt: 'formatCells', decimals: 'formatCells', formatPainter: 'formatCells', formatCells: 'formatCells',
  insertRows: 'insertRows', deleteRows: 'deleteRows', insertCols: 'insertColumns', deleteCols: 'deleteColumns', autosum: 'edit', insertFunction: 'edit', clear: 'edit',
  sort: 'sort', filter: 'autoFilter', rowHeightDialog: 'formatRows', colWidthDialog: 'formatColumns', autofitCols: 'formatColumns', autofitRows: 'formatRows',
  hideRows: 'formatRows', unhideRows: 'formatRows', hideCols: 'formatColumns', unhideCols: 'formatColumns',
};

interface Clip { text: string; si: number; range: Range; cells: { dr: number; dc: number; cell: Cell | undefined }[]; cut: boolean; owner: string }
let excelClip: Clip | null = null;

type UndoEntry = { si: number; cells: { r: number; c: number; before: Cell | undefined }[] };

function stemOf(name: string) { const i = name.lastIndexOf('.'); return i > 0 ? name.slice(0, i) : name; }
function baseName(p: string) { return p.split('/').pop() ?? p; }

const fmtLabel = (code: string): string => {
  if (code === 'General') return 'General';
  if (code === '@') return 'Text';
  if (/%/.test(code)) return 'Percentage';
  if (/_\(.*\$/.test(code) || /_\(\* /.test(code)) return 'Accounting';
  if (/\$/.test(code)) return 'Currency';
  if (/E\+/.test(code)) return 'Scientific';
  if (/mmmm|dddd/.test(code)) return 'Long Date';
  if (/[dy]/.test(code.replace(/"[^"]*"/g, ''))) return 'Short Date';
  if (/[hs]/.test(code.replace(/"[^"]*"/g, ''))) return 'Time';
  if (/\?\/\?/.test(code)) return 'Fraction';
  return 'Number';
};

export function ExcelApp({ win, req }: AppProps) {
  const persona = useStore(session, s => s.persona);
  const args = (req.args ?? {}) as { protectedView?: boolean; url?: string; name?: string; source?: string };
  const [path, setPath] = useState<string | undefined>(req.file);
  const [name, setName] = useState<string>(req.file ? baseName(req.file) : (args.name ?? 'Book1'));
  const [wb, setWb] = useState<Workbook | null>(null);
  const [err, setErr] = useState<string | null>(null);
  const [pv, setPv] = useState<boolean>(!!args.protectedView);
  const [pvReason, setPvReason] = useState<'attachment' | 'internet'>(args.source === 'attachment' || args.url ? 'attachment' : 'internet');
  const [version, bump] = useReducer((x: number) => x + 1, 0);
  const [si, setSi] = useState(0);
  const selsRef = useRef(new Map<number, Sel>());
  const scrollRef = useRef(new Map<number, { top: number; left: number }>());
  const [editor, setEditor] = useState<EditorState | null>(null);
  const editorRef = useRef<HTMLTextAreaElement | null>(null);
  const fbarRef = useRef<HTMLInputElement | null>(null);
  const rootRef = useRef<HTMLDivElement | null>(null);
  const [tab, setTab] = useState<TabName>('Home');
  // the ribbon starts collapsed to its tabs (Show tabs only); Ctrl+F1 toggles
  const [collapsed, setCollapsed] = useState(true);
  const [floatTab, setFloatTab] = useState<TabName | null>(null);
  const [backstage, setBackstage] = useState<null | 'home' | 'info' | 'saveas' | 'open'>(null);
  const [zoom, setZoom] = useState(1);
  const [showGrid, setShowGrid] = useState(true);
  const [showHeadings, setShowHeadings] = useState(true);
  const [showFbar, setShowFbar] = useState(true);
  const [showFormulas, setShowFormulas] = useState(false);
  const [ants, setAnts] = useState<{ si: number; range: Range } | null>(null);
  const [filters, setFilters] = useState<Map<number, Range>>(new Map());
  const [dialog, setDialog] = useState<null | { kind: 'find' | 'replace' } | { kind: 'goto' } | { kind: 'colWidth' | 'rowHeight' } | { kind: 'zoom' } | { kind: 'format'; tab: 'number' | 'alignment' | 'font' } | { kind: 'insertFn' } | { kind: 'moveCopy' | 'protect' | 'unprotect'; si: number } | { kind: 'macro' }>(null);
  // Select All Sheets: every visible sheet is selected (title shows [Group]; entries go to every sheet)
  const [grouped, setGrouped] = useState(false);
  const [findOpts, setFindOptsState] = useState<FindOptions & { expanded: boolean }>({ what: '', repl: '', within: 'sheet', by: 'rows', lookIn: 'formulas', matchCase: false, entire: false, expanded: false });
  const [findResults, setFindResults] = useState<FindHit[] | null>(null);
  const [findStatus, setFindStatus] = useState('');
  const [renaming, setRenaming] = useState<number | null>(null);
  const [savedAt, setSavedAt] = useState<number | null>(null);
  const [dirty, setDirty] = useState(false);
  const viewport = useRef({ rows: 30, cols: 15 });
  const [thumbs, setThumbs] = useState<ScrollThumbs | null>(null);
  const undoStack = useRef<UndoEntry[]>([]);
  const redoStack = useRef<UndoEntry[]>([]);
  const painter = useRef<number | null>(null);
  const [nameBoxText, setNameBoxText] = useState<string | null>(null);
  const [focused, setFocused] = useState(true);
  const [rootW, setRootW] = useState(1440);
  // a modeless dialog with focus (Find and Replace) makes the workbook window inactive, as in Windows
  const [dlgFocus, setDlgFocus] = useState(false);

  // ---------------------------------------------------------------- load
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        let url: string;
        if (args.url) url = args.url;
        else if (path) url = `/api/files/raw?path=${encodeURIComponent(path)}`;
        else { const w = Workbook.load(blankWorkbook()); if (!cancelled) { setWb(w); setSi(0); } return; }
        const res = await fetch(url);
        if (!res.ok) throw new Error(`${BRAND.excel} cannot open the file '${name}' because the file format or file extension is not valid.`);
        const bytes = new Uint8Array(await res.arrayBuffer());
        const w = Workbook.load(bytes);
        w.nowMs = now();
        // Protected View only opens when the caller asks for it (args.protectedView); nothing turns it on by itself.
        if (cancelled) return;
        setWb(w);
        setSi(w.activeTab);
        logAppEvent('excel', 'open', { path: path ?? null, name, protectedView: !!args.protectedView, source: args.source ?? (path ? 'file' : 'new') });
      } catch (e) {
        if (!cancelled) setErr(e instanceof Error ? e.message : String(e));
      }
    })();
    return () => { cancelled = true; };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps


  const titleText = `${stemOf(name)}${pv ? ' [Protected View]' : ''}${grouped ? ' [Group]' : ''}`;
  useEffect(() => { win.setTitle(`${titleText} - ${BRAND.excel}`); }, [titleText, win]);

  useLayoutEffect(() => {
    const el = rootRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setRootW(el.clientWidth));
    ro.observe(el);
    setRootW(el.clientWidth);
    return () => ro.disconnect();
  }, [wb]);

  // focus tracking for the inactive title bar
  useEffect(() => {
    const t = setInterval(() => setFocused(win.isFocused()), 400);
    return () => clearInterval(t);
  }, [win]);

  const sheet = wb?.sheets[si];
  const sel: Sel = selsRef.current.get(si) ?? (sheet ? { active: sheet.active, anchor: sheet.active, range: { r1: sheet.active.r, c1: sheet.active.c, r2: sheet.active.r, c2: sheet.active.c }, kind: 'cells' } : { active: { r: 0, c: 0 }, anchor: { r: 0, c: 0 }, range: { r1: 0, c1: 0, r2: 0, c2: 0 }, kind: 'cells' });
  const scroll = scrollRef.current.get(si) ?? (sheet ? { top: sheet.freeze?.rows ? Math.max(sheet.topLeft.r, sheet.freeze.rows) : sheet.topLeft.r, left: sheet.freeze?.cols ? Math.max(sheet.topLeft.c, sheet.freeze.cols) : sheet.topLeft.c } : { top: 0, left: 0 });

  const markDirty = () => { setDirty(true); bump(); };

  // ---------------------------------------------------------------- sheet protection / groups
  /** the sheets an entry goes to: the active one, or every visible sheet while grouped (Select All Sheets) */
  const targetSheets = (): number[] => (wb && grouped ? wb.sheets.map((s, i) => (s.state === 'visible' ? i : -1)).filter(i => i >= 0) : [si]);
  const lockedIn = (g: Range, sIdx = si): boolean => {
    if (!wb) return false;
    const sh = wb.sheets[sIdx];
    const r2 = Math.min(g.r2, Math.max(g.r1, sh.maxRow - 1)), c2 = Math.min(g.c2, Math.max(g.c1, sh.maxCol - 1));
    let n = 0;
    for (let r = g.r1; r <= r2; r++) for (let c = g.c1; c <= c2; c++) { if (wb.style(wb.xfAt(sIdx, r, c)).locked || ++n > 50000) return true; }
    return g.r2 > r2 || g.c2 > c2 ? wb.style(wb.xfAt(sIdx, g.r2, g.c2)).locked : false;
  };
  /** a protected sheet that does not allow this ('edit' = change locked cells, else a Protect Sheet allow-list
   *  attribute): shows Excel's message and returns true */
  const blocked = (what: string, g: Range = sel.range): boolean => {
    if (!wb) return false;
    const hit = targetSheets().some(i => { const p = wb.sheets[i]?.protection; return !!p && (what === 'edit' ? lockedIn(g, i) : !protectionAllows(p, what)); });
    if (hit) void win.dialog.message({ title: MS_EXCEL, text: LOCKED_MSG, icon: 'warning' });
    return hit;
  };

  const setSel = useCallback((s: Sel) => { selsRef.current.set(si, s); setNameBoxText(null); bump(); }, [si]);
  const setScroll = useCallback((top: number, left: number) => {
    const fz = wb?.sheets[si]?.freeze;
    scrollRef.current.set(si, { top: Math.max(fz?.rows ?? 0, Math.min(MAX_ROW, top)), left: Math.max(fz?.cols ?? 0, Math.min(MAX_COL, left)) });
    bump();
  }, [si, wb]);

  const ensureVisible = (r: number, c: number) => {
    if (!sheet) return;
    const fz = sheet.freeze ?? { rows: 0, cols: 0 };
    let { top, left } = scroll;
    const vr = viewport.current.rows, vc = viewport.current.cols;
    if (r >= fz.rows) { if (r < top) top = r; else if (r >= top + vr) top = r - vr + 1; }
    if (c >= fz.cols) { if (c < left) left = c; else if (c >= left + vc) left = c - vc + 1; }
    if (top !== scroll.top || left !== scroll.left) setScroll(top, left);
  };

  const select = (a: Addr, extendFrom?: Addr, kind: Sel['kind'] = 'cells') => {
    const anchor = extendFrom ?? a;
    let range: Range = { r1: Math.min(anchor.r, a.r), c1: Math.min(anchor.c, a.c), r2: Math.max(anchor.r, a.r), c2: Math.max(anchor.c, a.c) };
    if (sheet) for (let pass = 0; pass < 3; pass++) for (const m of sheet.merges) if (m.r1 <= range.r2 && m.r2 >= range.r1 && m.c1 <= range.c2 && m.c2 >= range.c1) range = { r1: Math.min(range.r1, m.r1), c1: Math.min(range.c1, m.c1), r2: Math.max(range.r2, m.r2), c2: Math.max(range.c2, m.c2) };
    const s: Sel = extendFrom ? { active: sel.active, anchor, range, kind } : { active: a, anchor: a, range, kind };
    setSel(s);
    // keep the moving end visible
    ensureVisible(a.r, a.c);
  };

  // ---------------------------------------------------------------- editing
  const recordUndo = (cells: { r: number; c: number }[]) => {
    if (!wb) return;
    undoStack.current.push({ si, cells: cells.map(({ r, c }) => ({ r, c, before: cloneCell(wb.cell(si, r, c)) })) });
    if (undoStack.current.length > 100) undoStack.current.shift();
    redoStack.current = [];
  };
  const restore = (e: UndoEntry, into: UndoEntry[]) => {
    if (!wb) return;
    into.push({ si: e.si, cells: e.cells.map(({ r, c }) => ({ r, c, before: cloneCell(wb.cell(e.si, r, c)) })) });
    for (const { r, c, before } of e.cells) wb.setCellRaw(e.si, r, c, before ? before.v : null, before?.f, before?.s ?? 0);
    setSi(e.si);
    const first = e.cells[0];
    if (first) select({ r: first.r, c: first.c });
    markDirty();
  };

  const startEdit = (mode: 'enter' | 'edit', initial?: string, inFormulaBar = false) => {
    if (!wb || pv) return;
    const { r, c } = sel.active;
    if (blocked('edit', { r1: r, c1: c, r2: r, c2: c })) return;
    const text = initial ?? wb.editText(si, r, c);
    setEditor({ r, c, text, mode, caret: text.length });
    if (inFormulaBar) setTimeout(() => fbarRef.current?.focus(), 0);
  };

  const commit = (move: 'down' | 'up' | 'right' | 'left' | 'none' = 'down', ctrlEnter = false) => {
    if (!wb || !editor) return;
    const { r, c, text } = editor;
    const targets: Addr[] = ctrlEnter ? rangeCells(sel.range) : [{ r, c }];
    if (ctrlEnter && blocked('edit', sel.range)) return;
    recordUndo(targets);
    try {
      for (const t of targets) {
        let txt = text;
        if (ctrlEnter && text.startsWith('=') && (t.r !== r || t.c !== c)) txt = shiftFormula(text, t.r - r, t.c - c);
        // close an unbalanced formula the way Excel does
        if (txt.startsWith('=')) { const open = (txt.match(/\(/g) ?? []).length - (txt.match(/\)/g) ?? []).length; if (open > 0) txt += ')'.repeat(open); }
        // grouped sheets get the same entry, as in Excel
        for (const gi of targetSheets()) wb.setInput(gi, t.r, t.c, txt);
      }
    } catch (e) {
      undoStack.current.pop();
      void win.dialog.message({ title: MS_EXCEL, text: e instanceof Error ? e.message : String(e), icon: 'warning' });
      return;
    }
    setEditor(null);
    setAnts(null);
    markDirty();
    rootRef.current?.focus({ preventScroll: true });
    const g = sel.range;
    const inRangeMove = !(g.r1 === g.r2 && g.c1 === g.c2);
    if (move === 'none' || ctrlEnter) return;
    if (inRangeMove) {
      // Enter cycles inside a selected range
      let { r: nr, c: nc } = sel.active;
      if (move === 'down') { nr++; if (nr > g.r2) { nr = g.r1; nc = nc + 1 > g.c2 ? g.c1 : nc + 1; } }
      else if (move === 'up') { nr--; if (nr < g.r1) { nr = g.r2; nc = nc - 1 < g.c1 ? g.c2 : nc - 1; } }
      else if (move === 'right') { nc++; if (nc > g.c2) { nc = g.c1; nr = nr + 1 > g.r2 ? g.r1 : nr + 1; } }
      else if (move === 'left') { nc--; if (nc < g.c1) { nc = g.c2; nr = nr - 1 < g.r1 ? g.r2 : nr - 1; } }
      setSel({ ...sel, active: { r: nr, c: nc } });
      return;
    }
    const d = { down: [1, 0], up: [-1, 0], right: [0, 1], left: [0, -1] }[move];
    const m = sheet?.merges.find(mm => mm.r1 === r && mm.c1 === c);
    const nr = Math.max(0, Math.min(MAX_ROW, (d[0] > 0 && m ? m.r2 : r) + d[0]));
    const nc = Math.max(0, Math.min(MAX_COL, (d[1] > 0 && m ? m.c2 : c) + d[1]));
    select({ r: nr, c: nc });
  };
  const cancelEdit = () => { setEditor(null); rootRef.current?.focus({ preventScroll: true }); };

  /** formula point mode: after an operator, comma or '(' a click or arrow inserts a reference */
  const lastPoint = useRef<string | null>(null);
  const pointing = (() => {
    if (!editor || !editor.text.startsWith('=')) return false;
    let before = editor.text.slice(0, editor.caret ?? editor.text.length);
    if (lastPoint.current && before.endsWith(lastPoint.current)) before = before.slice(0, -lastPoint.current.length);
    return /[=+\-*/^&,(<>:]$/.test(before);
  })();
  const onPointRef = (ref: string, extend: boolean) => {
    if (!editor) return;
    const caret = editor.caret ?? editor.text.length;
    let before = editor.text.slice(0, caret);
    const after = editor.text.slice(caret);
    if ((extend || lastPoint.current) && lastPoint.current && before.endsWith(lastPoint.current)) before = before.slice(0, -lastPoint.current.length);
    const sheetRef = ref;
    lastPoint.current = sheetRef;
    const text = before + sheetRef + after;
    setEditor({ ...editor, text, caret: before.length + sheetRef.length });
  };

  // ---------------------------------------------------------------- navigation helpers
  const nonEmpty = (r: number, c: number) => { const v = wb?.cell(si, r, c); return !!v && (v.v !== null || !!v.f); };
  const edge = (a: Addr, dr: number, dc: number): Addr => {
    let { r, c } = a;
    const inb = (rr: number, cc: number) => rr >= 0 && cc >= 0 && rr <= MAX_ROW && cc <= MAX_COL;
    if (!inb(r + dr, c + dc)) return a;
    const maxR = sheet?.maxRow ?? 0, maxC = sheet?.maxCol ?? 0;
    if (nonEmpty(r, c) && nonEmpty(r + dr, c + dc)) {
      while (inb(r + dr, c + dc) && nonEmpty(r + dr, c + dc)) { r += dr; c += dc; }
      return { r, c };
    }
    r += dr; c += dc;
    while (inb(r, c) && !nonEmpty(r, c)) {
      if ((dr > 0 && r >= maxR) || (dc > 0 && c >= maxC)) return { r: dr ? MAX_ROW : r, c: dc ? MAX_COL : c };
      r += dr; c += dc;
    }
    if (!inb(r, c)) return { r: Math.max(0, Math.min(MAX_ROW, r)), c: Math.max(0, Math.min(MAX_COL, c)) };
    return { r, c };
  };
  const currentRegion = (a: Addr): Range => {
    let g: Range = { r1: a.r, c1: a.c, r2: a.r, c2: a.c };
    for (let changed = true, n = 0; changed && n < 50; n++) {
      changed = false;
      const grow = (test: () => boolean, apply: () => void) => { if (test()) { apply(); changed = true; } };
      const rowHas = (r: number) => { for (let c = Math.max(0, g.c1 - 1); c <= g.c2 + 1; c++) if (nonEmpty(r, c)) return true; return false; };
      const colHas = (c: number) => { for (let r = Math.max(0, g.r1 - 1); r <= g.r2 + 1; r++) if (nonEmpty(r, c)) return true; return false; };
      grow(() => g.r1 > 0 && rowHas(g.r1 - 1), () => { g = { ...g, r1: g.r1 - 1 }; });
      grow(() => rowHas(g.r2 + 1), () => { g = { ...g, r2: g.r2 + 1 }; });
      grow(() => g.c1 > 0 && colHas(g.c1 - 1), () => { g = { ...g, c1: g.c1 - 1 }; });
      grow(() => colHas(g.c2 + 1), () => { g = { ...g, c2: g.c2 + 1 }; });
    }
    return g;
  };

  // ---------------------------------------------------------------- clipboard
  const usedClip = (g: Range): Range => {
    if (!sheet) return g;
    return { r1: g.r1, c1: g.c1, r2: Math.min(g.r2, Math.max(g.r1, sheet.maxRow - 1)), c2: Math.min(g.c2, Math.max(g.c1, sheet.maxCol - 1)) };
  };
  const doCopy = (cut: boolean) => {
    if (!wb) return;
    if (cut && blocked('edit')) return;
    const g = usedClip(sel.range);
    const rows: string[] = [];
    const html: string[] = [];
    const cells: Clip['cells'] = [];
    for (let r = g.r1; r <= g.r2; r++) {
      const t: string[] = [], h: string[] = [];
      if (wb.rowHidden(si, r)) continue;
      for (let c = g.c1; c <= g.c2; c++) {
        if (wb.colHidden(si, c)) continue;
        const d = wb.display(si, r, c);
        const st = wb.style(wb.xfAt(si, r, c));
        t.push(/[\t\n"]/.test(d.text) ? `"${d.text.replace(/"/g, '""')}"` : d.text);
        const css = [`font-family:${cssFamily(st.font.name).replace(/"/g, "'")}`, `font-size:${st.font.sz || 11}pt`, st.font.b ? 'font-weight:700' : '', st.font.i ? 'font-style:italic' : '', st.font.color ? `color:${st.font.color}` : '',
          st.fill.pattern === 'solid' && st.fill.fg ? `background:${st.fill.fg}` : '', typeof d.value === 'number' ? 'text-align:right' : '', 'white-space:nowrap'].filter(Boolean).join(';');
        h.push(`<td style="${css}">${escHtml(d.text)}</td>`);
        cells.push({ dr: r - g.r1, dc: c - g.c1, cell: cloneCell(wb.cell(si, r, c)) });
      }
      rows.push(t.join('\t'));
      html.push(`<tr style="height:${Math.round(wb.rowHeightPt(si, r) * 4 / 3)}px">${h.join('')}</tr>`);
    }
    const cols = [];
    for (let c = g.c1; c <= g.c2; c++) if (!wb.colHidden(si, c)) cols.push(`<col width="${charsToPx(wb.colWidthChars(si, c))}">`);
    const text = rows.join('\r\n') + '\r\n';
    const htmlDoc = `<table border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse">${cols.join('')}${html.join('')}</table>`;
    excelClip = { text, si, range: g, cells, cut, owner: win.id };
    window.dispatchEvent(new CustomEvent('sheet:clipboard-write', { detail: { text, html: htmlDoc, source: 'excel' } }));
    setAnts({ si, range: g });
    logAppEvent('excel', cut ? 'cut' : 'copy', { range: `${addrName(g.r1, g.c1)}:${addrName(g.r2, g.c2)}` });
  };
  const doPaste = () => {
    if (!wb || pv) return;
    if (blocked('edit')) return;
    const data = clipboard.read();
    const at = sel.active;
    if (excelClip && data && data.text === excelClip.text && data.source === 'excel') {
      const clip = excelClip;
      const h = clip.range.r2 - clip.range.r1 + 1, w = clip.range.c2 - clip.range.c1 + 1;
      const targets: Addr[] = [];
      for (let dr = 0; dr < h; dr++) for (let dc = 0; dc < w; dc++) targets.push({ r: at.r + dr, c: at.c + dc });
      const sources = clip.cut ? rangeCells(clip.range).filter(a => !targets.some(t => t.r === a.r && t.c === a.c)) : [];
      recordUndo([...targets, ...(clip.si === si ? sources : [])]);
      for (const { dr, dc, cell } of clip.cells) {
        const f = cell?.f ? (clip.cut ? cell.f : shiftFormula('=' + cell.f, at.r - clip.range.r1, at.c - clip.range.c1).slice(1)) : undefined;
        wb.setCellRaw(si, at.r + dr, at.c + dc, cell ? cell.v : null, f, cell?.s ?? 0);
      }
      if (clip.cut) {
        for (const a of sources) wb.setCellRaw(clip.si, a.r, a.c, null, undefined, 0);
        excelClip = null; setAnts(null);
      }
      select({ r: at.r + h - 1, c: at.c + w - 1 }, at);
      markDirty();
      return;
    }
    if (!data?.text) return;
    const lines = data.text.replace(/\r\n/g, '\n').replace(/\n$/, '').split('\n');
    const grid = lines.map(l => l.split('\t'));
    const targets: Addr[] = [];
    grid.forEach((row, dr) => row.forEach((_, dc) => targets.push({ r: at.r + dr, c: at.c + dc })));
    recordUndo(targets);
    grid.forEach((row, dr) => row.forEach((v, dc) => wb.setInput(si, at.r + dr, at.c + dc, v.replace(/^"(.*)"$/s, '$1').replace(/""/g, '"'))));
    const w = Math.max(...grid.map(r => r.length));
    select({ r: at.r + grid.length - 1, c: at.c + w - 1 }, at);
    markDirty();
  };

  // ---------------------------------------------------------------- formatting
  const fmt = (patch: Parameters<Workbook['formatRange']>[2], range = sel.range) => {
    if (!wb || pv) return;
    const g = usedClip(range);
    recordUndo(rangeCells({ ...g, r2: Math.min(g.r2, g.r1 + 2000), c2: Math.min(g.c2, g.c1 + 200) }));
    for (const gi of targetSheets()) wb.formatRange(gi, range, patch);
    markDirty();
  };
  const activeStyle = wb ? wb.style(wb.xfAt(si, sel.active.r, sel.active.c)) : null;

  const numFmtFor = (label: string): string | null => {
    const cur = activeStyle?.numFmt ?? 'General';
    switch (label) {
      case 'General': return 'General';
      case 'Number': return '0.00';
      case 'Currency': return '"$"#,##0.00';
      case 'Accounting': return NUMBER_CATEGORIES[3].code(2);
      case 'Short Date': return 'm/d/yyyy';
      case 'Long Date': return '[$-F800]dddd, mmmm dd, yyyy'.replace('[$-F800]', '');
      case 'Time': return '[$-F400]h:mm:ss AM/PM'.replace('[$-F400]', '');
      case 'Percentage': return '0.00%';
      case 'Percentage0': return '0%';
      case 'Fraction': return '# ?/?';
      case 'Scientific': return '0.00E+00';
      case 'Text': return '@';
      case 'Comma': return '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
      default: return cur;
    }
  };

  // ---------------------------------------------------------------- find / replace
  const searchCells = (o: FindOptions): FindHit[] => {
    if (!wb) return [];
    const out: FindHit[] = [];
    const sheets = o.within === 'workbook' ? wb.sheets.map((_, i) => i) : [si];
    const norm = (s: string) => (o.matchCase ? s : s.toLowerCase());
    const what = norm(o.what);
    for (const i of sheets) {
      const sh = wb.sheets[i];
      const keys = [...sh.cells.keys()].sort((a, b) => {
        const ra = Math.floor(a / 16384), ca = a % 16384, rb = Math.floor(b / 16384), cb = b % 16384;
        return o.by === 'rows' ? ra - rb || ca - cb : ca - cb || ra - rb;
      });
      for (const k of keys) {
        const r = Math.floor(k / 16384), c = k % 16384;
        const cell = sh.cells.get(k)!;
        if (cell.v === null && !cell.f) continue;
        const shown = wb.display(i, r, c).text;
        const hay = o.lookIn === 'values' ? shown : cell.f ? '=' + cell.f : wb.editText(i, r, c);
        const pattern = what.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?');
        const re = new RegExp(o.entire ? `^${pattern}$` : pattern, o.matchCase ? '' : 'i');
        if (re.test(hay)) out.push({ si: i, sheet: sh.name, r, c, addr: addrName(r, c), value: shown, formula: cell.f ? '=' + cell.f : '' });
      }
    }
    return out;
  };
  const findNext = () => {
    const hits = searchCells(findOpts);
    if (!hits.length) { void win.dialog.message({ title: MS_EXCEL, text: "We couldn't find what you were looking for. Select Options for more ways to search.", icon: 'info' }); return; }
    const cur = { si, r: sel.active.r, c: sel.active.c };
    const order = (h: { si: number; r: number; c: number }) => h.si * 1e12 + (findOpts.by === 'rows' ? h.r * 16384 + h.c : h.c * 2e6 + h.r);
    const next = hits.find(h => order(h) > order(cur)) ?? hits[0];
    goToHit(next);
  };
  const goToHit = (h: FindHit) => { setSi(h.si); selsRef.current.set(h.si, { active: { r: h.r, c: h.c }, anchor: { r: h.r, c: h.c }, range: { r1: h.r, c1: h.c, r2: h.r, c2: h.c }, kind: 'cells' }); setTimeout(() => ensureVisible(h.r, h.c), 0); bump(); };
  const findAll = () => { const hits = searchCells(findOpts); setFindResults(hits); setFindStatus(`${hits.length} cell(s) found`); if (hits[0]) goToHit(hits[0]); };
  const replaceOne = (hit: FindHit) => {
    if (!wb) return;
    const cell = wb.cell(hit.si, hit.r, hit.c);
    const src = cell?.f ? '=' + cell.f : wb.editText(hit.si, hit.r, hit.c);
    const pattern = findOpts.what.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
    const re = new RegExp(findOpts.entire ? `^${pattern}$` : pattern, findOpts.matchCase ? 'g' : 'gi');
    wb.setInput(hit.si, hit.r, hit.c, src.replace(re, findOpts.repl));
  };
  const doReplace = () => {
    if (!wb || pv) return;
    const hits = searchCells({ ...findOpts, lookIn: 'formulas' });
    const curHit = hits.find(h => h.si === si && h.r === sel.active.r && h.c === sel.active.c);
    if (curHit && blocked('edit', { r1: curHit.r, c1: curHit.c, r2: curHit.r, c2: curHit.c })) return;
    if (curHit) { recordUndo([{ r: curHit.r, c: curHit.c }]); replaceOne(curHit); markDirty(); }
    findNext();
  };
  const doReplaceAll = () => {
    if (!wb || pv) return;
    const hits = searchCells({ ...findOpts, lookIn: 'formulas' });
    if (!hits.length) { void win.dialog.message({ title: MS_EXCEL, text: "We couldn't find anything to replace. Click Options for more ways to search.", icon: 'info' }); return; }
    if (hits.some(h => wb.sheets[h.si].protection && lockedIn({ r1: h.r, c1: h.c, r2: h.r, c2: h.c }, h.si))) { void win.dialog.message({ title: MS_EXCEL, text: LOCKED_MSG, icon: 'warning' }); return; }
    recordUndo(hits.filter(h => h.si === si).map(h => ({ r: h.r, c: h.c })));
    hits.forEach(replaceOne);
    markDirty();
    void win.dialog.message({ title: MS_EXCEL, text: `All done. We made ${hits.length} replacement${hits.length === 1 ? '' : 's'}.`, icon: 'info' });
    logAppEvent('excel', 'replace_all', { count: hits.length });
  };

  const goTo = (ref: string) => {
    if (!wb) return;
    let target = ref.trim();
    let sIdx = si;
    const bang = target.lastIndexOf('!');
    if (bang > 0) { const sn = target.slice(0, bang).replace(/^'|'$/g, ''); const i = wb.sheetIndex(sn); if (i < 0) return false; sIdx = i; target = target.slice(bang + 1); }
    const dn = wb.definedNames.get(target.toUpperCase());
    if (dn) return goTo(dn.replace(/\$/g, ''));
    const g = parseRange(target);
    if (!g) { void win.dialog.message({ title: MS_EXCEL, text: 'Reference isn\'t valid.', icon: 'error' }); return false; }
    setSi(sIdx);
    selsRef.current.set(sIdx, { active: { r: g.r1, c: g.c1 }, anchor: { r: g.r1, c: g.c1 }, range: g, kind: g.r1 === 0 && g.r2 === MAX_ROW ? 'cols' : g.c1 === 0 && g.c2 === MAX_COL ? 'rows' : 'cells' });
    const fz = wb.sheets[sIdx].freeze;
    // Go To leaves the view alone when the target is on screen; otherwise it scrolls the target to the top-left
    const cur = scrollRef.current.get(sIdx) ?? { top: fz?.rows ?? 0, left: fz?.cols ?? 0 };
    const vr = viewport.current.rows, vc = viewport.current.cols;
    const rowVisible = g.r1 < (fz?.rows ?? 0) || (g.r1 >= cur.top && g.r1 < cur.top + vr);
    const colVisible = g.c1 < (fz?.cols ?? 0) || (g.c1 >= cur.left && g.c1 < cur.left + vc);
    if (!(rowVisible && colVisible) || sIdx !== si) scrollRef.current.set(sIdx, { top: rowVisible && sIdx === si ? cur.top : Math.max(fz?.rows ?? 0, g.r1), left: colVisible && sIdx === si ? cur.left : Math.max(fz?.cols ?? 0, g.c1 < vc ? 0 : g.c1) });
    bump();
    return true;
  };

  // ---------------------------------------------------------------- structure (rows / columns / sheets)
  const shiftCells = (axis: 'row' | 'col', at: number, n: number) => {
    if (!wb || !sheet) return;
    // move every cell at/after `at` by n (n < 0 deletes), adjusting formulas everywhere
    const all = [...sheet.cells.entries()].map(([k, cell]) => ({ r: Math.floor(k / 16384), c: k % 16384, cell }));
    const moving = all.filter(x => (axis === 'row' ? x.r : x.c) >= at);
    const sorted = moving.sort((a, b) => (n > 0 ? -1 : 1) * ((axis === 'row' ? a.r - b.r : a.c - b.c)));
    recordUndo([]);
    for (const x of sorted) {
      const nr = axis === 'row' ? x.r + n : x.r, nc = axis === 'col' ? x.c + n : x.c;
      const deleted = n < 0 && (axis === 'row' ? x.r < at - n : x.c < at - n);
      wb.setCellRaw(si, x.r, x.c, null, undefined, 0);
      if (!deleted && nr >= 0 && nc >= 0) wb.setCellRaw(si, nr, nc, x.cell.v, x.cell.f, x.cell.s);
    }
    // formulas that point at moved cells follow them
    const sheetName = sheet.name;
    wb.sheets.forEach((sh, i) => {
      for (const [k, cell] of sh.cells) {
        if (!cell.f) continue;
        const next = adjustRefs(cell.f, sh.name === sheetName, sheetName, axis, at, n);
        if (next !== cell.f) wb.setCellRaw(i, Math.floor(k / 16384), k % 16384, null, next, cell.s);
      }
    });
    undoStack.current.pop();
    markDirty();
  };

  // ---------------------------------------------------------------- commands
  const act = (cmd: string, arg?: unknown) => {
    if (!wb) return;
    setFloatTab(null);
    const g = sel.range;
    const gate = cmd === 'fill' ? (['down', 'up', 'left', 'right'].includes(String(arg)) ? 'edit' : 'formatCells') : PROTECT_GATE[cmd];
    if (gate && !pv && blocked(gate)) return;
    switch (cmd) {
      case 'copy': doCopy(false); break;
      case 'cut': if (!pv) doCopy(true); break;
      case 'paste': doPaste(); break;
      case 'bold': fmt({ font: { b: !activeStyle?.font.b } }); break;
      case 'italic': fmt({ font: { i: !activeStyle?.font.i } }); break;
      case 'underline': fmt({ font: { u: activeStyle?.font.u ? undefined : 'single' } }); break;
      case 'fontName': fmt({ font: { name: String(arg), scheme: undefined } }); break;
      case 'fontSize': if (Number(arg) > 0) fmt({ font: { sz: Number(arg) } }); break;
      case 'fontGrow': { const sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]; const cur = activeStyle?.font.sz ?? 11; const i = sizes.findIndex(s => s >= cur); const next = Number(arg) > 0 ? sizes[Math.min(sizes.length - 1, sizes[i] === cur ? i + 1 : i)] : sizes[Math.max(0, i - 1)]; fmt({ font: { sz: next } }); break; }
      case 'border': fmt({ border: arg as 'all' }); break;
      case 'fontColor': fmt({ font: { color: (arg as string | null) ?? undefined } }); break;
      case 'vAlign': fmt({ align: { v: String(arg) } }); break;
      case 'hAlign': fmt({ align: { h: activeStyle?.align.h === arg ? undefined : String(arg) } }); break;
      case 'wrap': fmt({ align: { wrap: !activeStyle?.align.wrap } }); break;
      case 'indent': fmt({ align: { indent: Math.max(0, (activeStyle?.align.indent ?? 0) + Number(arg)) || undefined, h: activeStyle?.align.h ?? 'left' } }); break;
      case 'rotate': fmt({ align: { rotation: Number(arg) } }); break;
      case 'mergeCenter': case 'mergeCells': case 'mergeAcross': {
        if (pv) break;
        const m = sheet?.merges.find(mm => mm.r1 === g.r1 && mm.c1 === g.c1 && mm.r2 === g.r2 && mm.c2 === g.c2);
        if (m && cmd === 'mergeCenter') { wb.unmerge(si, g); markDirty(); break; }
        if (g.r1 === g.r2 && g.c1 === g.c2) break;
        const filled = rangeCells(usedClip(g)).filter(a => nonEmpty(a.r, a.c));
        const go = () => {
          recordUndo(rangeCells(usedClip(g)));
          if (cmd === 'mergeAcross') for (let r = g.r1; r <= g.r2; r++) wb.merge(si, { r1: r, c1: g.c1, r2: r, c2: g.c2 }, false);
          else wb.merge(si, g, cmd === 'mergeCenter');
          markDirty();
        };
        if (filled.length > 1) void win.dialog.message({ title: MS_EXCEL, text: 'Merging cells only keeps the upper-left value and discards other values.', icon: 'warning', buttons: ['OK', 'Cancel'] }).then(b => { if (b === 'OK') go(); });
        else go();
        break;
      }
      case 'unmerge': if (!pv) { wb.unmerge(si, g); markDirty(); } break;
      case 'numFmt': { const code = numFmtFor(String(arg)); if (code) fmt({ numFmt: code }); break; }
      case 'decimals': {
        const cur = activeStyle?.numFmt ?? 'General';
        let code = cur;
        if (cur === 'General') { const v = wb.value(si, sel.active.r, sel.active.c); const d = typeof v === 'number' ? (String(v).split('.')[1]?.length ?? 0) : 0; code = Number(arg) > 0 ? `0.${'0'.repeat(d + 1)}` : d > 1 ? `0.${'0'.repeat(d - 1)}` : '0'; }
        else if (Number(arg) > 0) code = /\.0+/.test(cur) ? cur.replace(/\.(0+)/g, (_, z: string) => `.${z}0`) : cur.replace(/(0)(?![\d.])/, '$1.0');
        else code = cur.replace(/\.(0+)/g, (_, z: string) => (z.length > 1 ? `.${z.slice(1)}` : ''));
        fmt({ numFmt: code });
        break;
      }
      case 'insertRows': if (!pv) { shiftCells('row', g.r1, g.r2 - g.r1 + 1); } break;
      case 'deleteRows': if (!pv) { shiftCells('row', g.r1, -(g.r2 - g.r1 + 1)); } break;
      case 'insertCols': if (!pv) { shiftCells('col', g.c1, g.c2 - g.c1 + 1); } break;
      case 'deleteCols': if (!pv) { shiftCells('col', g.c1, -(g.c2 - g.c1 + 1)); } break;
      case 'insertSheet': if (!pv) { const i = wb.addSheet(si); setGrouped(false); setSi(i); markDirty(); } break;
      case 'moveCopySheet': if (!pv) setDialog({ kind: 'moveCopy', si }); break;
      case 'protectSheet': if (!pv) protectSheet(si); break;
      case 'deleteSheet': deleteSheet(si); break;
      case 'renameSheet': if (!pv) setRenaming(si); break;
      case 'autosum': {
        if (pv) break;
        const fn = String(arg ?? 'SUM');
        const { r, c } = sel.active;
        let rng = '';
        if (g.r1 !== g.r2 || g.c1 !== g.c2) {
          // a selected block: total below each column
          recordUndo(Array.from({ length: g.c2 - g.c1 + 1 }, (_, i) => ({ r: g.r2 + 1, c: g.c1 + i })));
          for (let cc = g.c1; cc <= g.c2; cc++) wb.setInput(si, g.r2 + 1, cc, `=${fn}(${addrName(g.r1, cc)}:${addrName(g.r2, cc)})`);
          markDirty();
          break;
        }
        let r0 = r - 1;
        while (r0 >= 0 && typeof wb.value(si, r0, c) === 'number') r0--;
        if (r0 < r - 1) rng = `${addrName(r0 + 1, c)}:${addrName(r - 1, c)}`;
        else {
          let c0 = c - 1;
          while (c0 >= 0 && typeof wb.value(si, r, c0) === 'number') c0--;
          if (c0 < c - 1) rng = `${addrName(r, c0 + 1)}:${addrName(r, c - 1)}`;
        }
        const text = `=${fn}(${rng})`;
        lastPoint.current = rng || null;
        setEditor({ r, c, text, mode: 'edit', caret: text.length - 1 });
        break;
      }
      case 'insertFunction': {
        if (pv) break;
        if (!arg) { setDialog({ kind: 'insertFn' }); break; }
        const { r, c } = sel.active;
        const text = editor ? editor.text + `${arg}(` : `=${arg}(`;
        lastPoint.current = null;
        setEditor({ r, c, text, mode: 'edit', caret: text.length });
        break;
      }
      case 'fill': {
        // One command name, two Excel features: Home > Fill Color passes a colour (or null for
        // No Fill); Home > Editing > Fill and Ctrl+D / Ctrl+R pass a direction.
        if (arg !== 'down' && arg !== 'up' && arg !== 'left' && arg !== 'right') { fmt({ fill: (arg as string | null) }); break; }
        if (pv) break;
        const dir = String(arg);
        // clip the cross axis to the used range, but keep the fill axis as selected (the targets past the last
        // used row/column are the point of Fill Down/Right); only a whole-column/row selection is cut back
        const u = usedClip(g);
        const vert = dir === 'down' || dir === 'up';
        const span = vert ? g.r2 - g.r1 : g.c2 - g.c1;
        const fr = span > 5000 ? u : vert ? { ...u, r2: g.r2 } : { ...u, c2: g.c2 };
        const targets = rangeCells(fr);
        recordUndo(targets);
        for (const t of targets) {
          const src = dir === 'down' ? { r: g.r1, c: t.c } : dir === 'up' ? { r: g.r2, c: t.c } : dir === 'right' ? { r: t.r, c: g.c1 } : { r: t.r, c: g.c2 };
          if (src.r === t.r && src.c === t.c) continue;
          const cell = wb.cell(si, src.r, src.c);
          const f = cell?.f ? shiftFormula('=' + cell.f, t.r - src.r, t.c - src.c).slice(1) : undefined;
          wb.setCellRaw(si, t.r, t.c, cell ? cell.v : null, f, cell?.s ?? 0);
        }
        markDirty();
        break;
      }
      case 'clear': {
        if (pv) break;
        const targets = rangeCells(usedClip(g));
        recordUndo(targets);
        for (const t of targets) {
          const cell = wb.cell(si, t.r, t.c);
          if (!cell) continue;
          if (arg === 'contents') wb.setCellRaw(si, t.r, t.c, null, undefined, cell.s);
          else if (arg === 'formats') wb.setCellRaw(si, t.r, t.c, cell.v, cell.f, 0);
          else wb.setCellRaw(si, t.r, t.c, null, undefined, 0);
        }
        markDirty();
        break;
      }
      case 'sort': {
        if (pv) break;
        const region = g.r1 === g.r2 && g.c1 === g.c2 ? currentRegion(sel.active) : usedClip(g);
        const header = region.r2 > region.r1 && typeof wb.value(si, region.r1, sel.active.c) === 'string' && wb.style(wb.xfAt(si, region.r1, sel.active.c)).font.b;
        const r1 = header ? region.r1 + 1 : region.r1;
        const col = sel.active.c;
        const rows: { cells: (Cell | undefined)[]; key: Scalar }[] = [];
        for (let r = r1; r <= region.r2; r++) {
          const cells: (Cell | undefined)[] = [];
          for (let c = region.c1; c <= region.c2; c++) cells.push(cloneCell(wb.cell(si, r, c)));
          rows.push({ cells, key: wb.value(si, r, col) });
        }
        const dir = Number(arg) || 1;
        const rank = (v: Scalar) => (v === null ? 3 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
        rows.sort((a, b) => {
          const ra = rank(a.key), rb = rank(b.key);
          if (a.key === null || b.key === null) return ra - rb; // blanks last either way
          if (ra !== rb) return dir * (ra - rb);
          if (typeof a.key === 'number' && typeof b.key === 'number') return dir * (a.key - b.key);
          return dir * String(a.key).localeCompare(String(b.key), undefined, { sensitivity: 'base' });
        });
        recordUndo(rangeCells({ r1, c1: region.c1, r2: region.r2, c2: region.c2 }));
        rows.forEach((row, i) => row.cells.forEach((cell, j) => wb.setCellRaw(si, r1 + i, region.c1 + j, cell ? cell.v : null, cell?.f, cell?.s ?? 0)));
        markDirty();
        break;
      }
      case 'filter': {
        const next = new Map(filters);
        if (next.has(si)) {
          // clearing the filter shows hidden rows again
          const fr = next.get(si)!;
          const hidden: number[] = [];
          for (let r = fr.r1 + 1; r <= fr.r2; r++) if (wb.rowHidden(si, r)) hidden.push(r);
          if (hidden.length) { wb.setRowHeight(si, hidden[0], hidden[hidden.length - 1], null, false); markDirty(); }
          next.delete(si);
        } else next.set(si, currentRegion(sel.active));
        setFilters(next);
        break;
      }
      case 'find': setDialog({ kind: 'find' }); setFindResults(null); setFindStatus(''); break;
      case 'replace': if (!pv) { setDialog({ kind: 'replace' }); setFindResults(null); setFindStatus(''); } break;
      case 'goto': setDialog({ kind: 'goto' }); break;
      case 'selectSpecial': {
        const hits: Addr[] = [];
        for (const [k, cell] of sheet!.cells) if (arg === 'formulas' ? !!cell.f : (cell.v !== null && !cell.f)) hits.push({ r: Math.floor(k / 16384), c: k % 16384 });
        if (!hits.length) { void win.dialog.message({ title: MS_EXCEL, text: 'No cells were found.', icon: 'info' }); break; }
        const rr = { r1: Math.min(...hits.map(h => h.r)), c1: Math.min(...hits.map(h => h.c)), r2: Math.max(...hits.map(h => h.r)), c2: Math.max(...hits.map(h => h.c)) };
        setSel({ active: hits[0], anchor: hits[0], range: rr, kind: 'cells' });
        break;
      }
      case 'formatPainter': {
        if (pv) break;
        painter.current = wb.xfAt(si, sel.active.r, sel.active.c);
        setAnts({ si, range: g });
        const once = () => {
          window.removeEventListener('mouseup', once, true);
          setTimeout(() => {
            const xf = painter.current; painter.current = null;
            if (xf === null) return;
            const s = selsRef.current.get(siRef.current);
            if (!s) return;
            const tgt = rangeCells(usedClip(s.range));
            recordUndo(tgt);
            for (const t of tgt) { const cell = wb.cell(siRef.current, t.r, t.c); wb.setCellRaw(siRef.current, t.r, t.c, cell ? cell.v : null, cell?.f, xf); }
            setAnts(null); markDirty();
          }, 0);
        };
        setTimeout(() => window.addEventListener('mouseup', once, true), 0);
        break;
      }
      case 'toggleGrid': setShowGrid(v => !v); break;
      case 'toggleHeadings': setShowHeadings(v => !v); break;
      case 'toggleFormulaBar': setShowFbar(v => !v); break;
      case 'toggleFormulas': setShowFormulas(v => !v); break;
      case 'zoom': setZoom(Number(arg) || 1); break;
      case 'zoomDialog': setDialog({ kind: 'zoom' }); break;
      case 'zoomSel': {
        const geo = new Geometry(wb, si, 1);
        let w = 0; for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 200); c++) w += geo.colW(c);
        let h = 0; for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 500); r++) h += geo.rowH(r);
        const box = rootRef.current?.querySelector('.xl-grid')?.getBoundingClientRect();
        if (box) { setZoom(Math.max(0.1, Math.min(4, Math.min((box.width - 40) / Math.max(w, 1), (box.height - 30) / Math.max(h, 1))))); setScroll(g.r1, g.c1); }
        break;
      }
      case 'freeze': if (sel.active.r || sel.active.c) { wb.setFreeze(si, sel.active.r, sel.active.c); setScroll(Math.max(scroll.top, sel.active.r), Math.max(scroll.left, sel.active.c)); markDirty(); } break;
      case 'unfreeze': wb.setFreeze(si, 0, 0); markDirty(); break;
      case 'freezeTopRow': wb.setFreeze(si, 1, 0); setScroll(Math.max(1, scroll.top), scroll.left); markDirty(); break;
      case 'freezeFirstCol': wb.setFreeze(si, 0, 1); setScroll(scroll.top, Math.max(1, scroll.left)); markDirty(); break;
      case 'calculate': wb.recalc(); bump(); break;
      case 'spelling': void win.dialog.message({ title: MS_EXCEL, text: 'Spell check complete. You\'re good to go!', icon: 'info' }); break;
      case 'accessibility': void win.dialog.message({ title: 'Accessibility', text: 'No accessibility issues found. People with disabilities should not have difficulty reading this workbook.', icon: 'info' }); break;
      case 'help': void win.dialog.message({ title: 'Help', text: `Help content is not available. Your organization manages help and support for ${BRAND.suite}.`, icon: 'info' }); break;
      case 'rowHeightDialog': setDialog({ kind: 'rowHeight' }); break;
      case 'colWidthDialog': setDialog({ kind: 'colWidth' }); break;
      case 'autofitCols': for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 100); c++) autofit(c); break;
      case 'autofitRows': wb.setRowHeight(si, g.r1, Math.min(g.r2, g.r1 + 500), 15, undefined); markDirty(); break;
      case 'hideRows': wb.setRowHeight(si, g.r1, Math.min(g.r2, g.r1 + 5000), null, true); markDirty(); break;
      case 'unhideRows': wb.setRowHeight(si, Math.max(0, g.r1 - 1), Math.min(g.r2 + 1, g.r1 + 5000), null, false); markDirty(); break;
      case 'hideCols': wb.setColWidth(si, g.c1, Math.min(g.c2, g.c1 + 500), null, true); markDirty(); break;
      case 'unhideCols': wb.setColWidth(si, Math.max(0, g.c1 - 1), Math.min(g.c2 + 1, g.c1 + 500), null, false); markDirty(); break;
      case 'formatCells': if (!pv) setDialog({ kind: 'format', tab: (arg as 'number') ?? 'number' }); break;
      case 'switchWindows': {
        const el = arg as HTMLElement;
        const list = win.windows?.().filter(w => w.appId === 'excel') ?? [];
        const r = el.getBoundingClientRect();
        menus.open({ x: r.left, y: r.bottom, variant: 'classic', items: list.map((w, i) => ({ label: `${i + 1} ${w.title.replace(new RegExp(` - ${BRAND.excel}$`), '')}`, checked: w.id === win.id, onClick: () => win.focusWindow?.(w.id) })) });
        break;
      }
    }
  };
  const siRef = useRef(si);
  siRef.current = si;

  const autofit = (c: number) => {
    if (!wb || !sheet) return;
    let w = 0;
    for (let r = 0; r < sheet.maxRow; r++) {
      const d = wb.display(si, r, c);
      if (!d.text) continue;
      const st = wb.style(wb.xfAt(si, r, c));
      const css = `${st.font.i ? 'italic ' : ''}${st.font.b ? '700 ' : ''}${(st.font.sz || 11) * 96 / 72}px ${cssFamily(st.font.name)}`;
      w = Math.max(w, textWidth(d.text, css));
    }
    wb.setColWidth(si, c, c, w ? Math.round(((w + 7) / 7) * 100) / 100 + 0.71 : sheet.defaultColWidth);
    markDirty();
  };

  const deleteSheet = (i: number) => {
    if (!wb || pv) return;
    if (wb.sheets.filter(s => s.state === 'visible').length <= 1) { void win.dialog.message({ title: MS_EXCEL, text: 'A workbook must contain at least one visible worksheet.', icon: 'warning' }); return; }
    const go = () => { wb.deleteSheet(i); setSi(Math.min(i, wb.sheets.length - 1)); undoStack.current = []; markDirty(); };
    if (wb.sheets[i].cells.size) void win.dialog.message({ title: MS_EXCEL, text: `${MS_EXCEL} will permanently delete this sheet. Do you want to continue?`, icon: 'warning', buttons: ['Delete', 'Cancel'] }).then(b => { if (b === 'Delete') go(); });
    else go();
  };

  // ---------------------------------------------------------------- save
  const doSave = async (saveAs: boolean): Promise<boolean> => {
    if (!wb) return false;
    if (pv) { return false; }
    if (editor) commit('none');
    let target = path;
    let kind: 'save' | 'save_as' = 'save';
    if (!target || saveAs || args.url) {
      const res = await (win.dialog.saveFileEx?.({ title: 'Save As', defaultName: `${stemOf(name)}.xlsx`, filters: [{ label: `${BRAND.excel} Workbook (*.xlsx)`, ext: ['xlsx'] }], officeFields: true, authors: persona?.name, controlId: ID.save })
        ?? win.dialog.saveFile({ title: 'Save As', defaultName: `${stemOf(name)}.xlsx`, filters: [{ label: `${BRAND.excel} Workbook (*.xlsx)`, ext: ['xlsx'] }] }).then(p => (p ? { path: p } : null)));
      if (!res) return false;
      target = res.path.toLowerCase().endsWith('.xlsx') ? res.path : `${res.path}.xlsx`;
      kind = 'save_as';
    }
    const bytes = wb.save();
    try {
      if (bytes.length > MAX_WORKBOOK_BYTES) throw new Error('too large');
      await api.saveFile(target!, bytes, kind);
    } catch (e) {
      await win.dialog.message({ title: MS_EXCEL, text: e instanceof Error && e.message === 'too large' ? `Your changes couldn't be saved to '${baseName(target!)}': the workbook is larger than 1.5 MB.` : `Your changes couldn't be saved to '${baseName(target!)}'.`, icon: 'error' });
      return false;
    }
    // continue from the saved package so later saves keep these bytes as the original
    const next = Workbook.load(bytes);
    next.nowMs = now();
    next.activeTab = si;
    setWb(next);
    setPath(target);
    setName(baseName(target!));
    setDirty(false);
    setSavedAt(now());
    undoStack.current = []; redoStack.current = [];
    logAppEvent('excel', 'save', { path: target, kind, bytes: bytes.length });
    return true;
  };

  // the page host (Download, Close) reads the current state through this handle
  const hostRef = useRef<HostHandle>({ save: () => Promise.resolve(false), dirty: false, path: undefined, name: '' });
  hostRef.current = { save: () => doSave(false), dirty, path, name };
  useEffect(() => win.bindHost?.(hostRef), [win]);

  useEffect(() => win.onBeforeClose?.(async () => {
    if (!dirty || pv) return true;
    const b = await win.dialog.message({ title: MS_EXCEL, text: `Want to save your changes to '${name}'?`, icon: 'none' as 'info', buttons: ['Save', "Don't Save", 'Cancel'] });
    if (b === 'Cancel') return false;
    if (b === 'Save') return doSave(false);
    return true;
  }), [dirty, pv, name, path, wb]); // eslint-disable-line react-hooks/exhaustive-deps

  const enableEditing = () => {
    setPv(false); setFloatTab(null);
    logAppEvent('excel', 'enable_editing', { path: path ?? name });
    rootRef.current?.focus({ preventScroll: true });
  };

  // ---------------------------------------------------------------- keyboard
  const onKeyDown = (e: React.KeyboardEvent) => {
    if (!wb || !sheet) return;
    if (e.target instanceof HTMLInputElement || e.target instanceof HTMLSelectElement) return;
    if (e.target instanceof HTMLTextAreaElement) return; // editor handles its own keys
    const k = e.key;
    const ctrl = e.ctrlKey || e.metaKey;
    const a = sel.active;
    const handled = () => { e.preventDefault(); e.stopPropagation(); };
    if (backstage) { if (k === 'Escape') { handled(); setBackstage(null); } return; }
    // shortcuts that work in Protected View too
    if (ctrl && k.toLowerCase() === 'f') { handled(); act('find'); return; }
    if (ctrl && k.toLowerCase() === 'g' || k === 'F5') { handled(); act('goto'); return; }
    if (ctrl && k.toLowerCase() === 'c') { handled(); doCopy(false); return; }
    if (ctrl && k.toLowerCase() === 'a') { handled(); const reg = currentRegion(a); const isSingle = sel.range.r1 === sel.range.r2 && sel.range.c1 === sel.range.c2; if (isSingle && (reg.r1 !== reg.r2 || reg.c1 !== reg.c2)) setSel({ active: a, anchor: a, range: reg, kind: 'cells' }); else setSel({ active: a, anchor: a, range: { r1: 0, c1: 0, r2: MAX_ROW, c2: MAX_COL }, kind: 'all' }); return; }
    if (ctrl && (k === 'PageDown' || k === 'PageUp')) { handled(); const vis = wb.sheets.map((s, i) => ({ s, i })).filter(x => x.s.state === 'visible'); const pos = vis.findIndex(x => x.i === si); const nx = vis[Math.max(0, Math.min(vis.length - 1, pos + (k === 'PageDown' ? 1 : -1)))]; if (nx) switchSheet(nx.i); return; }
    if (k === 'Escape') { handled(); setAnts(null); excelClip = excelClip?.cut ? null : excelClip; painter.current = null; setFloatTab(null); return; }
    if (k === 'F12') { handled(); void doSave(true); return; }
    if (ctrl && k.toLowerCase() === 's') { handled(); void doSave(false); return; }
    if (ctrl && k === 'F1') { handled(); setCollapsed(c => !c); return; }
    if (ctrl && k === '`') { handled(); act('toggleFormulas'); return; }
    if (e.altKey && k === 'F8') { handled(); setDialog({ kind: 'macro' }); return; }   // Macro dialog: nothing to run (VBA off by policy)
    if (e.altKey && k === 'F11') { handled(); return; }                                // VBA editor: blocked, no effect
    // movement
    const move = (dr: number, dc: number) => {
      handled();
      const from = e.shiftKey ? (sel.kind === 'cells' ? lastEnd(sel) : a) : a;
      let to: Addr;
      if (ctrl) to = edge(from, dr, dc);
      else {
        const m = sheet.merges.find(mm => mm.r1 <= from.r && mm.r2 >= from.r && mm.c1 <= from.c && mm.c2 >= from.c);
        let r = (dr > 0 && m ? m.r2 : dr < 0 && m ? m.r1 : from.r) + dr;
        let c = (dc > 0 && m ? m.c2 : dc < 0 && m ? m.c1 : from.c) + dc;
        // skip hidden rows/columns
        while (dr && r > 0 && r < MAX_ROW && wb.rowHidden(si, r)) r += dr;
        while (dc && c > 0 && c < MAX_COL && wb.colHidden(si, c)) c += dc;
        to = { r: Math.max(0, Math.min(MAX_ROW, r)), c: Math.max(0, Math.min(MAX_COL, c)) };
      }
      if (e.shiftKey) { const g = { r1: Math.min(sel.anchor.r, to.r), c1: Math.min(sel.anchor.c, to.c), r2: Math.max(sel.anchor.r, to.r), c2: Math.max(sel.anchor.c, to.c) }; selsRef.current.set(si, { active: a, anchor: sel.anchor, range: g, kind: 'cells' }); endRef.current = to; ensureVisible(to.r, to.c); bump(); }
      else { endRef.current = null; select(to); }
    };
    switch (k) {
      case 'ArrowDown': move(1, 0); return;
      case 'ArrowUp': move(-1, 0); return;
      case 'ArrowRight': move(0, 1); return;
      case 'ArrowLeft': move(0, -1); return;
      case 'PageDown': handled(); { const n = viewport.current.rows; setScroll(scroll.top + n, scroll.left); select({ r: Math.min(MAX_ROW, a.r + n), c: a.c }); } return;
      case 'PageUp': handled(); { const n = viewport.current.rows; setScroll(Math.max(0, scroll.top - n), scroll.left); select({ r: Math.max(0, a.r - n), c: a.c }); } return;
      case 'Home': handled(); if (ctrl) { const fz = sheet.freeze ?? { rows: 0, cols: 0 }; setScroll(fz.rows, fz.cols); select({ r: fz.rows, c: fz.cols }); } else select({ r: a.r, c: sheet.freeze?.cols ?? 0 }); return;
      case 'End': if (ctrl) { handled(); select({ r: Math.max(0, sheet.maxRow - 1), c: Math.max(0, sheet.maxCol - 1) }); } return;
      case 'Tab': handled(); select({ r: a.r, c: Math.max(0, a.c + (e.shiftKey ? -1 : 1)) }); return;
      case 'Enter': handled(); {
        const g = sel.range;
        if (g.r1 !== g.r2 || g.c1 !== g.c2) { let { r, c } = a; if (e.shiftKey) { r--; if (r < g.r1) { r = g.r2; c = c - 1 < g.c1 ? g.c2 : c - 1; } } else { r++; if (r > g.r2) { r = g.r1; c = c + 1 > g.c2 ? g.c1 : c + 1; } } setSel({ ...sel, active: { r, c } }); }
        else select({ r: Math.max(0, a.r + (e.shiftKey ? -1 : 1)), c: a.c });
      } return;
      case ' ': if (ctrl) { handled(); setSel({ active: a, anchor: a, range: { r1: 0, c1: sel.range.c1, r2: MAX_ROW, c2: sel.range.c2 }, kind: 'cols' }); return; } if (e.shiftKey) { handled(); setSel({ active: a, anchor: a, range: { r1: sel.range.r1, c1: 0, r2: sel.range.r2, c2: MAX_COL }, kind: 'rows' }); return; } break;
    }
    if (pv) return; // everything below edits
    if (ctrl && k.toLowerCase() === 'x') { handled(); doCopy(true); return; }
    if (ctrl && k.toLowerCase() === 'v') { handled(); doPaste(); return; }
    if (ctrl && k.toLowerCase() === 'h') { handled(); act('replace'); return; }
    if (ctrl && k.toLowerCase() === 'z') { handled(); const u = undoStack.current.pop(); if (u) restore(u, redoStack.current); return; }
    if (ctrl && k.toLowerCase() === 'y') { handled(); const u = redoStack.current.pop(); if (u) restore(u, undoStack.current); return; }
    if (ctrl && k.toLowerCase() === 'b') { handled(); act('bold'); return; }
    if (ctrl && k.toLowerCase() === 'i') { handled(); act('italic'); return; }
    if (ctrl && k.toLowerCase() === 'u') { handled(); act('underline'); return; }
    if (ctrl && k === '1') { handled(); act('formatCells', 'number'); return; }
    if (ctrl && k.toLowerCase() === 'd') { handled(); act('fill', 'down'); return; }
    if (ctrl && k.toLowerCase() === 'r') { handled(); act('fill', 'right'); return; }
    if (ctrl && k === ';') { handled(); startEdit('enter', formatValue(Math.floor(serialNow()), 'm/d/yyyy').text); return; }
    if (e.altKey && k === '=') { handled(); act('autosum', 'SUM'); return; }
    if (k === 'F2') { handled(); startEdit('edit'); return; }
    if (k === 'Delete') {
      handled();
      const targets = rangeCells(usedClip(sel.range)).filter(t => nonEmpty(t.r, t.c));
      if (!targets.length && !grouped) return;
      if (blocked('edit')) return;
      recordUndo(targets);
      for (const gi of targetSheets()) for (const t of (gi === si ? targets : rangeCells(usedClip(sel.range)))) { const cell = wb.cell(gi, t.r, t.c); if (cell) wb.setCellRaw(gi, t.r, t.c, null, undefined, cell.s ?? 0); }
      markDirty();
      return;
    }
    if (k === 'Backspace') { handled(); startEdit('enter', ''); return; }
    if (!ctrl && !e.altKey && k.length === 1) { handled(); startEdit('enter', k); return; }
  };
  const endRef = useRef<Addr | null>(null);
  const lastEnd = (s: Sel): Addr => endRef.current ?? { r: s.range.r1 === s.anchor.r ? s.range.r2 : s.range.r1, c: s.range.c1 === s.anchor.c ? s.range.c2 : s.range.c1 };

  const onEditorKey = (e: React.KeyboardEvent<HTMLTextAreaElement | HTMLInputElement>) => {
    if (!editor) return;
    e.stopPropagation();
    const k = e.key;
    if (k === 'Enter' && e.altKey) {
      e.preventDefault();
      const t = e.currentTarget; const s = t.selectionStart ?? editor.text.length;
      const text = editor.text.slice(0, s) + '\n' + editor.text.slice(t.selectionEnd ?? s);
      setEditor({ ...editor, text, caret: s + 1 });
      requestAnimationFrame(() => { t.setSelectionRange(s + 1, s + 1); });
      return;
    }
    if (k === 'Enter') { e.preventDefault(); commit(e.shiftKey ? 'up' : 'down', e.ctrlKey); return; }
    if (k === 'Tab') { e.preventDefault(); commit(e.shiftKey ? 'left' : 'right'); return; }
    if (k === 'Escape') { e.preventDefault(); cancelEdit(); return; }
    if (k === 'F2') { e.preventDefault(); setEditor({ ...editor, mode: editor.mode === 'edit' ? 'enter' : 'edit' }); return; }
    if (editor.mode === 'enter' && ['ArrowDown', 'ArrowUp', 'ArrowLeft', 'ArrowRight'].includes(k) && !pointing) {
      e.preventDefault();
      commit(k === 'ArrowDown' ? 'down' : k === 'ArrowUp' ? 'up' : k === 'ArrowLeft' ? 'left' : 'right');
      return;
    }
    if (pointing && editor.mode === 'enter' && ['ArrowDown', 'ArrowUp', 'ArrowLeft', 'ArrowRight'].includes(k)) {
      // arrow keys in point mode walk a reference from the edited cell
      e.preventDefault();
      const cur = lastPoint.current ? parseAddr(lastPoint.current) : null;
      const base = cur ?? { r: editor.r, c: editor.c };
      const d = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k as 'ArrowDown']!;
      const to = { r: Math.max(0, base.r + d[0]), c: Math.max(0, base.c + d[1]) };
      onPointRef(addrName(to.r, to.c), !!cur);
      return;
    }
    if ((e.ctrlKey || e.metaKey) && k.toLowerCase() === 'a') return; // select all text
    if (k === 'F4') {
      // cycle absolute/relative on the reference before the caret
      e.preventDefault();
      const caret = e.currentTarget.selectionStart ?? editor.text.length;
      const before = editor.text.slice(0, caret);
      const m = /(\$?)([A-Z]{1,3})(\$?)(\d+)$/.exec(before);
      if (m) {
        const state = (m[1] ? 2 : 0) + (m[3] ? 1 : 0);
        const next = [[true, true], [false, true], [true, false], [false, false]][state === 3 ? 0 : state === 1 ? 1 : state === 2 ? 2 : 3];
        // Excel order: A1 -> $A$1 -> A$1 -> $A1 -> A1
        const order = ['', '$$', '_$', '$_'];
        const curKey = `${m[1] ? '$' : '_'}${m[3] ? '$' : '_'}`.replace('__', '');
        const seq = ['', '$$', '_$', '$_'];
        const idx = seq.indexOf(curKey === '' ? '' : curKey);
        const nk = order[(idx + 1) % 4];
        void next;
        const rep = `${nk[0] === '$' ? '$' : ''}${m[2]}${nk[1] === '$' ? '$' : ''}${m[4]}`;
        const text = before.slice(0, before.length - m[0].length) + rep + editor.text.slice(caret);
        setEditor({ ...editor, text, caret: caret - m[0].length + rep.length });
      }
    }
  };

  const switchSheet = (i: number) => {
    if (editor) commit('none');
    wb?.setActiveTab(i);
    setSi(i);
    setAnts(a => a);
  };

  // ---------------------------------------------------------------- derived UI state
  const ribbonState: RibbonState = useMemo(() => ({
    pv, bold: !!activeStyle?.font.b, italic: !!activeStyle?.font.i, underline: !!activeStyle?.font.u,
    fontName: activeStyle?.font.name || 'Calibri', fontSize: activeStyle?.font.sz || 11, hAlign: activeStyle?.align.h, vAlign: activeStyle?.align.v,
    wrap: !!activeStyle?.align.wrap, merged: !!sheet?.merges.some(m => m.r1 === sel.active.r && m.c1 === sel.active.c),
    numFmtLabel: fmtLabel(activeStyle?.numFmt ?? 'General'), showGrid, showHeadings, showFormulaBar: showFbar, showFormulas,
    filterOn: filters.has(si), frozen: !!sheet?.freeze, zoom, sheetProtected: !!sheet?.protection,
  }), [pv, activeStyle, sheet, sel.active.r, sel.active.c, showGrid, showHeadings, showFbar, showFormulas, filters, si, zoom, version]); // eslint-disable-line react-hooks/exhaustive-deps

  const aggregates = useMemo(() => {
    if (!wb || !sheet) return null;
    const g = usedClip(sel.range);
    let count = 0, n = 0, sum = 0, cells = 0;
    let firstFmt: string | null = null;
    outer: for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
      if (++cells > 200000) break outer;
      const v = wb.value(si, r, c);
      if (v === null || v === '') continue;
      count++;
      if (typeof v === 'number') { n++; sum += v; if (firstFmt === null) firstFmt = wb.style(wb.xfAt(si, r, c)).numFmt; }
    }
    if (count < 2) return null;
    const f = (x: number) => formatValue(x, firstFmt && firstFmt !== 'General' ? firstFmt : 'General').text;
    return { count, avg: n ? f(sum / n) : null, sum: n ? f(sum) : null };
  }, [wb, sheet, sel.range, si, version]); // eslint-disable-line react-hooks/exhaustive-deps

  const activeLocked = !!sheet?.protection && lockedIn({ r1: sel.active.r, c1: sel.active.c, r2: sel.active.r, c2: sel.active.c });
  const formulaText = editor ? editor.text : wb ? wb.editText(si, sel.active.r, sel.active.c) : '';
  const nameBox = nameBoxText ?? addrName(sel.active.r, sel.active.c);

  // ---------------------------------------------------------------- sheet tab menu
  const visibleSheets = () => (wb ? wb.sheets.map((s, i) => (s.state === 'visible' ? i : -1)).filter(i => i >= 0) : []);
  /** per-sheet UI state (selection, scroll, filters) follows its sheet after Move or Copy; undo does not survive it */
  const reindex = (before: Sheet[]) => {
    if (!wb) return;
    const remap = <T,>(m: Map<number, T>) => { const out = new Map<number, T>(); wb.sheets.forEach((s, n) => { const o = before.indexOf(s); if (o >= 0 && m.has(o)) out.set(n, m.get(o)!); }); return out; };
    selsRef.current = remap(selsRef.current); scrollRef.current = remap(scrollRef.current); setFilters(f => remap(f));
    undoStack.current = []; redoStack.current = []; setAnts(null); excelClip = null;
  };
  /** Move or Copy OK: `pos` indexes the Before sheet list (visible sheets, then (move to end)) */
  const moveOrCopy = (from: number, pos: number, copy: boolean) => {
    if (!wb || pv) return;
    if (editor) commit('none');
    const before = wb.sheets.slice();
    const vis = visibleSheets();
    const targetSheet = pos < vis.length ? before[vis[pos]] : null;
    const sources = (grouped ? vis : [from]).map(i => before[i]);
    let active = before[from];
    if (copy) {
      let at = targetSheet ? wb.sheets.indexOf(targetSheet) : wb.sheets.length;
      sources.forEach((src, n) => { const ni = wb.copySheet(wb.sheets.indexOf(src), at); if (n === 0) active = wb.sheets[ni]; at = ni + 1; });
      setGrouped(false);
    } else {
      for (const src of sources) { if (src === targetSheet) continue; wb.moveSheet(wb.sheets.indexOf(src), targetSheet ? wb.sheets.indexOf(targetSheet) : wb.sheets.length); }
    }
    reindex(before);
    const ai = wb.sheets.indexOf(active);
    wb.setActiveTab(ai); setSi(ai);
    logAppEvent('excel', copy ? 'copy_sheet' : 'move_sheet', { sheet: before[from].name, before: targetSheet?.name ?? '(move to end)' });
    markDirty();
  };
  /** Protect Sheet... / Unprotect Sheet... */
  const protectSheet = (i: number) => {
    if (!wb || pv) return;
    const p = wb.sheets[i].protection;
    if (!p) { setDialog({ kind: 'protect', si: i }); return; }
    if (hasPassword(p)) { setDialog({ kind: 'unprotect', si: i }); return; }
    wb.setProtection(i, null); undoStack.current = []; redoStack.current = [];
    logAppEvent('excel', 'unprotect_sheet', { sheet: wb.sheets[i].name });
    markDirty();
  };
  /** Link to this Sheet: a hyperlink to the workbook at this sheet's A1 goes to the clipboard */
  const copySheetLink = (i: number) => {
    if (!wb) return;
    const sn = wb.sheets[i].name;
    const loc = `${/^[A-Za-z_][\w.]*$/.test(sn) ? sn : `'${sn.replace(/'/g, "''")}'`}!A1`;
    const url = `${path ? `${location.origin}/?file=${encodeURIComponent(path)}` : ''}#${encodeURI(loc)}`;
    const html = `<a href="${escHtml(url)}">${escHtml(name)}</a>`;
    window.dispatchEvent(new CustomEvent('sheet:clipboard-write', { detail: { text: url, html, source: 'excel' } }));
    logAppEvent('excel', 'link_to_sheet', { sheet: sn, url });
  };
  const tabMenu = (i: number, x: number, y: number) => {
    if (!wb) return;
    const hidden = wb.sheets.filter(s => s.state !== 'visible');
    const colors: [string, string | undefined][] = [['No Color', undefined], ['Red', '#C00000'], ['Orange', '#ED7D31'], ['Yellow', '#FFC000'], ['Green', '#70AD47'], ['Blue', '#4472C4'], ['Purple', '#7030A0']];
    const protectedNow = !!wb.sheets[i].protection;
    const lastSheet = () => void win.dialog.message({ title: MS_EXCEL, text: LAST_SHEET_MSG, icon: 'warning' });
    const items: XlMenuItem[] = [
      { label: 'Insert...', key: 'I', disabled: pv, onClick: () => { const n = wb.addSheet(i); setGrouped(false); setSi(n); markDirty(); } },
      { label: 'Delete', key: 'D', icon: 'menuDeleteSheet', disabled: pv, onClick: () => (grouped ? lastSheet() : deleteSheet(i)) },
      { label: 'Rename', key: 'R', icon: 'menuRenameSheet', disabled: pv || grouped, onClick: () => { switchSheet(i); setRenaming(i); } },
      { label: 'Move or Copy...', key: 'M', disabled: pv, onClick: () => setDialog({ kind: 'moveCopy', si: i }) },
      { label: 'View Code', key: 'V', icon: 'menuViewCode', onClick: () => setDialog({ kind: 'macro' }) },
      { label: protectedNow ? 'Unprotect Sheet...' : 'Protect Sheet...', key: 'P', icon: 'menuProtectSheet', disabled: pv || grouped, onClick: () => protectSheet(i) },
      { label: 'Tab Color', key: 'T', disabled: pv, submenu: colors.map(([l, c]) => ({ label: l, swatch: c ?? '#ffffff', onClick: () => { wb.sheets[i].tabColor = c; bump(); } })) },
      { separator: true },
      { label: 'Hide', key: 'H', disabled: pv || (!grouped && wb.sheets.filter(s => s.state === 'visible').length <= 1), onClick: () => { if (grouped) { lastSheet(); return; } wb.sheets[i].state = 'hidden'; const v = wb.sheets.findIndex(s => s.state === 'visible'); setSi(v); markDirty(); } },
      { label: 'Unhide...', key: 'U', disabled: pv || !hidden.length, onClick: () => { for (const s of hidden) s.state = 'visible'; setGrouped(false); markDirty(); } },
      { separator: true },
      grouped
        ? { label: 'Ungroup Sheets', key: 'N', onClick: () => setGrouped(false) }
        : { label: 'Select All Sheets', key: 'S', onClick: () => { if (visibleSheets().length > 1) setGrouped(true); } },
      { separator: true },
      { label: 'Link to this Sheet', key: 'L', icon: 'menuLinkSheet', onClick: () => copySheetLink(i) },
    ];
    xlMenu.open({ x: x + 1, y: y + 2, items, controlId: ID.tabs, above: true });
  };

  const cellMenu = (kind: 'cell' | 'col' | 'row' | 'all', x: number, y: number) => {
    const items: MenuItem[] = [
      { label: 'Cut', shortcut: 'Ctrl+X', disabled: pv, controlId: ID.copy, onClick: () => doCopy(true) },
      { label: 'Copy', shortcut: 'Ctrl+C', controlId: ID.copy, onClick: () => doCopy(false) },
      { label: 'Paste Options:', disabled: pv || !clipboard.read()?.text, controlId: ID.copy, onClick: () => doPaste() },
      { label: 'Paste Special...', disabled: true },
      { separator: true },
      { label: 'Smart Lookup', disabled: true },
      { separator: true },
      ...(kind === 'col' ? [{ label: 'Insert', disabled: pv, onClick: () => act('insertCols') }, { label: 'Delete', disabled: pv, onClick: () => act('deleteCols') }] as MenuItem[]
        : kind === 'row' ? [{ label: 'Insert', disabled: pv, onClick: () => act('insertRows') }, { label: 'Delete', disabled: pv, onClick: () => act('deleteRows') }] as MenuItem[]
          : [{ label: 'Insert...', disabled: pv, submenu: [{ label: 'Entire row', onClick: () => act('insertRows') }, { label: 'Entire column', onClick: () => act('insertCols') }] }, { label: 'Delete...', disabled: pv, submenu: [{ label: 'Entire row', onClick: () => act('deleteRows') }, { label: 'Entire column', onClick: () => act('deleteCols') }] }] as MenuItem[]),
      { label: 'Clear Contents', disabled: pv, onClick: () => act('clear', 'contents') },
      { separator: true },
      ...(kind === 'cell' || kind === 'all' ? [{ label: 'Filter', disabled: pv, submenu: [{ label: filters.has(si) ? 'Clear Filter' : 'Filter by Selected Cell\'s Value', onClick: () => act('filter') }] }, { label: 'Sort', disabled: pv, submenu: [{ label: 'Sort A to Z', onClick: () => act('sort', 1) }, { label: 'Sort Z to A', onClick: () => act('sort', -1) }] }, { separator: true }] as MenuItem[] : []),
      { label: 'New Comment', disabled: true },
      { label: 'New Note', disabled: true },
      { label: 'Format Cells...', disabled: pv, onClick: () => act('formatCells', 'number') },
      ...(kind === 'col' ? [{ label: 'Column Width...', disabled: pv, onClick: () => act('colWidthDialog') }, { label: 'Hide', disabled: pv, onClick: () => act('hideCols') }, { label: 'Unhide', disabled: pv, onClick: () => act('unhideCols') }] as MenuItem[] : []),
      ...(kind === 'row' ? [{ label: 'Row Height...', disabled: pv, onClick: () => act('rowHeightDialog') }, { label: 'Hide', disabled: pv, onClick: () => act('hideRows') }, { label: 'Unhide', disabled: pv, onClick: () => act('unhideRows') }] as MenuItem[] : []),
      ...(kind === 'cell' ? [{ label: 'Pick From Drop-down List...', disabled: true }, { label: 'Define Name...', disabled: true }, { label: 'Link', disabled: true }] as MenuItem[] : []),
    ];
    menus.open({ x, y, items, controlId: 'excel.grid_cells', variant: 'classic', minWidth: 200 });
  };

  const filterMenu = (c: number, el: HTMLElement) => {
    const fr = filters.get(si);
    if (!wb || !fr) return;
    const values = new Map<string, number[]>();
    for (let r = fr.r1 + 1; r <= fr.r2; r++) { const t = wb.display(si, r, c).text || '(Blanks)'; (values.get(t) ?? values.set(t, []).get(t)!).push(r); }
    const r0 = el.getBoundingClientRect();
    const sorted = [...values.keys()].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    menus.open({
      x: r0.left, y: r0.bottom, variant: 'classic', controlId: 'excel.data_sort_and_filter', minWidth: 220, items: [
        { label: 'Sort A to Z', onClick: () => { setSel({ ...sel, active: { r: fr.r1 + 1, c }, anchor: { r: fr.r1 + 1, c }, range: { r1: fr.r1 + 1, c1: c, r2: fr.r1 + 1, c2: c } }); setTimeout(() => act('sort', 1), 0); } },
        { label: 'Sort Z to A', onClick: () => { setSel({ ...sel, active: { r: fr.r1 + 1, c }, anchor: { r: fr.r1 + 1, c }, range: { r1: fr.r1 + 1, c1: c, r2: fr.r1 + 1, c2: c } }); setTimeout(() => act('sort', -1), 0); } },
        { separator: true },
        { label: `Clear Filter From "${wb.display(si, fr.r1, c).text}"`, onClick: () => { wb.setRowHeight(si, fr.r1 + 1, fr.r2, null, false); markDirty(); } },
        { separator: true },
        ...sorted.slice(0, 40).map(v => ({ label: v, checked: values.get(v)!.some(r => !wb.rowHidden(si, r)), keepOpen: false, onClick: () => {
          const rowsFor = values.get(v)!;
          const anyVisible = rowsFor.some(r => !wb.rowHidden(si, r));
          for (const r of rowsFor) wb.setRowHeight(si, r, r, null, anyVisible);
          markDirty();
        } })),
      ],
    });
  };

  // ---------------------------------------------------------------- render
  const docTitle = pv ? `${stemOf(name)}  -  Protected View` : stemOf(name);
  const docStatus = !dirty && (path || savedAt) ? '\u2002\u2022\u2002Saved' : '';
  if (err) return <div className="xl-root"><TitleBar title={BRAND.excel} status="" pv={false} persona={persona} onSave={() => {}} focused={focused} /><div className="xl-loading">{err}</div></div>;
  if (!wb || !sheet) return <div className="xl-root"><TitleBar title={docTitle} status="" pv={pv} persona={persona} onSave={() => {}} focused={focused} /><div className="xl-loading">Opening - {BRAND.excel}</div></div>;

  const fr = filters.get(si) ?? null;
  // sheet tabs | horizontal scroll bar split: Excel's default tab ratio is 0.6 (0.6 of the window width, offset 13 px)
  const tabSplit = Math.round(rootW * 0.6 + 13);
  const quickAnalysis = (() => { const g = sel.range; if (g.r1 === g.r2 && g.c1 === g.c2) return false; let n = 0; for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 200); r++) for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 50); c++) if (nonEmpty(r, c) && ++n > 1) return true; return false; })();
  const showRibbon = (!collapsed || floatTab) && !backstage;
  return (
    <div ref={rootRef} className={`xl-root${focused && !(dlgFocus && dialog) ? '' : ' inactive'}`} tabIndex={-1} onKeyDown={onKeyDown} data-autofocus data-book={name}>
      <TitleBar title={docTitle} tag={grouped ? '[Group]' : undefined} status={docStatus} pv={pv} persona={persona} onSave={() => void doSave(false)} focused={focused}
        onUndo={() => { const u = undoStack.current.pop(); if (u) restore(u, redoStack.current); }} onRedo={() => { const u = redoStack.current.pop(); if (u) restore(u, undoStack.current); }}
        canUndo={!pv && undoStack.current.length > 0} canRedo={!pv && redoStack.current.length > 0} />
      <div className="xl-tabs" role="tablist">
        {TABS.map(t => (
          <button key={t} className={`xl-tab${t === tab && (!collapsed || floatTab === t) && t !== 'File' ? ' active' : ''}`} role="tab" style={t === 'File' ? { width: 48 } : { padding: '0 10.8px 2px' }} data-control-id={t === 'Home' ? 'excel.home_tab_number_format_font_alignment_borders' : t === 'Formulas' ? 'excel.formulas_autosum_function_library' : t === 'Data' ? 'excel.data_sort_and_filter' : t === 'Insert' ? 'excel.insert_charts_pivottable_power_query_copilot' : t === 'File' ? ID.save : undefined}
            data-control-item={`${t} tab`} onMouseDown={e => e.preventDefault()}
            onClick={() => { if (t === 'File') { setBackstage('info'); return; } setTab(t); if (collapsed) setFloatTab(f => (f === t ? null : t)); }}
            onDoubleClick={() => { if (t !== 'File' && !pv) { setCollapsed(c => !c); setFloatTab(null); } }}>{t}</button>
        ))}
        <div className="xl-tabs-right">
          <button className="xl-pill" title="Comments" disabled><Icon name="commentPill" size={16} />Comments</button>
          <button className="xl-pill primary" title="Share" disabled><Icon name="sharePill" size={16} />Share<svg width="5" height="3" viewBox="0 0 5 3" shapeRendering="crispEdges" fill="currentColor" style={{ marginLeft: 6, flex: 'none' }} aria-hidden><rect x="0" y="0" width="1" height="1" /><rect x="1" y="1" width="1" height="1" /><rect x="2" y="2" width="1" height="1" /><rect x="3" y="1" width="1" height="1" /><rect x="4" y="0" width="1" height="1" /></svg></button>
        </div>
      </div>
      {showRibbon && (
        <div style={{ position: 'relative', height: collapsed ? 0 : undefined, zIndex: 30 }}>
          <Ribbon tab={floatTab ?? tab} st={ribbonState} act={act} floating={!!floatTab && collapsed} onCollapse={pv ? undefined : () => setCollapsed(true)} />
        </div>
      )}
      {pv && (
        <div className="xl-pvbar" data-control-id={ID.pv} role="alert" aria-label="Protected View">
          <Icon name="pvInfo" size={18} />
          <span className="xl-pv-title">PROTECTED VIEW</span>
          <span className="xl-pv-msg">{pvReason === 'attachment' ? 'Be careful—email attachments can contain viruses. Unless you need to edit, it\'s safer to stay in Protected View.' : 'Be careful—files from the Internet can contain viruses. Unless you need to edit, it\'s safer to stay in Protected View.'}</span>
          <button className="xl-pv-btn" data-control-id={ID.pv} data-control-item="Enable Editing" onClick={enableEditing}>Enable Editing</button>
          <button className="xl-pv-close" aria-label="Close" data-control-id={ID.pv} data-control-item="Close" onClick={() => {}} title="Close"><Icon name="close" size={10} /></button>
        </div>
      )}
      {showFbar && (
        <div className="xl-fbar">
          <div className="xl-fbox xl-namebox" data-control-id={ID.name}>
            <input value={nameBox} spellCheck={false} aria-label="Name Box" data-control-id={ID.name}
              onFocus={e => { const el = e.target; setNameBoxText(el.value); requestAnimationFrame(() => { if (document.activeElement === el) el.select(); }); }}
              onChange={e => setNameBoxText(e.target.value)}
              onBlur={() => setNameBoxText(null)}
              onKeyDown={e => {
                e.stopPropagation();
                if (e.key === 'Enter') { e.preventDefault(); const t = (nameBoxText ?? '').trim(); setNameBoxText(null); if (t) goTo(t); rootRef.current?.focus({ preventScroll: true }); }
                if (e.key === 'Escape') { e.preventDefault(); setNameBoxText(null); rootRef.current?.focus({ preventScroll: true }); }
              }} />
            <button aria-label="Name Box list" onMouseDown={e => e.preventDefault()} onClick={e => { const r = e.currentTarget.getBoundingClientRect(); menus.open({ x: r.left - 80, y: r.bottom, variant: 'classic', minWidth: 96, controlId: ID.name, items: [...wb.definedNames.keys()].length ? [...wb.definedNames.keys()].map(n => ({ label: n, onClick: () => goTo(n) })) : [{ label: '(no names)', disabled: true }] }); }}><Icon name="chevronV" size={12} style={{ transform: 'translate(0.5px, 1.5px)' }} /></button>
          </div>
          <div className="xl-fbar-grip"><Icon name="fbarGrip" size={16} /></div>
          <div className="xl-fbox xl-fbar-btns">
            <button className="xl-fx-cancel" aria-label="Cancel" title="Cancel" disabled={!editor} onClick={cancelEdit}><Icon name="fxCancel" size={16} /></button>
            <button className="xl-fx-enter" aria-label="Enter" title="Enter" disabled={!editor} onClick={() => commit('none')}><Icon name="fxEnter" size={16} /></button>
            <button className="xl-fx-fn" aria-label="Insert Function" title="Insert Function" disabled={pv} data-control-id="excel.formulas_autosum_function_library" data-control-item="fx" onClick={() => act('insertFunction', '')}><Icon name="fx" size={18} /></button>
            <button className="xl-fx-more" aria-label="More" tabIndex={-1} onMouseDown={e => e.preventDefault()}><Icon name="chevronV" size={12} style={{ transform: 'translate(-1px, 1px)' }} /></button>
          </div>
          <div className="xl-fx-gap" />
          <div className="xl-fbox xl-fbar-input">
            <input ref={fbarRef} className={`xl-formula${formulaText.startsWith('=') ? ' mono' : ''}`} value={formulaText} spellCheck={false} readOnly={pv || (!editor && activeLocked)} data-control-id={ID.fbar} aria-label="Formula Bar"
              onFocus={() => { if (!editor && !pv && !activeLocked) startEdit('edit'); }}
              onChange={e => editor && setEditor({ ...editor, text: e.target.value, caret: e.target.selectionStart ?? e.target.value.length, mode: 'edit' })}
              onKeyDown={e => { if (editor) onEditorKey(e); else e.stopPropagation(); }} />
            <div className="xl-fbar-expand" title="Expand Formula Bar (Ctrl+Shift+U)"><Icon name="chevronV" size={12} style={{ transform: 'translate(2.5px, -1px)' }} /></div>
          </div>
        </div>
      )}
      <div className="xl-gridwrap">
        <div className="xl-gridarea">
          <Grid wb={wb} si={si} zoom={zoom} version={version} top={scroll.top} left={scroll.left} setScroll={setScroll}
            sel={sel} onSelect={setSel} editor={editor}
            onEditorText={(text, caret) => { if (!editor) return; if (text !== editor.text) lastPoint.current = null; if (text !== editor.text || caret !== editor.caret) setEditor({ ...editor, text, caret }); }}
            onEditorKey={onEditorKey} editorRef={editorRef}
            showGrid={showGrid} showHeadings={showHeadings} showFormulas={showFormulas}
            ants={ants && ants.si === si ? ants.range : null} filter={fr} onFilterMenu={filterMenu}
            onDoubleClick={(r, c) => { if (pv || blocked('edit', { r1: r, c1: c, r2: r, c2: c })) return; selsRef.current.set(si, { active: { r, c }, anchor: { r, c }, range: { r1: r, c1: c, r2: r, c2: c }, kind: 'cells' }); const text = wb.editText(si, r, c); setEditor({ r, c, text, mode: 'edit', caret: text.length }); }}
            onContext={(kind, x, y) => cellMenu(kind, x, y)}
            onResizeCol={(c, px) => { if (pv || blocked('formatColumns')) return; const cs = sel.kind === 'cols' && c >= sel.range.c1 && c <= sel.range.c2 ? [sel.range.c1, sel.range.c2] : [c, c]; wb.setColWidth(si, cs[0], cs[1], Math.max(0, Math.round((px / 7) * 256) / 256), px < 1 ? true : undefined); markDirty(); }}
            onResizeRow={(r, px) => { if (pv || blocked('formatRows')) return; wb.setRowHeight(si, r, r, Math.round(px * 0.75 * 4) / 4, px < 1 ? true : undefined); markDirty(); }}
            onAutofitCol={c => !pv && !blocked('formatColumns') && autofit(c)}
            onPointRef={onPointRef} pointing={pointing}
            onViewport={(rows, cols, t) => { viewport.current = { rows, cols }; if (t) setThumbs(o => o && Math.abs(o.h.at - t.h.at) + Math.abs(o.h.frac - t.h.frac) + Math.abs(o.v.at - t.v.at) + Math.abs(o.v.frac - t.v.frac) < 1e-4 ? o : t); }}
            quickAnalysis={!pv && !editor && quickAnalysis} fab={false} />
          <Scrollbar vertical pos={scroll.top} max={(Math.max(sheet.maxRow, scroll.top + viewport.current.rows) + 1) / 0.97} page={viewport.current.rows} thumb={thumbs?.v} onChange={v => setScroll(v, scroll.left)} />
        </div>
        <div className="xl-sheetbar">
          <div className="xl-sheetnav">
            <button aria-label="Scroll to the previous sheet" onClick={() => { const vis = wb.sheets.map((s, i) => ({ s, i })).filter(x => x.s.state === 'visible'); const pos = vis.findIndex(x => x.i === si); if (pos > 0) switchSheet(vis[pos - 1].i); }}><Icon name="sheetPrev" size={16} /></button>
            <button aria-label="Scroll to the next sheet" onClick={() => { const vis = wb.sheets.map((s, i) => ({ s, i })).filter(x => x.s.state === 'visible'); const pos = vis.findIndex(x => x.i === si); if (pos < vis.length - 1) switchSheet(vis[pos + 1].i); }}><Icon name="sheetNext" size={16} /></button>
          </div>
          <div className="xl-sheettabs" data-control-id={ID.tabs} role="tablist">
            {wb.sheets.map((s, i) => s.state !== 'visible' ? null : (
              <div key={`${i}:${s.name}`} className={`xl-sheettab${i === si ? ' active' : grouped ? ' grouped' : ''}`} role="tab" aria-selected={i === si || grouped} data-control-id={ID.tabs} data-control-item={s.name}
                onMouseDown={e => { if (e.button === 0) { if (grouped && i !== si) setGrouped(false); switchSheet(i); } }}
                onDoubleClick={() => { if (!pv) setRenaming(i); }}
                onContextMenu={e => { e.preventDefault(); e.stopPropagation(); if (i !== si) switchSheet(i); tabMenu(i, e.clientX, e.clientY); }}>
                {renaming === i ? (
                  <input autoFocus defaultValue={s.name} maxLength={31} onFocus={e => e.target.select()} onMouseDown={e => e.stopPropagation()}
                    onKeyDown={e => {
                      e.stopPropagation();
                      if (e.key === 'Enter') {
                        const v = (e.target as HTMLInputElement).value.trim();
                        if (!v || /[\\/?*[\]:]/.test(v) || (wb.sheetIndex(v) >= 0 && wb.sheetIndex(v) !== i)) { void win.dialog.message({ title: MS_EXCEL, text: 'You typed an invalid name for a sheet or chart. Make sure that:\n\n• The name that you type does not exceed 31 characters.\n• The name does not contain any of the following characters: \\ / ? * [ or ]\n• You did not leave the name blank.', icon: 'warning' }); return; }
                        wb.renameSheet(i, v); setRenaming(null); markDirty(); rootRef.current?.focus({ preventScroll: true });
                      }
                      if (e.key === 'Escape') { setRenaming(null); rootRef.current?.focus({ preventScroll: true }); }
                    }}
                    onBlur={e => { const v = e.target.value.trim(); if (v && v !== s.name && !/[\\/?*[\]:]/.test(v) && wb.sheetIndex(v) < 0) { wb.renameSheet(i, v); markDirty(); } setRenaming(null); }} />
                ) : <span className="xl-sheettab-label">{s.name}</span>}
                {s.tabColor && <span className="xl-tabcolor" style={{ background: s.tabColor }} />}
              </div>
            ))}
          </div>
          <button className="xl-addsheet" aria-label="New sheet" title="New sheet" data-control-id={ID.tabs} data-control-item="New sheet" disabled={pv} onClick={() => { const n = wb.addSheet(si + 1); setGrouped(false); setSi(n); markDirty(); }}><Icon name="addSheet" size={16} /></button>
          <div className="xl-sheetbar-split" style={{ position: 'absolute', left: tabSplit - 20 }}><Icon name="splitGrip" size={16} /></div>
          <div style={{ position: 'absolute', left: tabSplit + 6, right: 20, top: 0 }}>
            <Scrollbar pos={scroll.left} max={(Math.max(sheet.maxCol, scroll.left + viewport.current.cols) + 1) / 0.97} page={viewport.current.cols} thumb={thumbs?.h} onChange={v => setScroll(scroll.top, v)} />
          </div>
        </div>
      </div>
      <div className="xl-status" data-control-id={ID.status}>
        <span className="xl-status-ready">{editor ? (editor.mode === 'edit' ? 'Edit' : pointing ? 'Point' : 'Enter') : ants && excelClip ? 'Select destination and press ENTER or choose Paste' : 'Ready'}</span>
        {!(ants && excelClip && !editor) && !pv && <span className="xl-status-access"><Icon name="accessStatus" size={16} />Accessibility: Good to go</span>}
        <div className="xl-status-right">
          {aggregates && (
            <div className="xl-status-agg" data-control-id={ID.status} data-control-item="Sum Count Average">
              {aggregates.avg !== null && <span>Average: {aggregates.avg}</span>}
              <span>Count: {aggregates.count}</span>
              {aggregates.sum !== null && <span>Sum: {aggregates.sum}</span>}
            </div>
          )}
          <div className="xl-views">
            <button className="on" title="Normal" aria-label="Normal"><Icon name="normalViewSb" size={16} /></button>
            <button title="Page Layout" aria-label="Page Layout" disabled><Icon name="pageLayoutViewSb" size={16} /></button>
            <button title="Page Break Preview" aria-label="Page Break Preview" disabled><Icon name="pageBreakViewSb" size={16} /></button>
          </div>
          <ZoomSlider zoom={zoom} onZoom={setZoom} onDialog={() => act('zoomDialog')} />
        </div>
      </div>
      {dialog && (dialog.kind === 'find' || dialog.kind === 'replace') && (
        <FindReplace mode={dialog.kind} setMode={m => setDialog({ kind: m })} canReplace={!pv} opts={findOpts} setOpts={o => setFindOptsState(s => ({ ...s, ...o }))}
          onFindNext={findNext} onFindAll={findAll} onReplace={doReplace} onReplaceAll={doReplaceAll} onClose={() => { setDialog(null); rootRef.current?.focus({ preventScroll: true }); }}
          onPick={goToHit} results={findResults} status={findStatus} onFocusChange={setDlgFocus} hidden={!focused && !dlgFocus} />
      )}
      {dialog?.kind === 'goto' && <GoTo names={[...wb.definedNames.keys()]} onGo={r => { if (goTo(r) !== false) { setDialog(null); rootRef.current?.focus({ preventScroll: true }); } }} onClose={() => { setDialog(null); rootRef.current?.focus({ preventScroll: true }); }} />}
      {dialog?.kind === 'colWidth' && <NumberPrompt title="Column Width" label="Column width:" value={String(Math.round((wb.colWidthChars(si, sel.active.c) - 0.71) * 100) / 100)} onClose={() => setDialog(null)} onOk={v => { wb.setColWidth(si, sel.range.c1, Math.min(sel.range.c2, sel.range.c1 + 500), v + 0.71); setDialog(null); markDirty(); }} />}
      {dialog?.kind === 'rowHeight' && <NumberPrompt title="Row Height" label="Row height:" value={String(wb.rowHeightPt(si, sel.active.r))} onClose={() => setDialog(null)} onOk={v => { wb.setRowHeight(si, sel.range.r1, Math.min(sel.range.r2, sel.range.r1 + 5000), v); setDialog(null); markDirty(); }} />}
      {dialog?.kind === 'zoom' && <ZoomDialog zoom={zoom} onClose={() => setDialog(null)} onOk={z => { setZoom(z); setDialog(null); }} />}
      {dialog?.kind === 'insertFn' && <InsertFunction names={FUNCTION_NAMES} onClose={() => setDialog(null)} onOk={fn => { setDialog(null); act('insertFunction', fn); }} />}
      {dialog?.kind === 'format' && <FormatCells tab={dialog.tab} sample={wb.display(si, sel.active.r, sel.active.c).text} onClose={() => setDialog(null)}
        onApply={pp => { setDialog(null); fmt({ numFmt: pp.numFmt, align: pp.h || pp.v || pp.wrap !== undefined ? { h: pp.h === 'general' ? undefined : pp.h, v: pp.v, wrap: pp.wrap } : undefined, font: pp.bold !== undefined || pp.italic !== undefined || pp.size ? { b: pp.bold, i: pp.italic, ...(pp.size ? { sz: pp.size } : {}) } : undefined }); }} />}
      {dialog?.kind === 'moveCopy' && (
        <MoveCopyDialog book={name} sheets={visibleSheets().map(i => wb.sheets[i].name)} current={Math.max(0, visibleSheets().indexOf(dialog.si))}
          onClose={() => { setDialog(null); rootRef.current?.focus({ preventScroll: true }); }}
          onOk={(pos, copy) => { const from = dialog.si; setDialog(null); moveOrCopy(from, pos, copy); rootRef.current?.focus({ preventScroll: true }); }} />
      )}
      {dialog?.kind === 'protect' && (
        <ProtectSheetDialog flags={PROTECTION_FLAGS} alert={text => win.dialog.message({ title: MS_EXCEL, text, icon: 'warning' })}
          onClose={() => { setDialog(null); rootRef.current?.focus({ preventScroll: true }); }}
          onOk={(pw, allow) => {
            const i = dialog.si;
            wb.setProtection(i, buildProtection(pw, allow));
            undoStack.current = []; redoStack.current = [];
            setDialog(null); markDirty(); rootRef.current?.focus({ preventScroll: true });
            logAppEvent('excel', 'protect_sheet', { sheet: wb.sheets[i].name, password: !!pw, allow: Object.keys(allow).filter(k => allow[k]) });
          }} />
      )}
      {dialog?.kind === 'unprotect' && (
        <UnprotectSheetDialog onClose={() => { setDialog(null); rootRef.current?.focus({ preventScroll: true }); }}
          onOk={pw => {
            const i = dialog.si, p = wb.sheets[i].protection;
            setDialog(null); rootRef.current?.focus({ preventScroll: true });
            if (p && !verifyPassword(p, pw)) { void win.dialog.message({ title: MS_EXCEL, text: 'The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.', icon: 'warning' }); return; }
            wb.setProtection(i, null); undoStack.current = []; redoStack.current = [];
            logAppEvent('excel', 'unprotect_sheet', { sheet: wb.sheets[i].name });
            markDirty();
          }} />
      )}
      {dialog?.kind === 'macro' && <MacroDialog onClose={() => { setDialog(null); rootRef.current?.focus({ preventScroll: true }); }} />}
      {backstage && <Backstage page={backstage} setPage={setBackstage} name={name} path={path} pv={pv} onClose={() => setBackstage(null)}
        onSave={async () => { setBackstage(null); await doSave(false); }} onSaveAs={async () => { setBackstage(null); await doSave(true); }}
        onCloseWindow={() => win.close()} onOpen={async () => { const o = { title: 'Open', filters: [{ label: `${BRAND.excel} Files`, ext: ['xlsx'] }], splitMenu: true }; const r = win.dialog.openFileEx ? await win.dialog.openFileEx(o) : await win.dialog.openFile(o).then(path => path ? { path, paths: [path], mode: 'open' as const } : null); setBackstage(null); if (r) win.open('excel', { file: r.path, args: r.mode === 'protectedView' ? { protectedView: true } : r.mode === 'open' ? undefined : { openMode: r.mode } }); }} persona={persona?.name ?? ''} />}
      <XlMenuHost />
    </div>
  );
}

// ------------------------------------------------------------------ pieces

function TitleBar({ title, tag, status, pv, persona, onSave, onUndo, onRedo, canUndo, canRedo, focused }: { title: string; tag?: string; status: string; pv: boolean; saved?: number | null; dirty?: boolean; persona: { name: string; initials: string } | undefined; onSave: () => void; onUndo?: () => void; onRedo?: () => void; canUndo?: boolean; canRedo?: boolean; focused: boolean }) {
  void focused;
  const barRef = useRef<HTMLDivElement>(null);
  const nameRef = useRef<HTMLButtonElement>(null);
  const [box, setBox] = useState<{ left: number; width: number; barW: number } | null>(null);
  const barW = box?.barW ?? 1440;
  const nameMax = pv ? 252 : status ? Math.round(barW * 0.117) : Math.round(barW * 0.1335);
  // Office's title layout: the search box is centred unless the document name reaches it, then it moves right; it
  // shrinks (down to 304 px) before the controls at the right, and disappears when even that does not fit
  useLayoutEffect(() => {
    const bar = barRef.current, nm = nameRef.current;
    if (!bar || !nm) return;
    const place = () => {
      const w = bar.clientWidth;
      const end = nm.offsetLeft + nm.offsetWidth;
      let width = Math.floor(Math.min(368, Math.max(304, w * 0.27)));
      let left = Math.max(Math.round((w - width) / 2 + 8), end + 27);
      const limit = w - 373;
      if (left + width > limit) width = limit - left;
      setBox(b => (b && b.left === left && b.width === width && b.barW === w ? b : { left, width, barW: w }));
    };
    place();
    const ro = new ResizeObserver(place);
    ro.observe(bar);
    return () => ro.disconnect();
  }, [title, tag, status, nameMax]);
  return (
    <div className="xl-titlebar" data-window-drag="true" ref={barRef}>
      <span className="xl-appicon"><ExcelAppIcon size={16} /></span>
      <div className="xl-qat">
        <button className="xl-autosave" aria-label="AutoSave" aria-pressed={false} onMouseDown={e => e.preventDefault()} title={`AutoSave is available when the file is stored on ${BRAND.onedrive} or ${BRAND.sharepoint}`}>
          <span>AutoSave</span>
          <span className="xl-toggle"><span className="xl-toggle-knob" /><span className="xl-toggle-text">Off</span></span>
        </button>
        <button className="xl-qat-save" data-control-id="excel.save_save_as_xlsx" data-control-item="Save" aria-label="Save" title="Save (Ctrl+S)" disabled={pv} onClick={onSave} onMouseDown={e => e.preventDefault()}><Icon name="saveQat" size={20} /></button>
        <button className="xl-qat-undo" aria-label="Undo" title="Undo (Ctrl+Z)" disabled={!canUndo} onClick={onUndo} onMouseDown={e => e.preventDefault()}><Icon name="undoQat" size={20} /></button>
        <button className="xl-qat-redo" aria-label="Redo" title="Redo (Ctrl+Y)" disabled={!canRedo} onClick={onRedo} onMouseDown={e => e.preventDefault()}><Icon name="redoQat" size={20} /></button>
        <button className="xl-qat-more" aria-label="Customize Quick Access Toolbar" onMouseDown={e => e.preventDefault()}><Icon name="qatMore" size={12} /></button>
      </div>
      <button className="xl-docname" aria-label={title} onMouseDown={e => e.preventDefault()} ref={nameRef}>
        <span className="xl-docname-text" style={{ maxWidth: nameMax }}>{title}</span>
        {tag && <span className="xl-docname-tag">{tag}</span>}
        {status && <span className="xl-docname-status">{status}</span>}
        <Icon name="chevronV" size={12} style={{ marginLeft: 2, transform: 'translate(-0.5px, 0.5px)' }} />
      </button>
      <div className="xl-search" style={{ left: box?.left ?? '50%', width: box?.width, visibility: !box || box.width < 304 ? 'hidden' : undefined }}>
        <Icon name="searchQat" size={16} />
        <input placeholder="Search" aria-label="Search (Alt+Q)" onKeyDown={e => e.stopPropagation()} />
      </div>
      <div className="xl-title-right">
        <button className="xl-account" aria-label={`Account: ${persona?.name ?? ''}`} title={persona?.name} onMouseDown={e => e.preventDefault()}>
          <span className="xl-avatar">{persona?.initials}</span>
        </button>
      </div>
    </div>
  );
}

function Scrollbar({ vertical, pos, max, page, thumb: px, onChange }: { vertical?: boolean; pos: number; max: number; page: number; thumb?: { at: number; frac: number }; onChange: (v: number) => void }) {
  const track = useRef<HTMLDivElement>(null);
  const frac = px ? px.frac : Math.min(1, page / Math.max(1, max));
  const at = px ? px.at : Math.min(1 - frac, pos / Math.max(1, max));
  const drag = (e: React.MouseEvent) => {
    e.preventDefault(); e.stopPropagation();
    const el = track.current!;
    const r = el.getBoundingClientRect();
    const len = vertical ? r.height : r.width;
    const start = vertical ? e.clientY : e.clientX;
    const p0 = pos;
    const move = (ev: MouseEvent) => { const d = ((vertical ? ev.clientY : ev.clientX) - start) / len * max; onChange(Math.max(0, Math.round(p0 + d))); };
    const up = () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
    window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
  };
  const thumb = vertical ? { top: `${at * 100}%`, height: `${Math.max(frac * 100, 4)}%` } : { left: `${at * 100}%`, width: `${Math.max(frac * 100, 4)}%` };
  const body = (
    <>
      <button className="xl-scroll-btn" aria-label={vertical ? 'Scroll up' : 'Scroll left'} onClick={() => onChange(Math.max(0, pos - 1))}><Icon name={vertical ? 'triUp' : 'triLeft'} size={16} /></button>
      <div className="xl-scroll-track" ref={track} onMouseDown={e => { const r = track.current!.getBoundingClientRect(); const click = vertical ? (e.clientY - r.top) / r.height : (e.clientX - r.left) / r.width; onChange(Math.max(0, pos + (click < at ? -page : page))); }}>
        <div className="xl-scroll-thumb" style={thumb} onMouseDown={drag} />
      </div>
      <button className="xl-scroll-btn" aria-label={vertical ? 'Scroll down' : 'Scroll right'} onClick={() => onChange(pos + 1)}><Icon name={vertical ? 'triDown' : 'triRight'} size={16} /></button>
    </>
  );
  if (vertical) return <div className="xl-vscroll" onMouseDown={e => e.stopPropagation()}><div className="xl-vscroll-inner">{body}</div></div>;
  return <div className="xl-hscroll" onMouseDown={e => e.stopPropagation()}>{body}</div>;
}

function ZoomSlider({ zoom, onZoom, onDialog }: { zoom: number; onZoom: (z: number) => void; onDialog: () => void }) {
  const ref = useRef<HTMLDivElement>(null);
  const toPos = (z: number) => (z <= 1 ? (z - 0.1) / 0.9 * 0.5 : 0.5 + (z - 1) / 3 * 0.5);
  const fromPos = (t: number) => (t <= 0.5 ? 0.1 + t / 0.5 * 0.9 : 1 + (t - 0.5) / 0.5 * 3);
  const drag = (e: React.MouseEvent) => {
    const r = ref.current!.getBoundingClientRect();
    const set = (x: number) => { let z = fromPos(Math.max(0, Math.min(1, (x - r.left) / r.width))); if (Math.abs(z - 1) < 0.06) z = 1; onZoom(Math.round(z * 100) / 100); };
    set(e.clientX);
    const mv = (ev: MouseEvent) => set(ev.clientX);
    const up = () => { window.removeEventListener('mousemove', mv); window.removeEventListener('mouseup', up); };
    window.addEventListener('mousemove', mv); window.addEventListener('mouseup', up);
  };
  return (
    <div className="xl-zoom">
      <button aria-label="Zoom Out" onClick={() => onZoom(Math.max(0.1, Math.round((zoom - 0.1) * 10) / 10))}><Icon name="zoomOut" size={16} /></button>
      <div className="xl-zoom-track" ref={ref} onMouseDown={drag}><div className="xl-zoom-thumb" style={{ left: `${toPos(zoom) * 100}%` }} /></div>
      <button aria-label="Zoom In" onClick={() => onZoom(Math.min(4, Math.round((zoom + 0.1) * 10) / 10))}><Icon name="zoomIn" size={16} /></button>
      <button className="xl-zoom-pct" style={{ width: 'auto' }} onClick={onDialog} title="Zoom level">{Math.round(zoom * 100)}%</button>
    </div>
  );
}

function Backstage({ page, setPage, name, path, pv, onClose, onSave, onSaveAs, onCloseWindow, onOpen, persona }: { page: 'home' | 'info' | 'saveas' | 'open'; setPage: (p: 'home' | 'info' | 'saveas' | 'open') => void; name: string; path?: string; pv: boolean; onClose: () => void; onSave: () => void; onSaveAs: () => void; onCloseWindow: () => void; onOpen: () => void; persona: string }) {
  return (
    <div className="xl-backstage" data-control-id="excel.save_save_as_xlsx">
      <div className="xl-bs-nav">
        <button onClick={onClose} aria-label="Back"><Icon name="back" size={20} /></button>
        <button className={page === 'home' ? 'active' : ''} onClick={() => setPage('home')}><Icon name="home" />Home</button>
        <button disabled><Icon name="new" />New</button>
        <button className={page === 'open' ? 'active' : ''} onClick={onOpen}><Icon name="open" />Open</button>
        <hr />
        <button className={page === 'info' ? 'active' : ''} onClick={() => setPage('info')}><Icon name="info" />Info</button>
        <button onClick={onSave} disabled={pv} data-control-id="excel.save_save_as_xlsx" data-control-item="Save"><Icon name="save" />Save</button>
        <button onClick={onSaveAs} disabled={pv} data-control-id="excel.save_save_as_xlsx" data-control-item="Save As"><Icon name="save" />Save As</button>
        <button disabled><Icon name="print" />Print</button>
        <button disabled><Icon name="share" />Share</button>
        <button disabled><Icon name="export" />Export</button>
        <button onClick={onCloseWindow}><Icon name="close" />Close</button>
        <div style={{ flex: 1 }} />
        <button disabled>Account</button>
        <button disabled>Options</button>
      </div>
      <div className="xl-bs-main">
        {page === 'home' && <><h1>Good afternoon</h1><p>Recent workbooks appear here.</p></>}
        {page === 'info' && (
          <>
            <h1>Info</h1>
            <div style={{ fontSize: 20, fontWeight: 300 }}>{stemOf(name)}</div>
            <div style={{ color: '#666', marginTop: 4 }}>{path ? `Documents » ${path.split('/').slice(0, -1).join(' » ')}` : ''}</div>
            <div className="xl-bs-props">
              <span>Size</span><span>—</span>
              <span>Last Modified</span><span>—</span>
              <span>Author</span><span>{persona}</span>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

// ------------------------------------------------------------------ helpers

function cloneCell(c: Cell | undefined): Cell | undefined { return c ? { ...c, shared: undefined } : undefined; }
function rangeCells(g: Range): Addr[] { const out: Addr[] = []; for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) out.push({ r, c }); return out; }
function escHtml(s: string) { return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function serialNow() { const d = new Date(now()); const p = new Intl.DateTimeFormat('en-US', { year: 'numeric', month: 'numeric', day: 'numeric' }).formatToParts(d); const g = (t: string) => Number(p.find(x => x.type === t)?.value); return (Date.UTC(g('year'), g('month') - 1, g('day')) - Date.UTC(1899, 11, 30)) / 86400000; }

/** references after inserting (n > 0) or deleting (n < 0) rows/columns at `at` */
function adjustRefs(f: string, sameSheet: boolean, sheetName: string, axis: 'row' | 'col', at: number, n: number): string {
  return f.replace(/((?:'[^']+'|[A-Za-z_][\w.]*)!)?(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![\w(])/g, (m, sheet: string | undefined, d1: string, col: string, d2: string, row: string) => {
    const target = sheet ? sheet.slice(0, -1).replace(/^'|'$/g, '') : null;
    if (target ? target !== sheetName : !sameSheet) return m;
    let r = Number(row) - 1;
    let c = 0; for (const ch of col) c = c * 26 + ch.charCodeAt(0) - 64; c -= 1;
    if (axis === 'row' && r >= at) { if (n < 0 && r < at - n) return (sheet ?? '') + '#REF!'; r += n; }
    if (axis === 'col' && c >= at) { if (n < 0 && c < at - n) return (sheet ?? '') + '#REF!'; c += n; }
    return `${sheet ?? ''}${d1}${colName(c)}${d2}${r + 1}`;
  });
}

export { isErr };
