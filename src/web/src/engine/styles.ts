// SpreadsheetML styles: fonts, fills, borders, number formats and cell formats (cellXfs), resolved
// to CSS-ready values; plus appending new formats on save without touching existing ones.

import { kid, kids, attrNum, attrBool, textOf, escAttr, type XDocument, type XNode } from './xml.ts';
import { BUILTIN_FORMATS } from './numfmt.ts';

export interface Font { name: string; sz: number; b?: boolean; i?: boolean; u?: string; strike?: boolean; color?: string; vertAlign?: string; family?: number; scheme?: string }
export interface Fill { pattern: string; fg?: string; bg?: string }
export interface BorderSide { style: string; color?: string }
export interface Border { left?: BorderSide; right?: BorderSide; top?: BorderSide; bottom?: BorderSide }
export interface Align { h?: string; v?: string; wrap?: boolean; indent?: number; rotation?: number; shrink?: boolean }
export interface Xf { numFmtId: number; fontId: number; fillId: number; borderId: number; xfId: number; align?: Align; quotePrefix?: boolean; raw?: string; /** <protection locked="0">: editable on a protected sheet */ unlocked?: boolean; /** <protection hidden="1"> */ hiddenFormula?: boolean }

export interface ResolvedStyle { font: Font; fill: Fill; border: Border; numFmt: string; numFmtId: number; align: Align; quotePrefix: boolean; locked: boolean }

// Excel's default Office theme colours (used when theme1.xml is missing)
const DEFAULT_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];

const INDEXED = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];

export function parseTheme(doc: XDocument | null): string[] {
  if (!doc) return DEFAULT_THEME.slice();
  const out: string[] = [];
  const find = (n: XNode): XNode | undefined => n.local === 'clrScheme' ? n : n.children.map(find).find(Boolean);
  const scheme = find(doc.root);
  if (!scheme) return DEFAULT_THEME.slice();
  const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  for (const name of order) {
    const el = kid(scheme, name);
    const c = el?.children[0];
    out.push(c?.attrs.lastClr ?? c?.attrs.val ?? DEFAULT_THEME[out.length]);
  }
  return out;
}

function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const r = parseInt(hex.slice(0, 2), 16) / 255, g = parseInt(hex.slice(2, 4), 16) / 255, b = parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0; let l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const hue = (p: number, q: number, t: number) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
  let rr: number, gg: number, bb: number;
  if (s === 0) rr = gg = bb = l;
  else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s; const p = 2 * l - q; rr = hue(p, q, h + 1 / 3); gg = hue(p, q, h); bb = hue(p, q, h - 1 / 3); }
  const to = (x: number) => Math.round(Math.max(0, Math.min(1, x)) * 255).toString(16).padStart(2, '0');
  return (to(rr) + to(gg) + to(bb)).toUpperCase();
}

export function colorOf(node: XNode | undefined, theme: string[]): string | undefined {
  if (!node) return undefined;
  if (node.attrs.auto === '1') return undefined;
  let hex: string | undefined;
  if (node.attrs.rgb) hex = node.attrs.rgb.length === 8 ? node.attrs.rgb.slice(2) : node.attrs.rgb;
  else if (node.attrs.theme !== undefined) hex = theme[Number(node.attrs.theme)];
  else if (node.attrs.indexed !== undefined) { const i = Number(node.attrs.indexed); if (i === 64 || i === 65) return undefined; hex = INDEXED[i]; }
  if (!hex) return undefined;
  const tint = Number(node.attrs.tint ?? 0);
  return '#' + applyTint(hex.toUpperCase(), tint);
}

export class Styles {
  numFmts = new Map<number, string>();
  fonts: Font[] = [];
  fills: Fill[] = [];
  borders: Border[] = [];
  xfs: Xf[] = [];
  /** additions made by edits (appended on save) */
  added = { fonts: [] as Font[], fills: [] as Fill[], borders: [] as Border[], xfs: [] as Xf[], numFmts: [] as { id: number; code: string }[] };
  private cache = new Map<number, ResolvedStyle>();
  doc: XDocument | null;
  theme: string[];
  constructor(doc: XDocument | null, theme: string[]) {
    this.doc = doc; this.theme = theme;
    if (doc) this.parse(doc);
    if (!this.fonts.length) this.fonts.push({ name: 'Calibri', sz: 11 });
    if (!this.fills.length) this.fills.push({ pattern: 'none' }, { pattern: 'gray125' });
    if (!this.borders.length) this.borders.push({});
    if (!this.xfs.length) this.xfs.push({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0 });
  }

  private parse(doc: XDocument) {
    const root = doc.root;
    for (const n of kids(kid(root, 'numFmts'), 'numFmt')) this.numFmts.set(attrNum(n, 'numFmtId', 0), n.attrs.formatCode ?? 'General');
    for (const f of kids(kid(root, 'fonts'), 'font')) {
      const font: Font = { name: kid(f, 'name')?.attrs.val ?? '', sz: attrNum(kid(f, 'sz'), 'val', 0) };
      const flag = (n: string) => { const k = kid(f, n); return k ? (k.attrs.val === undefined || attrBool(k, 'val', true)) : false; };
      if (flag('b')) font.b = true;
      if (flag('i')) font.i = true;
      if (flag('strike')) font.strike = true;
      const u = kid(f, 'u'); if (u) font.u = u.attrs.val ?? 'single';
      const va = kid(f, 'vertAlign'); if (va) font.vertAlign = va.attrs.val;
      const c = colorOf(kid(f, 'color'), this.theme); if (c) font.color = c;
      const sch = kid(f, 'scheme'); if (sch) font.scheme = sch.attrs.val;
      this.fonts.push(font);
    }
    for (const f of kids(kid(root, 'fills'), 'fill')) {
      const p = kid(f, 'patternFill');
      const g = kid(f, 'gradientFill');
      if (g) { const stop = kid(g, 'stop'); this.fills.push({ pattern: 'solid', fg: colorOf(kid(stop, 'color'), this.theme) }); continue; }
      this.fills.push({ pattern: p?.attrs.patternType ?? 'none', fg: colorOf(kid(p, 'fgColor'), this.theme), bg: colorOf(kid(p, 'bgColor'), this.theme) });
    }
    for (const b of kids(kid(root, 'borders'), 'border')) {
      const side = (n: string): BorderSide | undefined => { const s = kid(b, n); if (!s || !s.attrs.style) return undefined; return { style: s.attrs.style, color: colorOf(kid(s, 'color'), this.theme) }; };
      this.borders.push({ left: side('left') ?? side('start'), right: side('right') ?? side('end'), top: side('top'), bottom: side('bottom') });
    }
    for (const x of kids(kid(root, 'cellXfs'), 'xf')) {
      const xf: Xf = { numFmtId: attrNum(x, 'numFmtId', 0), fontId: attrNum(x, 'fontId', 0), fillId: attrNum(x, 'fillId', 0), borderId: attrNum(x, 'borderId', 0), xfId: attrNum(x, 'xfId', 0), raw: doc.text.slice(x.start, x.end) };
      const al = kid(x, 'alignment');
      if (al) xf.align = { h: al.attrs.horizontal, v: al.attrs.vertical, wrap: attrBool(al, 'wrapText'), indent: attrNum(al, 'indent', 0) || undefined, rotation: attrNum(al, 'textRotation', 0) || undefined, shrink: attrBool(al, 'shrinkToFit') };
      if (attrBool(x, 'quotePrefix')) xf.quotePrefix = true;
      const pr = kid(x, 'protection');
      if (pr) { if (pr.attrs.locked === '0' || pr.attrs.locked === 'false') xf.unlocked = true; if (attrBool(pr, 'hidden')) xf.hiddenFormula = true; }
      this.xfs.push(xf);
    }
  }

  allFonts() { return [...this.fonts, ...this.added.fonts]; }
  allFills() { return [...this.fills, ...this.added.fills]; }
  allBorders() { return [...this.borders, ...this.added.borders]; }
  allXfs() { return [...this.xfs, ...this.added.xfs]; }

  formatCode(id: number): string {
    return this.numFmts.get(id) ?? this.added.numFmts.find(n => n.id === id)?.code ?? BUILTIN_FORMATS[id] ?? 'General';
  }

  resolve(xfIndex: number): ResolvedStyle {
    const hit = this.cache.get(xfIndex);
    if (hit) return hit;
    const xfs = this.allXfs();
    const xf = xfs[xfIndex] ?? xfs[0];
    const fonts = this.allFonts();
    const base = fonts[0];
    const f = fonts[xf.fontId] ?? base;
    const font: Font = { ...f, name: f.name || base.name, sz: f.sz || base.sz };
    const r: ResolvedStyle = {
      font, fill: this.allFills()[xf.fillId] ?? { pattern: 'none' }, border: this.allBorders()[xf.borderId] ?? {},
      numFmt: this.formatCode(xf.numFmtId), numFmtId: xf.numFmtId, align: xf.align ?? {}, quotePrefix: !!xf.quotePrefix, locked: !xf.unlocked,
    };
    this.cache.set(xfIndex, r);
    return r;
  }

  /** index of a cell format equal to `xf` (reusing an existing or added one, else adding it) */
  xfFor(xf: Xf): number {
    const all = this.allXfs();
    const same = (a: Xf) => a.numFmtId === xf.numFmtId && a.fontId === xf.fontId && a.fillId === xf.fillId && a.borderId === xf.borderId
      && JSON.stringify(a.align ?? {}) === JSON.stringify(xf.align ?? {}) && !!a.quotePrefix === !!xf.quotePrefix
      && !!a.unlocked === !!xf.unlocked && !!a.hiddenFormula === !!xf.hiddenFormula;
    const i = all.findIndex(same);
    if (i >= 0) return i;
    this.added.xfs.push({ ...xf, raw: undefined });
    this.cache.clear();
    return all.length;
  }
  fontFor(font: Font): number {
    const all = this.allFonts();
    const i = all.findIndex(f => JSON.stringify(normFont(f)) === JSON.stringify(normFont(font)));
    if (i >= 0) return i;
    this.added.fonts.push(font);
    return all.length;
  }
  fillFor(fill: Fill): number {
    const all = this.allFills();
    const i = all.findIndex(f => f.pattern === fill.pattern && (f.fg ?? '') === (fill.fg ?? ''));
    if (i >= 0) return i;
    this.added.fills.push(fill);
    return all.length;
  }
  borderFor(b: Border): number {
    const all = this.allBorders();
    const k = (x: Border) => JSON.stringify([x.left, x.right, x.top, x.bottom].map(s => s ? [s.style, s.color ?? ''] : null));
    const i = all.findIndex(x => k(x) === k(b));
    if (i >= 0) return i;
    this.added.borders.push(b);
    return all.length;
  }
  numFmtFor(code: string): number {
    for (const [id, c] of Object.entries(BUILTIN_FORMATS)) if (c === code) return Number(id);
    for (const [id, c] of this.numFmts) if (c === code) return id;
    const a = this.added.numFmts.find(n => n.code === code);
    if (a) return a.id;
    const id = Math.max(163, ...this.numFmts.keys(), ...this.added.numFmts.map(n => n.id)) + 1;
    this.added.numFmts.push({ id, code });
    return id;
  }
  dirty(): boolean { const a = this.added; return !!(a.fonts.length || a.fills.length || a.borders.length || a.xfs.length || a.numFmts.length); }
}

function normFont(f: Font) { return [f.name, f.sz, !!f.b, !!f.i, f.u ?? '', !!f.strike, (f.color ?? '').toUpperCase(), f.vertAlign ?? '']; }

// ------------------------------------------------------------------ writing

function rgb(c?: string) { return c ? `FF${c.replace('#', '').toUpperCase()}` : undefined; }

export function fontXml(f: Font): string {
  let s = '<font>';
  if (f.b) s += '<b/>';
  if (f.i) s += '<i/>';
  if (f.strike) s += '<strike/>';
  if (f.u) s += f.u === 'single' ? '<u/>' : `<u val="${escAttr(f.u)}"/>`;
  if (f.vertAlign) s += `<vertAlign val="${f.vertAlign}"/>`;
  s += `<sz val="${f.sz}"/>`;
  s += f.color ? `<color rgb="${rgb(f.color)}"/>` : '<color theme="1"/>';
  s += `<name val="${escAttr(f.name)}"/>`;
  s += `<family val="${f.family ?? 2}"/>`;
  if (f.scheme) s += `<scheme val="${f.scheme}"/>`;
  return s + '</font>';
}
export function fillXml(f: Fill): string {
  if (f.pattern === 'none' || !f.fg) return `<fill><patternFill patternType="${f.pattern === 'gray125' ? 'gray125' : 'none'}"/></fill>`;
  return `<fill><patternFill patternType="${f.pattern}"><fgColor rgb="${rgb(f.fg)}"/><bgColor indexed="64"/></patternFill></fill>`;
}
export function borderXml(b: Border): string {
  const side = (n: string, s?: BorderSide) => s ? `<${n} style="${s.style}">${s.color ? `<color rgb="${rgb(s.color)}"/>` : '<color indexed="64"/>'}</${n}>` : `<${n}/>`;
  return `<border>${side('left', b.left)}${side('right', b.right)}${side('top', b.top)}${side('bottom', b.bottom)}<diagonal/></border>`;
}
export function xfXml(x: Xf): string {
  const attrs = [`numFmtId="${x.numFmtId}"`, `fontId="${x.fontId}"`, `fillId="${x.fillId}"`, `borderId="${x.borderId}"`, `xfId="${x.xfId}"`];
  if (x.quotePrefix) attrs.push('quotePrefix="1"');
  if (x.numFmtId) attrs.push('applyNumberFormat="1"');
  if (x.fontId) attrs.push('applyFont="1"');
  if (x.fillId) attrs.push('applyFill="1"');
  if (x.borderId) attrs.push('applyBorder="1"');
  const a = x.align;
  const hasAlign = a && (a.h || a.v || a.wrap || a.indent || a.rotation || a.shrink);
  if (hasAlign) attrs.push('applyAlignment="1"');
  const prot = x.unlocked || x.hiddenFormula ? `<protection${x.unlocked ? ' locked="0"' : ''}${x.hiddenFormula ? ' hidden="1"' : ''}/>` : '';
  if (prot) attrs.push('applyProtection="1"');
  if (!hasAlign) return prot ? `<xf ${attrs.join(' ')}>${prot}</xf>` : `<xf ${attrs.join(' ')}/>`;
  const al: string[] = [];
  if (a!.h) al.push(`horizontal="${a!.h}"`);
  if (a!.v) al.push(`vertical="${a!.v}"`);
  if (a!.rotation) al.push(`textRotation="${a!.rotation}"`);
  if (a!.wrap) al.push('wrapText="1"');
  if (a!.indent) al.push(`indent="${a!.indent}"`);
  if (a!.shrink) al.push('shrinkToFit="1"');
  return `<xf ${attrs.join(' ')}><alignment ${al.join(' ')}/>${prot}</xf>`;
}

export { textOf };
