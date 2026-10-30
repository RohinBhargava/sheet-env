// Our own formula engine: tokenizer, parser and evaluator for the spreadsheet functions business and
// cap-table workbooks use. Written from the published function semantics; no third-party engine.

import { parseAddr, colIndex, colName, MAX_ROW, MAX_COL, type Range } from './ref.ts';
import { partsToSerial, serialToParts, formatValue, BUILTIN_FORMATS } from './numfmt.ts';

export interface XlError { error: string }
export type Scalar = number | string | boolean | null | XlError;
export type Matrix = Scalar[][];
export type Value = Scalar | Matrix | RefVal;
export interface RefVal { ref: true; sheet: string; range: Range }

export const ERR = {
  div0: { error: '#DIV/0!' }, na: { error: '#N/A' }, name: { error: '#NAME?' }, num: { error: '#NUM!' },
  ref: { error: '#REF!' }, value: { error: '#VALUE!' }, null: { error: '#NULL!' }, calc: { error: '#CALC!' }, spill: { error: '#SPILL!' },
} satisfies Record<string, XlError>;

export function isErr(v: unknown): v is XlError { return !!v && typeof v === 'object' && 'error' in (v as object); }

// ------------------------------------------------------------------ tokens

type Tok =
  | { k: 'num'; v: number } | { k: 'str'; v: string } | { k: 'bool'; v: boolean } | { k: 'err'; v: string }
  | { k: 'ref'; sheet?: string; a: string; b?: string } | { k: 'name'; v: string } | { k: 'fn'; v: string }
  | { k: 'op'; v: string } | { k: 'lp' } | { k: 'rp' } | { k: 'comma' } | { k: 'lb' } | { k: 'rb' } | { k: 'semi' };

const ERRORS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!'];

export function tokenize(src: string): Tok[] {
  const s = src.startsWith('=') ? src.slice(1) : src;
  const out: Tok[] = [];
  let i = 0;
  const refRe = /^(\$?[A-Za-z]{1,3}\$?\d{1,7})(?::(\$?[A-Za-z]{1,3}\$?\d{1,7}))?/;
  const colRe = /^(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})(?![\w(])/;
  const rowRe = /^(\$?\d{1,7}):(\$?\d{1,7})/;
  while (i < s.length) {
    const ch = s[i];
    if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t') { i++; continue; }
    if (ch === '"') {
      let j = i + 1, v = '';
      while (j < s.length) { if (s[j] === '"') { if (s[j + 1] === '"') { v += '"'; j += 2; continue; } break; } v += s[j++]; }
      out.push({ k: 'str', v }); i = j + 1; continue;
    }
    if (ch === '#') {
      const e = ERRORS.find(x => s.slice(i, i + x.length).toUpperCase() === x);
      if (e) { out.push({ k: 'err', v: e }); i += e.length; continue; }
    }
    if (/[0-9.]/.test(ch)) {
      const rr = rowRe.exec(s.slice(i));
      const prev = out[out.length - 1];
      if (rr && !(prev && (prev.k === 'num' || prev.k === 'ref'))) { out.push({ k: 'ref', a: rr[1], b: rr[2] }); i += rr[0].length; continue; }
      const m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(s.slice(i));
      if (m) { out.push({ k: 'num', v: Number(m[0]) }); i += m[0].length; continue; }
    }
    // sheet-qualified reference
    let sheet: string | undefined;
    let j = i;
    if (ch === "'") {
      let k = i + 1, name = '';
      while (k < s.length) { if (s[k] === "'") { if (s[k + 1] === "'") { name += "'"; k += 2; continue; } break; } name += s[k++]; }
      if (s[k + 1] === '!') { sheet = name; j = k + 2; }
    } else {
      const m = /^([A-Za-z_\u00c0-\uffff][\w.\u00c0-\uffff]*)!/.exec(s.slice(i));
      if (m) { sheet = m[1]; j = i + m[0].length; }
    }
    const rest = s.slice(j);
    const cm = colRe.exec(rest);
    if (cm) { out.push({ k: 'ref', sheet, a: cm[1], b: cm[2] }); i = j + cm[0].length; continue; }
    const rm2 = rowRe.exec(rest);
    if (sheet && rm2) { out.push({ k: 'ref', sheet, a: rm2[1], b: rm2[2] }); i = j + rm2[0].length; continue; }
    const rm = refRe.exec(rest);
    if (rm && !/^[\w(]/.test(rest.slice(rm[0].length))) { out.push({ k: 'ref', sheet, a: rm[1], b: rm[2] }); i = j + rm[0].length; continue; }
    if (sheet) { out.push({ k: 'err', v: '#REF!' }); i = j; continue; }
    const nm = /^[A-Za-z_\\][\w.]*/.exec(s.slice(i));
    if (nm) {
      const word = nm[0];
      const after = s.slice(i + word.length);
      if (/^\s*\(/.test(after)) { out.push({ k: 'fn', v: word.toUpperCase().replace(/^_XLFN\./, '').replace(/^_XLWS\./, '') }); i += word.length; continue; }
      if (/^(TRUE|FALSE)$/i.test(word)) { out.push({ k: 'bool', v: word.toUpperCase() === 'TRUE' }); i += word.length; continue; }
      out.push({ k: 'name', v: word }); i += word.length; continue;
    }
    if (ch === '(') { out.push({ k: 'lp' }); i++; continue; }
    if (ch === ')') { out.push({ k: 'rp' }); i++; continue; }
    if (ch === ',') { out.push({ k: 'comma' }); i++; continue; }
    if (ch === '{') { out.push({ k: 'lb' }); i++; continue; }
    if (ch === '}') { out.push({ k: 'rb' }); i++; continue; }
    if (ch === ';') { out.push({ k: 'semi' }); i++; continue; }
    const two = s.slice(i, i + 2);
    if (two === '<=' || two === '>=' || two === '<>') { out.push({ k: 'op', v: two }); i += 2; continue; }
    if ('+-*/^&=<>%:'.includes(ch)) { out.push({ k: 'op', v: ch }); i++; continue; }
    throw new Error(`unexpected "${ch}"`);
  }
  return out;
}

// ------------------------------------------------------------------ AST

export type Node =
  | { t: 'num'; v: number } | { t: 'str'; v: string } | { t: 'bool'; v: boolean } | { t: 'err'; v: string }
  | { t: 'ref'; sheet?: string; a: string; b?: string } | { t: 'name'; v: string }
  | { t: 'fn'; name: string; args: Node[] } | { t: 'un'; op: string; e: Node } | { t: 'bin'; op: string; l: Node; r: Node }
  | { t: 'pct'; e: Node } | { t: 'arr'; rows: Node[][] } | { t: 'missing' };

const PREC: Record<string, number> = { ':': 8, '^': 6, '*': 5, '/': 5, '+': 4, '-': 4, '&': 3, '=': 2, '<>': 2, '<': 2, '>': 2, '<=': 2, '>=': 2 };

export function parseFormula(src: string): Node {
  const toks = tokenize(src);
  let p = 0;
  const peek = () => toks[p];
  function primary(): Node {
    const t = toks[p++];
    if (!t) throw new Error('unexpected end');
    switch (t.k) {
      case 'num': return { t: 'num', v: t.v };
      case 'str': return { t: 'str', v: t.v };
      case 'bool': return { t: 'bool', v: t.v };
      case 'err': return { t: 'err', v: t.v };
      case 'ref': return { t: 'ref', sheet: t.sheet, a: t.a, b: t.b };
      case 'name': return { t: 'name', v: t.v };
      case 'op':
        if (t.v === '-' || t.v === '+') return { t: 'un', op: t.v, e: expr(7) };
        throw new Error(`unexpected ${t.v}`);
      case 'lp': { const e = expr(0); if (toks[p++]?.k !== 'rp') throw new Error('missing )'); return e; }
      case 'lb': {
        const rows: Node[][] = [[]];
        while (peek() && peek().k !== 'rb') {
          rows[rows.length - 1].push(expr(0));
          const n = toks[p];
          if (n?.k === 'comma') p++;
          else if (n?.k === 'semi') { p++; rows.push([]); }
        }
        p++;
        return { t: 'arr', rows };
      }
      case 'fn': {
        p++; // (
        const args: Node[] = [];
        if (peek()?.k === 'rp') { p++; return { t: 'fn', name: t.v, args }; }
        for (;;) {
          if (peek()?.k === 'comma' || peek()?.k === 'rp') args.push({ t: 'missing' });
          else args.push(expr(0));
          const n = toks[p++];
          if (!n) throw new Error('missing )');
          if (n.k === 'rp') break;
          if (n.k !== 'comma') throw new Error('expected ,');
        }
        return { t: 'fn', name: t.v, args };
      }
      default: throw new Error('unexpected token');
    }
  }
  function expr(minPrec: number): Node {
    let left = primary();
    for (;;) {
      const t = peek();
      if (!t || t.k !== 'op') break;
      if (t.v === '%') { p++; left = { t: 'pct', e: left }; continue; }
      const prec = PREC[t.v];
      if (prec === undefined || prec < minPrec) break;
      p++;
      const right = expr(prec + 1);
      left = { t: 'bin', op: t.v, l: left, r: right };
    }
    return left;
  }
  const e = expr(0);
  if (p < toks.length) throw new Error('unexpected input');
  return e;
}

// ------------------------------------------------------------------ reference shifting

/** move relative references by (dr, dc): shared formulas, fill, copy/paste */
export function shiftFormula(src: string, dr: number, dc: number): string {
  const body = src.startsWith('=') ? src.slice(1) : src;
  let out = '';
  let i = 0;
  const re = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/y;
  while (i < body.length) {
    const ch = body[i];
    if (ch === '"') { const j = body.indexOf('"', i + 1); out += body.slice(i, j < 0 ? undefined : j + 1); i = j < 0 ? body.length : j + 1; continue; }
    if (ch === "'") { const j = body.indexOf("'", i + 1); out += body.slice(i, j + 1); i = j + 1; continue; }
    const prev = body[i - 1] ?? '';
    if (/[A-Za-z$]/.test(ch) && !/[\w.]/.test(prev)) {
      re.lastIndex = i;
      const m = re.exec(body);
      if (m && !/[\w(!]/.test(body[i + m[0].length] ?? '')) {
        let c = colIndex(m[2]); let r = Number(m[4]) - 1;
        if (!m[1]) c += dc;
        if (!m[3]) r += dr;
        if (c < 0 || r < 0 || c > MAX_COL || r > MAX_ROW) out += '#REF!';
        else out += `${m[1]}${colName(c)}${m[3]}${r + 1}`;
        i += m[0].length; continue;
      }
      const w = /^[A-Za-z_][\w.]*/.exec(body.slice(i));
      if (w) { out += w[0]; i += w[0].length; continue; }
    }
    out += ch; i++;
  }
  return (src.startsWith('=') ? '=' : '') + out;
}

// ------------------------------------------------------------------ evaluation

export interface EvalContext {
  /** value of a cell (evaluating formulas as needed) */
  cell(sheet: string, r: number, c: number): Scalar;
  /** used extent of a sheet (to clip whole-column references) */
  extent(sheet: string): { rows: number; cols: number };
  sheet: string;
  /** current cell (for ROW()/COLUMN() without args) */
  r: number; c: number;
  hasSheet(name: string): boolean;
  name?(name: string): Node | null;
  /** scenario clock (TODAY/NOW) in epoch ms */
  nowMs: number;
  date1904?: boolean;
}

function toNum(v: Scalar): number | XlError {
  if (v === null) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isErr(v)) return v;
  const s = v.trim();
  if (s === '') return ERR.value;
  const pct = /%$/.test(s);
  const n = Number(s.replace(/[$,]/g, '').replace(/%$/, ''));
  if (Number.isFinite(n)) return pct ? n / 100 : n;
  const d = parseDateText(s);
  return d ?? ERR.value;
}
export function parseDateText(s: string): number | null {
  let m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(s);
  if (m) { let y = Number(m[3]); if (y < 100) y += y < 30 ? 2000 : 1900; return partsToSerial(y, Number(m[1]), Number(m[2])); }
  m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (m) return partsToSerial(Number(m[1]), Number(m[2]), Number(m[3]));
  return null;
}
function toStr(v: Scalar): string | XlError {
  if (v === null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isErr(v)) return v;
  return formatValue(v, 'General').text;
}
function toBool(v: Scalar): boolean | XlError {
  if (v === null) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (isErr(v)) return v;
  if (/^true$/i.test(v)) return true;
  if (/^false$/i.test(v)) return false;
  return ERR.value;
}

function isRef(v: Value): v is RefVal { return !!v && typeof v === 'object' && 'ref' in (v as object); }
function isMatrix(v: Value): v is Matrix { return Array.isArray(v); }

export class Evaluator {
  ctx: EvalContext;
  constructor(ctx: EvalContext) { this.ctx = ctx; }

  resolveRef(n: { sheet?: string; a: string; b?: string }): RefVal | XlError {
    const sheet = n.sheet ?? this.ctx.sheet;
    if (!this.ctx.hasSheet(sheet)) return ERR.ref;
    let range: Range;
    const a = n.a.replace(/\$/g, ''), b = (n.b ?? n.a).replace(/\$/g, '');
    if (/^[A-Za-z]+$/.test(a)) {
      range = { r1: 0, c1: colIndex(a), r2: MAX_ROW, c2: colIndex(b) };
    } else if (/^\d+$/.test(a)) {
      range = { r1: Number(a) - 1, c1: 0, r2: Number(b) - 1, c2: MAX_COL };
    } else {
      const x = parseAddr(a), y = parseAddr(b);
      if (!x || !y) return ERR.ref;
      range = { r1: Math.min(x.r, y.r), c1: Math.min(x.c, y.c), r2: Math.max(x.r, y.r), c2: Math.max(x.c, y.c) };
    }
    return { ref: true, sheet, range };
  }

  /** materialize a reference as a matrix (clipped to the used extent) */
  matrix(v: Value): Matrix {
    if (isMatrix(v)) return v;
    if (isRef(v)) {
      const ext = this.ctx.extent(v.sheet);
      const r2 = Math.min(v.range.r2, Math.max(ext.rows - 1, v.range.r1));
      const c2 = Math.min(v.range.c2, Math.max(ext.cols - 1, v.range.c1));
      const out: Matrix = [];
      for (let r = v.range.r1; r <= r2; r++) {
        const row: Scalar[] = [];
        for (let c = v.range.c1; c <= c2; c++) row.push(this.ctx.cell(v.sheet, r, c));
        out.push(row);
      }
      return out;
    }
    return [[v as Scalar]];
  }

  /** a single value: a 1x1 ref, or implicit intersection with the current row/column */
  scalar(v: Value): Scalar {
    if (isRef(v)) {
      const g = v.range;
      if (g.r1 === g.r2 && g.c1 === g.c2) return this.ctx.cell(v.sheet, g.r1, g.c1);
      if (g.c1 === g.c2 && this.ctx.r >= g.r1 && this.ctx.r <= g.r2) return this.ctx.cell(v.sheet, this.ctx.r, g.c1);
      if (g.r1 === g.r2 && this.ctx.c >= g.c1 && this.ctx.c <= g.c2) return this.ctx.cell(v.sheet, g.r1, this.ctx.c);
      return ERR.value;
    }
    if (isMatrix(v)) return v[0]?.[0] ?? null;
    return v;
  }

  eval(n: Node): Value {
    switch (n.t) {
      case 'num': return n.v;
      case 'str': return n.v;
      case 'bool': return n.v;
      case 'err': return { error: n.v };
      case 'missing': return null;
      case 'ref': return this.resolveRef(n);
      case 'name': {
        const def = this.ctx.name?.(n.v);
        if (def) return this.eval(def);
        return ERR.name;
      }
      case 'arr': return n.rows.map(row => row.map(x => this.scalar(this.eval(x))));
      case 'pct': { const x = toNum(this.scalar(this.eval(n.e))); return isErr(x) ? x : x / 100; }
      case 'un': {
        const v = this.eval(n.e);
        if (isMatrix(v) || (isRef(v) && !(v.range.r1 === v.range.r2 && v.range.c1 === v.range.c2) && false)) return this.mapM(this.matrix(v), x => { const y = toNum(x); return isErr(y) ? y : n.op === '-' ? -y : y; });
        const x = toNum(this.scalar(v));
        return isErr(x) ? x : n.op === '-' ? -x : x;
      }
      case 'bin': return this.binary(n.op, n.l, n.r);
      case 'fn': return this.call(n.name, n.args);
    }
  }

  private mapM(m: Matrix, f: (x: Scalar) => Scalar): Matrix { return m.map(r => r.map(f)); }

  private binary(op: string, l: Node, r: Node): Value {
    if (op === ':') {
      const a = this.eval(l), b = this.eval(r);
      if (isRef(a) && isRef(b)) return { ref: true, sheet: a.sheet, range: { r1: Math.min(a.range.r1, b.range.r1), c1: Math.min(a.range.c1, b.range.c1), r2: Math.max(a.range.r2, b.range.r2), c2: Math.max(a.range.c2, b.range.c2) } };
      return ERR.ref;
    }
    const lv = this.eval(l), rv = this.eval(r);
    const lm = isMatrix(lv) || (isRef(lv) && !single(lv)), rm = isMatrix(rv) || (isRef(rv) && !single(rv));
    if ((lm || rm) && this.arrayMode) {
      const A = this.matrix(lv), B = this.matrix(rv);
      const rows = Math.max(A.length, B.length), cols = Math.max(A[0]?.length ?? 0, B[0]?.length ?? 0);
      const out: Matrix = [];
      for (let i = 0; i < rows; i++) {
        const row: Scalar[] = [];
        for (let j = 0; j < cols; j++) {
          const a = A.length === 1 ? A[0][A[0].length === 1 ? 0 : j] : A[i]?.[A[0].length === 1 ? 0 : j];
          const b = B.length === 1 ? B[0][B[0].length === 1 ? 0 : j] : B[i]?.[B[0].length === 1 ? 0 : j];
          row.push(a === undefined || b === undefined ? ERR.na : scalarOp(op, a, b));
        }
        out.push(row);
      }
      return out;
    }
    return scalarOp(op, this.scalar(lv), this.scalar(rv));
  }

  /** inside SUMPRODUCT and array-taking arguments, operators work element-wise */
  arrayMode = false;

  private argList(args: Node[]): Value[] { return args.map(a => this.eval(a)); }

  /** flatten numbers for SUM-like functions (refs skip text/bools; direct args coerce) */
  private nums(vals: Value[], opts: { countText?: boolean } = {}): number[] | XlError {
    const out: number[] = [];
    for (const v of vals) {
      if (isRef(v) || isMatrix(v)) {
        for (const row of this.matrix(v)) for (const x of row) {
          if (isErr(x)) return x;
          if (typeof x === 'number') out.push(x);
          else if (opts.countText && x !== null) out.push(0);
        }
      } else {
        if (v === null) continue;
        const x = toNum(v as Scalar);
        if (isErr(x)) return x;
        out.push(x);
      }
    }
    return out;
  }

  private flat(v: Value): Scalar[] { return this.matrix(v).flat(); }

  call(name: string, args: Node[]): Value {
    const f = FUNCS[name];
    if (!f) return ERR.name;
    try { return f(this, args); } catch (e) { if (isErr(e)) return e as XlError; throw e; }
  }

  // helpers exposed to the function table
  num(n: Node | undefined, dflt?: number): number {
    if (!n || n.t === 'missing') { if (dflt !== undefined) return dflt; throw ERR.value; }
    const x = toNum(this.scalar(this.eval(n)));
    if (isErr(x)) throw x;
    return x;
  }
  str(n: Node | undefined, dflt?: string): string {
    if (!n || n.t === 'missing') { if (dflt !== undefined) return dflt; throw ERR.value; }
    const x = toStr(this.scalar(this.eval(n)));
    if (isErr(x)) throw x;
    return x;
  }
  bool(n: Node | undefined, dflt?: boolean): boolean {
    if (!n || n.t === 'missing') { if (dflt !== undefined) return dflt; throw ERR.value; }
    const x = toBool(this.scalar(this.eval(n)));
    if (isErr(x)) throw x;
    return x;
  }
  val(n: Node): Scalar { return this.scalar(this.eval(n)); }
  all(args: Node[]): number[] { const r = this.nums(this.argList(args)); if (isErr(r)) throw r; return r; }
  numsA(args: Node[]): number[] { const r = this.nums(this.argList(args), { countText: true }); if (isErr(r)) throw r; return r; }
  flatArgs(args: Node[]): Scalar[] { return args.flatMap(a => this.flat(this.eval(a))); }
  arr(n: Node): Matrix { const prev = this.arrayMode; this.arrayMode = true; try { return this.matrix(this.eval(n)); } finally { this.arrayMode = prev; } }
  ref(n: Node): RefVal { const v = this.eval(n); if (!isRef(v)) throw ERR.value; return v; }
}

function single(v: RefVal) { return v.range.r1 === v.range.r2 && v.range.c1 === v.range.c2; }

function cmp(a: Scalar, b: Scalar): number {
  // Excel ordering: numbers < text < booleans; text case-insensitive
  const rank = (x: Scalar) => (x === null ? 0 : typeof x === 'number' ? 1 : typeof x === 'string' ? 2 : 3);
  if (a === null && typeof b === 'string') a = '';
  if (b === null && typeof a === 'string') b = '';
  if (a === null && typeof b === 'number') a = 0;
  if (b === null && typeof a === 'number') b = 0;
  if (a === null && typeof b === 'boolean') a = false;
  if (b === null && typeof a === 'boolean') b = false;
  const ra = rank(a), rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'string' && typeof b === 'string') { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
  const x = Number(a), y = Number(b);
  return x < y ? -1 : x > y ? 1 : 0;
}

function scalarOp(op: string, a: Scalar, b: Scalar): Scalar {
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  switch (op) {
    case '&': { const x = toStr(a), y = toStr(b); if (isErr(x)) return x; if (isErr(y)) return y; return x + y; }
    case '=': return cmp(a, b) === 0;
    case '<>': return cmp(a, b) !== 0;
    case '<': return cmp(a, b) < 0;
    case '>': return cmp(a, b) > 0;
    case '<=': return cmp(a, b) <= 0;
    case '>=': return cmp(a, b) >= 0;
  }
  const x = toNum(a), y = toNum(b);
  if (isErr(x)) return x;
  if (isErr(y)) return y;
  switch (op) {
    case '+': return x + y;
    case '-': return x - y;
    case '*': return x * y;
    case '/': return y === 0 ? ERR.div0 : x / y;
    case '^': { const r = Math.pow(x, y); return Number.isFinite(r) ? r : ERR.num; }
  }
  return ERR.value;
}

// ------------------------------------------------------------------ criteria (SUMIF / COUNTIF ...)

function criterion(c: Scalar): (v: Scalar) => boolean {
  if (typeof c === 'number' || typeof c === 'boolean') return v => v === c || (typeof v === 'string' && toNum(v) === c);
  if (c === null) return v => v === null || v === '';
  if (isErr(c)) return v => isErr(v) && v.error === c.error;
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(c)!;
  const op = m[1] ?? '=';
  const rhs = m[2];
  const n = rhs !== '' && Number.isFinite(Number(rhs)) ? Number(rhs) : null;
  if (n !== null) {
    return v => {
      const x = typeof v === 'number' ? v : null;
      if (x === null) return op === '<>';
      switch (op) { case '=': return x === n; case '<>': return x !== n; case '<': return x < n; case '>': return x > n; case '<=': return x <= n; case '>=': return x >= n; }
      return false;
    };
  }
  if (op === '=' || op === '<>') {
    const pat = new RegExp('^' + rhs.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '[\\s\\S]*').replace(/\?/g, '.').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?') + '$', 'i');
    if (rhs === '') return op === '=' ? (v => v === null || v === '') : (v => !(v === null || v === ''));
    return op === '=' ? (v => typeof v === 'string' && pat.test(v)) : (v => !(typeof v === 'string' && pat.test(v)));
  }
  return v => typeof v === 'string' && (op === '<' ? cmp(v, rhs) < 0 : op === '>' ? cmp(v, rhs) > 0 : op === '<=' ? cmp(v, rhs) <= 0 : cmp(v, rhs) >= 0);
}

// ------------------------------------------------------------------ function table

type Fn = (e: Evaluator, a: Node[]) => Value;

const roundTo = (x: number, d: number, mode: 'n' | 'u' | 'd') => {
  const f = Math.pow(10, d);
  const y = x * f;
  const r = mode === 'n' ? Math.sign(y) * Math.round(Math.abs(y) + 1e-9) : mode === 'u' ? Math.sign(y) * Math.ceil(Math.abs(y) - 1e-9) : Math.sign(y) * Math.floor(Math.abs(y) + 1e-9);
  return r / f;
};

function lookupMatch(e: Evaluator, lookup: Scalar, vec: Scalar[], mode: number): number {
  if (mode === 0) {
    const test = typeof lookup === 'string' && /[*?]/.test(lookup) ? criterion(lookup) : (v: Scalar) => cmp(v, lookup) === 0 && (typeof v === typeof lookup || v === null && lookup === '');
    for (let i = 0; i < vec.length; i++) if (test(vec[i])) return i;
    return -1;
  }
  // approximate (sorted): last value <= lookup (mode 1) or >= (mode -1)
  let best = -1;
  for (let i = 0; i < vec.length; i++) {
    const v = vec[i];
    if (v === null || typeof v !== typeof lookup) continue;
    const c = cmp(v, lookup);
    if (mode === 1 ? c <= 0 : c >= 0) best = i; else if (mode === 1) break;
  }
  void e;
  return best;
}

const FUNCS: Record<string, Fn> = {
  SUM: (e, a) => e.all(a).reduce((s, x) => s + x, 0),
  PRODUCT: (e, a) => e.all(a).reduce((s, x) => s * x, 1),
  AVERAGE: (e, a) => { const n = e.all(a); return n.length ? n.reduce((s, x) => s + x, 0) / n.length : ERR.div0; },
  AVERAGEA: (e, a) => { const n = e.numsA(a); return n.length ? n.reduce((s, x) => s + x, 0) / n.length : ERR.div0; },
  MIN: (e, a) => { const n = e.all(a); return n.length ? Math.min(...n) : 0; },
  MAX: (e, a) => { const n = e.all(a); return n.length ? Math.max(...n) : 0; },
  MEDIAN: (e, a) => { const n = e.all(a).sort((x, y) => x - y); if (!n.length) return ERR.num; const m = n.length >> 1; return n.length % 2 ? n[m] : (n[m - 1] + n[m]) / 2; },
  COUNT: (e, a) => e.flatArgs(a).filter(x => typeof x === 'number').length,
  COUNTA: (e, a) => e.flatArgs(a).filter(x => x !== null && x !== '').length,
  COUNTBLANK: (e, a) => e.flatArgs(a).filter(x => x === null || x === '').length,
  ABS: (e, a) => Math.abs(e.num(a[0])),
  INT: (e, a) => Math.floor(e.num(a[0])),
  TRUNC: (e, a) => roundTo(e.num(a[0]), e.num(a[1], 0), 'd'),
  ROUND: (e, a) => roundTo(e.num(a[0]), e.num(a[1]), 'n'),
  ROUNDUP: (e, a) => roundTo(e.num(a[0]), e.num(a[1]), 'u'),
  ROUNDDOWN: (e, a) => roundTo(e.num(a[0]), e.num(a[1]), 'd'),
  MROUND: (e, a) => { const x = e.num(a[0]), m = e.num(a[1]); return m === 0 ? 0 : Math.round(x / m) * m; },
  CEILING: (e, a) => { const x = e.num(a[0]), m = e.num(a[1], 1); return m === 0 ? 0 : Math.ceil(x / m - 1e-12) * m; },
  'CEILING.MATH': (e, a) => { const x = e.num(a[0]), m = e.num(a[1], 1); return m === 0 ? 0 : Math.ceil(x / m - 1e-12) * m; },
  FLOOR: (e, a) => { const x = e.num(a[0]), m = e.num(a[1], 1); return m === 0 ? ERR.div0 : Math.floor(x / m + 1e-12) * m; },
  'FLOOR.MATH': (e, a) => { const x = e.num(a[0]), m = e.num(a[1], 1); return m === 0 ? 0 : Math.floor(x / m + 1e-12) * m; },
  MOD: (e, a) => { const x = e.num(a[0]), d = e.num(a[1]); if (d === 0) return ERR.div0; return x - d * Math.floor(x / d); },
  POWER: (e, a) => Math.pow(e.num(a[0]), e.num(a[1])),
  SQRT: (e, a) => { const x = e.num(a[0]); return x < 0 ? ERR.num : Math.sqrt(x); },
  EXP: (e, a) => Math.exp(e.num(a[0])),
  LN: (e, a) => { const x = e.num(a[0]); return x <= 0 ? ERR.num : Math.log(x); },
  LOG10: (e, a) => Math.log10(e.num(a[0])),
  LOG: (e, a) => Math.log(e.num(a[0])) / Math.log(e.num(a[1], 10)),
  PI: () => Math.PI,
  SIGN: (e, a) => Math.sign(e.num(a[0])),
  RAND: () => 0.5,
  IF: (e, a) => { const c = e.bool(a[0]); const n = c ? a[1] : a[2]; if (!n) return c ? true : false; if (n.t === 'missing') return 0; return e.eval(n); },
  IFS: (e, a) => { for (let i = 0; i + 1 < a.length; i += 2) if (e.bool(a[i])) return e.eval(a[i + 1]); return ERR.na; },
  IFERROR: (e, a) => { const v = e.val(a[0]); return isErr(v) ? e.eval(a[1]) : v; },
  IFNA: (e, a) => { const v = e.val(a[0]); return isErr(v) && v.error === '#N/A' ? e.eval(a[1]) : v; },
  AND: (e, a) => { const v = e.flatArgs(a).filter(x => x !== null && typeof x !== 'string'); if (!v.length) return ERR.value; for (const x of v) { const b = toBool(x); if (isErr(b)) return b; if (!b) return false; } return true; },
  OR: (e, a) => { const v = e.flatArgs(a).filter(x => x !== null && typeof x !== 'string'); if (!v.length) return ERR.value; for (const x of v) { const b = toBool(x); if (isErr(b)) return b; if (b) return true; } return false; },
  XOR: (e, a) => e.flatArgs(a).filter(x => x === true || (typeof x === 'number' && x !== 0)).length % 2 === 1,
  NOT: (e, a) => !e.bool(a[0]),
  TRUE: () => true,
  FALSE: () => false,
  NA: () => ERR.na,
  ISBLANK: (e, a) => e.val(a[0]) === null,
  ISNUMBER: (e, a) => typeof e.val(a[0]) === 'number',
  ISTEXT: (e, a) => typeof e.val(a[0]) === 'string',
  ISNONTEXT: (e, a) => typeof e.val(a[0]) !== 'string',
  ISLOGICAL: (e, a) => typeof e.val(a[0]) === 'boolean',
  ISERROR: (e, a) => isErr(e.val(a[0])),
  ISERR: (e, a) => { const v = e.val(a[0]); return isErr(v) && v.error !== '#N/A'; },
  ISNA: (e, a) => { const v = e.val(a[0]); return isErr(v) && v.error === '#N/A'; },
  ISEVEN: (e, a) => Math.floor(e.num(a[0])) % 2 === 0,
  ISODD: (e, a) => Math.abs(Math.floor(e.num(a[0])) % 2) === 1,
  CONCATENATE: (e, a) => a.map(x => e.str(x)).join(''),
  CONCAT: (e, a) => e.flatArgs(a).map(x => { const s = toStr(x); if (isErr(s)) throw s; return s; }).join(''),
  TEXTJOIN: (e, a) => { const d = e.str(a[0]); const skip = e.bool(a[1]); return a.slice(2).flatMap(x => e.arr(x).flat()).map(x => { const s = toStr(x); if (isErr(s)) throw s; return s; }).filter(s => !(skip && s === '')).join(d); },
  LEN: (e, a) => e.str(a[0]).length,
  LEFT: (e, a) => e.str(a[0]).slice(0, e.num(a[1], 1)),
  RIGHT: (e, a) => { const s = e.str(a[0]); const n = e.num(a[1], 1); return n <= 0 ? '' : s.slice(-n); },
  MID: (e, a) => { const s = e.str(a[0]); const st = e.num(a[1]); const n = e.num(a[2]); if (st < 1 || n < 0) return ERR.value; return s.substr(st - 1, n); },
  UPPER: (e, a) => e.str(a[0]).toUpperCase(),
  LOWER: (e, a) => e.str(a[0]).toLowerCase(),
  PROPER: (e, a) => e.str(a[0]).toLowerCase().replace(/(^|[^a-z])([a-z])/g, (m, p: string, c: string) => p + c.toUpperCase()),
  TRIM: (e, a) => e.str(a[0]).replace(/ +/g, ' ').trim(),
  CLEAN: (e, a) => e.str(a[0]).replace(/[\x00-\x1f]/g, ''),
  REPT: (e, a) => e.str(a[0]).repeat(Math.max(0, Math.floor(e.num(a[1])))),
  EXACT: (e, a) => e.str(a[0]) === e.str(a[1]),
  FIND: (e, a) => { const i = e.str(a[1]).indexOf(e.str(a[0]), e.num(a[2], 1) - 1); return i < 0 ? ERR.value : i + 1; },
  SEARCH: (e, a) => { const i = e.str(a[1]).toLowerCase().indexOf(e.str(a[0]).toLowerCase(), e.num(a[2], 1) - 1); return i < 0 ? ERR.value : i + 1; },
  SUBSTITUTE: (e, a) => {
    const s = e.str(a[0]), o = e.str(a[1]), nw = e.str(a[2]);
    if (!o) return s;
    if (a[3] && a[3].t !== 'missing') { const k = e.num(a[3]); let idx = -1; for (let i = 0; i < k; i++) { idx = s.indexOf(o, idx + 1); if (idx < 0) return s; } return s.slice(0, idx) + nw + s.slice(idx + o.length); }
    return s.split(o).join(nw);
  },
  REPLACE: (e, a) => { const s = e.str(a[0]); const st = e.num(a[1]) - 1; const n = e.num(a[2]); return s.slice(0, st) + e.str(a[3]) + s.slice(st + n); },
  VALUE: (e, a) => { const v = e.val(a[0]); const n = toNum(typeof v === 'string' ? v : (v as Scalar)); return n; },
  TEXT: (e, a) => { const v = e.val(a[0]); const code = e.str(a[1]); const n = typeof v === 'string' ? toNum(v) : v; return formatValue(isErr(n) ? v : n, code).text; },
  FIXED: (e, a) => { const x = roundTo(e.num(a[0]), e.num(a[1], 2), 'n'); const d = Math.max(0, e.num(a[1], 2)); const s = x.toFixed(d); return e.bool(a[2], false) ? s : s.replace(/\B(?=(\d{3})+(?!\d))/g, ','); },
  DOLLAR: (e, a) => formatValue(e.num(a[0]), BUILTIN_FORMATS[7]).text,
  CHAR: (e, a) => String.fromCharCode(e.num(a[0])),
  CODE: (e, a) => e.str(a[0]).charCodeAt(0) || ERR.value,
  N: (e, a) => { const v = e.val(a[0]); return typeof v === 'number' ? v : v === true ? 1 : 0; },
  T: (e, a) => { const v = e.val(a[0]); return typeof v === 'string' ? v : ''; },
  TODAY: e => { const d = scenarioDate(e.ctx.nowMs); return partsToSerial(d.y, d.m, d.d); },
  NOW: e => { const d = scenarioDate(e.ctx.nowMs); return partsToSerial(d.y, d.m, d.d) + (d.H * 3600 + d.M * 60 + d.S) / 86400; },
  DATE: (e, a) => { let y = e.num(a[0]); const m = e.num(a[1]), d = e.num(a[2]); if (y < 1900) y += 1900; const dt = new Date(Date.UTC(y, m - 1, d)); return partsToSerial(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate()); },
  TIME: (e, a) => (e.num(a[0]) * 3600 + e.num(a[1]) * 60 + e.num(a[2])) / 86400 % 1,
  YEAR: (e, a) => serialToParts(dateArg(e, a[0])).y,
  MONTH: (e, a) => serialToParts(dateArg(e, a[0])).m,
  DAY: (e, a) => serialToParts(dateArg(e, a[0])).d,
  HOUR: (e, a) => serialToParts(dateArg(e, a[0])).H,
  MINUTE: (e, a) => serialToParts(dateArg(e, a[0])).M,
  SECOND: (e, a) => serialToParts(dateArg(e, a[0])).S,
  WEEKDAY: (e, a) => { const d = serialToParts(dateArg(e, a[0])).dow; const t = e.num(a[1], 1); return t === 2 ? ((d + 6) % 7) + 1 : t === 3 ? (d + 6) % 7 : d + 1; },
  DATEVALUE: (e, a) => { const v = parseDateText(e.str(a[0])); return v ?? ERR.value; },
  EDATE: (e, a) => { const p = serialToParts(dateArg(e, a[0])); const dt = new Date(Date.UTC(p.y, p.m - 1 + e.num(a[1]), 1)); const last = new Date(Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth() + 1, 0)).getUTCDate(); return partsToSerial(dt.getUTCFullYear(), dt.getUTCMonth() + 1, Math.min(p.d, last)); },
  EOMONTH: (e, a) => { const p = serialToParts(dateArg(e, a[0])); const dt = new Date(Date.UTC(p.y, p.m + e.num(a[1]), 0)); return partsToSerial(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate()); },
  DAYS: (e, a) => Math.floor(dateArg(e, a[0])) - Math.floor(dateArg(e, a[1])),
  DATEDIF: (e, a) => {
    const s = serialToParts(dateArg(e, a[0])), t = serialToParts(dateArg(e, a[1])); const u = e.str(a[2]).toUpperCase();
    if (dateArg(e, a[0]) > dateArg(e, a[1])) return ERR.num;
    let months = (t.y - s.y) * 12 + (t.m - s.m); if (t.d < s.d) months--;
    if (u === 'Y') return Math.floor(months / 12);
    if (u === 'M') return months;
    if (u === 'D') return Math.floor(dateArg(e, a[1])) - Math.floor(dateArg(e, a[0]));
    if (u === 'YM') return months % 12;
    return ERR.num;
  },
  YEARFRAC: (e, a) => Math.abs(dateArg(e, a[1]) - dateArg(e, a[0])) / 365,
  NETWORKDAYS: (e, a) => { let s = Math.floor(dateArg(e, a[0])), t = Math.floor(dateArg(e, a[1])); const sign = s <= t ? 1 : -1; if (s > t) [s, t] = [t, s]; let n = 0; for (let d = s; d <= t; d++) { const w = serialToParts(d).dow; if (w !== 0 && w !== 6) n++; } return n * sign; },
  ROW: (e, a) => a.length ? e.ref(a[0]).range.r1 + 1 : e.ctx.r + 1,
  COLUMN: (e, a) => a.length ? e.ref(a[0]).range.c1 + 1 : e.ctx.c + 1,
  ROWS: (e, a) => e.arr(a[0]).length,
  COLUMNS: (e, a) => e.arr(a[0])[0]?.length ?? 0,
  SUMIF: (e, a) => {
    const range = e.arr(a[0]); const test = criterion(e.val(a[1])); const sum = a[2] ? e.arr(a[2]) : range;
    let s = 0; range.forEach((row, i) => row.forEach((v, j) => { if (test(v)) { const x = sum[i]?.[j]; if (typeof x === 'number') s += x; } })); return s;
  },
  SUMIFS: (e, a) => {
    const sum = e.arr(a[0]); const conds = pairs(e, a.slice(1));
    let s = 0; sum.forEach((row, i) => row.forEach((x, j) => { if (conds.every(([rg, t]) => t(rg[i]?.[j] ?? null)) && typeof x === 'number') s += x; })); return s;
  },
  COUNTIF: (e, a) => { const range = e.arr(a[0]); const test = criterion(e.val(a[1])); return range.flat().filter(v => test(v)).length; },
  COUNTIFS: (e, a) => { const conds = pairs(e, a); const first = conds[0]?.[0] ?? []; let n = 0; first.forEach((row, i) => row.forEach((_, j) => { if (conds.every(([rg, t]) => t(rg[i]?.[j] ?? null))) n++; })); return n; },
  AVERAGEIF: (e, a) => {
    const range = e.arr(a[0]); const test = criterion(e.val(a[1])); const sum = a[2] ? e.arr(a[2]) : range;
    let s = 0, n = 0; range.forEach((row, i) => row.forEach((v, j) => { if (test(v)) { const x = sum[i]?.[j]; if (typeof x === 'number') { s += x; n++; } } })); return n ? s / n : ERR.div0;
  },
  AVERAGEIFS: (e, a) => {
    const sum = e.arr(a[0]); const conds = pairs(e, a.slice(1));
    let s = 0, n = 0; sum.forEach((row, i) => row.forEach((x, j) => { if (conds.every(([rg, t]) => t(rg[i]?.[j] ?? null)) && typeof x === 'number') { s += x; n++; } })); return n ? s / n : ERR.div0;
  },
  MAXIFS: (e, a) => { const sum = e.arr(a[0]); const conds = pairs(e, a.slice(1)); const v: number[] = []; sum.forEach((row, i) => row.forEach((x, j) => { if (conds.every(([rg, t]) => t(rg[i]?.[j] ?? null)) && typeof x === 'number') v.push(x); })); return v.length ? Math.max(...v) : 0; },
  MINIFS: (e, a) => { const sum = e.arr(a[0]); const conds = pairs(e, a.slice(1)); const v: number[] = []; sum.forEach((row, i) => row.forEach((x, j) => { if (conds.every(([rg, t]) => t(rg[i]?.[j] ?? null)) && typeof x === 'number') v.push(x); })); return v.length ? Math.min(...v) : 0; },
  SUMPRODUCT: (e, a) => {
    const ms = a.map(x => e.arr(x));
    const rows = ms[0].length, cols = ms[0][0]?.length ?? 0;
    if (ms.some(m => m.length !== rows || (m[0]?.length ?? 0) !== cols)) return ERR.value;
    let s = 0;
    for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) {
      let p = 1;
      for (const m of ms) { const v = m[i][j]; if (isErr(v)) return v; p *= typeof v === 'number' ? v : typeof v === 'boolean' ? Number(v) : 0; }
      s += p;
    }
    return s;
  },
  VLOOKUP: (e, a) => {
    const look = e.val(a[0]); const tbl = e.arr(a[1]); const col = e.num(a[2]); const approx = a[3] && a[3].t !== 'missing' ? e.bool(a[3]) : true;
    if (col < 1 || col > (tbl[0]?.length ?? 0)) return ERR.ref;
    const i = lookupMatch(e, look, tbl.map(r => r[0]), approx ? 1 : 0);
    return i < 0 ? ERR.na : tbl[i][col - 1];
  },
  HLOOKUP: (e, a) => {
    const look = e.val(a[0]); const tbl = e.arr(a[1]); const row = e.num(a[2]); const approx = a[3] && a[3].t !== 'missing' ? e.bool(a[3]) : true;
    if (row < 1 || row > tbl.length) return ERR.ref;
    const i = lookupMatch(e, look, tbl[0], approx ? 1 : 0);
    return i < 0 ? ERR.na : tbl[row - 1][i];
  },
  MATCH: (e, a) => {
    const look = e.val(a[0]); const m = e.arr(a[1]); const mode = e.num(a[2], 1);
    const vec = m.length === 1 ? m[0] : m.map(r => r[0]);
    const i = lookupMatch(e, look, vec, mode === 0 ? 0 : mode > 0 ? 1 : -1);
    return i < 0 ? ERR.na : i + 1;
  },
  XMATCH: (e, a) => { const look = e.val(a[0]); const m = e.arr(a[1]); const vec = m.length === 1 ? m[0] : m.map(r => r[0]); const i = lookupMatch(e, look, vec, 0); return i < 0 ? ERR.na : i + 1; },
  INDEX: (e, a) => {
    const m = e.arr(a[0]); let r = e.num(a[1], 0); let c = e.num(a[2], 0);
    if (m.length === 1 && (!a[2] || a[2].t === 'missing')) { c = r; r = 1; }
    if (r === 0 && c === 0) return m;
    if (r === 0) return m.map(row => [row[c - 1]]);
    if (c === 0) c = m[0].length === 1 ? 1 : 0;
    if (c === 0) return [m[r - 1]];
    const v = m[r - 1]?.[c - 1];
    return v === undefined ? ERR.ref : v;
  },
  XLOOKUP: (e, a) => {
    const look = e.val(a[0]); const lm = e.arr(a[1]); const rm = e.arr(a[2]);
    const vertical = lm.length > 1 || (lm[0]?.length ?? 0) === 1;
    const vec = vertical ? lm.map(r => r[0]) : lm[0];
    const i = lookupMatch(e, look, vec, 0);
    if (i < 0) return a[3] && a[3].t !== 'missing' ? e.eval(a[3]) : ERR.na;
    return vertical ? (rm[i].length === 1 ? rm[i][0] : [rm[i]]) : (rm.length === 1 ? rm[0][i] : rm.map(r => [r[i]]));
  },
  CHOOSE: (e, a) => { const i = e.num(a[0]); if (i < 1 || i >= a.length) return ERR.value; return e.eval(a[Math.floor(i)]); },
  SWITCH: (e, a) => { const v = e.val(a[0]); let i = 1; for (; i + 1 < a.length; i += 2) if (cmp(e.val(a[i]), v) === 0) return e.eval(a[i + 1]); return i < a.length ? e.eval(a[i]) : ERR.na; },
  LARGE: (e, a) => { const n = e.all([a[0]]).sort((x, y) => y - x); const k = e.num(a[1]); return n[k - 1] ?? ERR.num; },
  SMALL: (e, a) => { const n = e.all([a[0]]).sort((x, y) => x - y); const k = e.num(a[1]); return n[k - 1] ?? ERR.num; },
  RANK: (e, a) => { const x = e.num(a[0]); const n = e.all([a[1]]); const asc = e.num(a[2], 0) !== 0; const i = n.filter(v => (asc ? v < x : v > x)).length; return n.includes(x) ? i + 1 : ERR.na; },
  'RANK.EQ': (e, a) => FUNCS.RANK(e, a),
  STDEV: (e, a) => { const n = e.all(a); if (n.length < 2) return ERR.div0; const m = n.reduce((s, x) => s + x, 0) / n.length; return Math.sqrt(n.reduce((s, x) => s + (x - m) ** 2, 0) / (n.length - 1)); },
  'STDEV.S': (e, a) => FUNCS.STDEV(e, a),
  PMT: (e, a) => { const r = e.num(a[0]), n = e.num(a[1]), pv = e.num(a[2]), fv = e.num(a[3], 0), t = e.num(a[4], 0); if (r === 0) return -(pv + fv) / n; const f = Math.pow(1 + r, n); return -(r * (pv * f + fv)) / ((1 + r * t) * (f - 1)); },
  FV: (e, a) => { const r = e.num(a[0]), n = e.num(a[1]), p = e.num(a[2]), pv = e.num(a[3], 0), t = e.num(a[4], 0); if (r === 0) return -(pv + p * n); const f = Math.pow(1 + r, n); return -(pv * f + p * (1 + r * t) * (f - 1) / r); },
  PV: (e, a) => { const r = e.num(a[0]), n = e.num(a[1]), p = e.num(a[2]), fv = e.num(a[3], 0), t = e.num(a[4], 0); if (r === 0) return -(fv + p * n); const f = Math.pow(1 + r, n); return -(fv + p * (1 + r * t) * (f - 1) / r) / f; },
  NPV: (e, a) => { const r = e.num(a[0]); return e.all(a.slice(1)).reduce((s, x, i) => s + x / Math.pow(1 + r, i + 1), 0); },
  TRANSPOSE: (e, a) => { const m = e.arr(a[0]); return (m[0] ?? []).map((_, j) => m.map(r => r[j])); },
  UNIQUE: (e, a) => { const m = e.arr(a[0]); const seen = new Set<string>(); return m.filter(r => { const k = JSON.stringify(r); if (seen.has(k)) return false; seen.add(k); return true; }); },
  HYPERLINK: (e, a) => (a[1] ? e.val(a[1]) : e.val(a[0])),
};

function pairs(e: Evaluator, a: Node[]): [Matrix, (v: Scalar) => boolean][] {
  const out: [Matrix, (v: Scalar) => boolean][] = [];
  for (let i = 0; i + 1 < a.length; i += 2) out.push([e.arr(a[i]), criterion(e.val(a[i + 1]))]);
  return out;
}

function dateArg(e: Evaluator, n: Node): number {
  const v = e.val(n);
  if (typeof v === 'string') { const d = parseDateText(v); if (d === null) throw ERR.value; return d; }
  const x = toNum(v);
  if (isErr(x)) throw x;
  return x;
}

function scenarioDate(ms: number) {
  // America/Los_Angeles wall clock of the scenario instant
  const f = new Intl.DateTimeFormat('en-US', { timeZone: 'America/Los_Angeles', year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: 'numeric', second: 'numeric', hourCycle: 'h23' });
  const parts = Object.fromEntries(f.formatToParts(new Date(ms)).map(p => [p.type, p.value]));
  return { y: Number(parts.year), m: Number(parts.month), d: Number(parts.day), H: Number(parts.hour), M: Number(parts.minute), S: Number(parts.second) };
}

export const FUNCTION_NAMES = Object.keys(FUNCS).sort();

/** references a formula reads, for dependency tracking (sheet-qualified ranges) */
export function formulaRefs(n: Node, sheet: string, out: { sheet: string; a: string; b?: string }[] = []) {
  switch (n.t) {
    case 'ref': out.push({ sheet: n.sheet ?? sheet, a: n.a, b: n.b }); break;
    case 'fn': n.args.forEach(x => formulaRefs(x, sheet, out)); break;
    case 'un': case 'pct': formulaRefs(n.e, sheet, out); break;
    case 'bin': formulaRefs(n.l, sheet, out); formulaRefs(n.r, sheet, out); break;
    case 'arr': n.rows.flat().forEach(x => formulaRefs(x, sheet, out)); break;
  }
  return out;
}
