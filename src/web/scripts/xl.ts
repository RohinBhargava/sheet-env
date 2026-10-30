// Small helpers over the xlsx engine for building and reading workbooks from scripts (seed, solve, checks).
import { Workbook, parseAddr, parseRange, type Scalar } from '../src/engine/index.ts';
import { blankWorkbook } from '../src/engine/blank.ts';

export const FMT = {
  usd: '#,##0.00;(#,##0.00);"-"',
  usd0: '#,##0;(#,##0);"-"',
  pct: '0.00%',
  pct4: '0.0000%',
  date: 'mm/dd/yyyy',
  int: '0',
};

/** Excel serial for a calendar date */
export function serial(y: number, m: number, d: number): number {
  return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
}

export class Book {
  wb: Workbook;
  constructor(wb?: Workbook) { this.wb = wb ?? Workbook.load(blankWorkbook()); }
  static open(bytes: Uint8Array) { return new Book(Workbook.load(bytes)); }

  /** first sheet gets `first`'s name; returns the sheet index */
  sheet(name: string): number {
    const i = this.wb.sheetIndex(name);
    if (i >= 0) return i;
    if (this.wb.sheets.length === 1 && this.wb.sheets[0].name === 'Sheet1' && !this.wb.sheets[0].cells.size) { this.wb.renameSheet(0, name); return 0; }
    return this.wb.addSheet(this.wb.sheets.length, name);
  }

  si(name: string): number {
    const i = this.wb.sheetIndex(name);
    if (i < 0) throw new Error(`no sheet ${name}`);
    return i;
  }

  /** text, number, formula ("=...") or null (clear) */
  set(sheet: string, a1: string, v: string | number | null) {
    const si = this.si(sheet);
    const a = parseAddr(a1);
    if (!a) throw new Error(`bad address ${a1}`);
    if (v === null) this.wb.setInput(si, a.r, a.c, '');
    else if (typeof v === 'number') this.wb.setCellRaw(si, a.r, a.c, v);
    else if (v.startsWith('=')) this.wb.setInput(si, a.r, a.c, v);
    else this.wb.setCellRaw(si, a.r, a.c, v);
  }

  /** write a row of values starting at a1 */
  row(sheet: string, a1: string, values: (string | number | null | undefined)[]) {
    const a = parseAddr(a1)!;
    values.forEach((v, i) => { if (v !== undefined) this.set(sheet, colA1(a.c + i, a.r), v); });
  }

  fmt(sheet: string, range: string, patch: Parameters<Workbook['formatRange']>[2]) {
    const g = parseRange(range);
    if (!g) throw new Error(`bad range ${range}`);
    this.wb.formatRange(this.si(sheet), g, patch);
  }

  widths(sheet: string, widths: Record<string, number>) {
    const si = this.si(sheet);
    for (const [col, w] of Object.entries(widths)) {
      const g = parseRange(`${col}1`)!;
      this.wb.setColWidth(si, g.c1, g.c1, w);
    }
  }

  freeze(sheet: string, rows: number, cols = 0) { this.wb.setFreeze(this.si(sheet), rows, cols); }

  /** cached value as written in the file (what a non-calculating reader sees) */
  cached(sheet: string, a1: string): Scalar {
    const a = parseAddr(a1)!;
    return this.wb.cell(this.si(sheet), a.r, a.c)?.v ?? null;
  }
  /** value recomputed by the engine */
  computed(sheet: string, a1: string): Scalar {
    const a = parseAddr(a1)!;
    return this.wb.value(this.si(sheet), a.r, a.c);
  }

  save(): Uint8Array {
    this.wb.activeTab = 0;
    return this.wb.save();
  }
}

function colA1(c: number, r: number): string {
  let s = '', n = c + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return `${s}${r + 1}`;
}
