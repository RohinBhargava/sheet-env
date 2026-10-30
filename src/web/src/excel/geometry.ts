// Pixel geometry for the grid: column widths (Excel character widths at Calibri 11, max digit width
// 7 px), row heights (points at 96 dpi), visible rows/columns for a scroll position and frozen panes.
import type { Workbook } from '../engine/index.ts';

export const HEAD_H = 20;
export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

export function charsToPx(w: number): number {
  // Excel: pixels = Truncate(((256 * width + Truncate(128 / MDW)) / 256) * MDW), MDW = 7 for Calibri 11
  return Math.trunc(((256 * w + Math.trunc(128 / 7)) / 256) * 7);
}
export function pxToChars(px: number): number {
  return Math.round(((px - 5) / 7) * 100 + 0.5) / 100 + 0.71;
}

export interface Band { i: number; pos: number; size: number }

export class Geometry {
  wb: Workbook;
  si: number;
  zoom: number;
  constructor(wb: Workbook, si: number, zoom: number) { this.wb = wb; this.si = si; this.zoom = zoom; }
  colW(c: number): number { return this.wb.colHidden(this.si, c) ? 0 : Math.round(charsToPx(this.wb.colWidthChars(this.si, c)) * this.zoom); }
  rowH(r: number): number { return this.wb.rowHidden(this.si, r) ? 0 : Math.round(this.wb.rowHeightPt(this.si, r) * 96 / 72 * this.zoom); }
  rowHeadW(lastRow: number): number { const d = String(lastRow + 1).length; return Math.round((d <= 2 ? 26 : d * 7 + 12) * Math.min(1, this.zoom)); }
  headH(): number { return Math.round(HEAD_H * this.zoom); }

  /** rows to draw: frozen rows, then rows from `top` until `height` is filled (+1 partial) */
  rows(top: number, height: number, frozen: number): Band[] {
    const out: Band[] = [];
    let y = 0;
    for (let r = 0; r < frozen; r++) { const h = this.rowH(r); out.push({ i: r, pos: y, size: h }); y += h; }
    for (let r = Math.max(top, frozen); r < MAX_ROWS && y < height; r++) { const h = this.rowH(r); out.push({ i: r, pos: y, size: h }); y += h; }
    return out;
  }
  cols(left: number, width: number, frozen: number): Band[] {
    const out: Band[] = [];
    let x = 0;
    for (let c = 0; c < frozen; c++) { const w = this.colW(c); out.push({ i: c, pos: x, size: w }); x += w; }
    for (let c = Math.max(left, frozen); c < MAX_COLS && x < width; c++) { const w = this.colW(c); out.push({ i: c, pos: x, size: w }); x += w; }
    return out;
  }
}

/** cell text size in px: Excel (GDI) renders at a whole-pixel em, so 11 pt = 15 px at 96 dpi */
export function cellFontPx(pt: number, zoom: number): number { return Math.max(1, Math.round(pt * 96 / 72 * zoom)); }

/** CSS font-family for a workbook font name (substitutes are aliased in the shell's fonts.css) */
export function cssFamily(name: string): string {
  const n = (name || 'Calibri').trim();
  const map: Record<string, string> = {
    'Calibri': 'Calibri, Carlito', 'Calibri Light': 'Calibri, Carlito', 'Aptos': '"Aptos", "Isomet Sans"', 'Aptos Display': '"Aptos Display", "Isomet Sans"', 'Aptos Narrow': 'Calibri, Carlito',
    'Cambria': 'Cambria, Caladea', 'Arial': 'Arial, "Liberation Sans"', 'Helvetica': 'Arial, "Liberation Sans"', 'Times New Roman': '"Times New Roman", "Liberation Serif"',
    'Segoe UI': 'Selawik', 'Consolas': 'Consolas, "Liberation Mono"', 'Courier New': '"Liberation Mono"',
  };
  return `${map[n] ?? `"${n}", Calibri, Carlito`}, sans-serif`;
}

let measureCtx: CanvasRenderingContext2D | null = null;
export function textWidth(text: string, font: string): number {
  if (!measureCtx) measureCtx = document.createElement('canvas').getContext('2d');
  if (!measureCtx) return text.length * 7;
  measureCtx.font = font;
  // cells draw unkerned (excel.css .xl-cell): measure the same way, or overflow and #### decisions drift
  (measureCtx as CanvasRenderingContext2D & { fontKerning: string }).fontKerning = 'none';
  return measureCtx.measureText(text).width;
}
