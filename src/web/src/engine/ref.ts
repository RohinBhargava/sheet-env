// A1 references. Rows and columns are 0-based internally.

export interface Addr { r: number; c: number }
export interface Range { r1: number; c1: number; r2: number; c2: number }

export const MAX_ROW = 1048575;
export const MAX_COL = 16383;

export function colName(c: number): string {
  let s = '';
  let n = c + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
export function colIndex(name: string): number {
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
export function addrName(r: number, c: number): string { return colName(c) + (r + 1); }

export function parseAddr(s: string): Addr | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(s.trim());
  if (!m) return null;
  const c = colIndex(m[1]); const r = Number(m[2]) - 1;
  if (r < 0 || r > MAX_ROW || c < 0 || c > MAX_COL) return null;
  return { r, c };
}

/** "A1", "A1:C3", "A:C", "2:5" */
export function parseRange(s: string): Range | null {
  const t = s.trim().replace(/\$/g, '');
  const parts = t.split(':');
  if (parts.length === 1) { const a = parseAddr(parts[0]); return a ? { r1: a.r, c1: a.c, r2: a.r, c2: a.c } : null; }
  if (parts.length !== 2) return null;
  const [x, y] = parts;
  if (/^[A-Za-z]{1,3}$/.test(x) && /^[A-Za-z]{1,3}$/.test(y)) { const c1 = colIndex(x), c2 = colIndex(y); return norm({ r1: 0, c1, r2: MAX_ROW, c2 }); }
  if (/^\d+$/.test(x) && /^\d+$/.test(y)) return norm({ r1: Number(x) - 1, c1: 0, r2: Number(y) - 1, c2: MAX_COL });
  const a = parseAddr(x), b = parseAddr(y);
  return a && b ? norm({ r1: a.r, c1: a.c, r2: b.r, c2: b.c }) : null;
}
export function norm(g: Range): Range {
  return { r1: Math.min(g.r1, g.r2), c1: Math.min(g.c1, g.c2), r2: Math.max(g.r1, g.r2), c2: Math.max(g.c1, g.c2) };
}
export function rangeName(g: Range): string {
  const a = addrName(g.r1, g.c1); const b = addrName(g.r2, g.c2);
  return a === b ? a : `${a}:${b}`;
}
export function inRange(g: Range, r: number, c: number) { return r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2; }
export function key(r: number, c: number) { return r * 16384 + c; }
export function unkey(k: number): Addr { return { r: Math.floor(k / 16384), c: k % 16384 }; }
