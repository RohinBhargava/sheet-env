// The worksheet grid: headers, gridlines, cell fills/borders/text (with Excel's text overflow into
// empty neighbours and #### for numbers that do not fit), merged cells, frozen panes, selection,
// marching ants, the in-cell editor, filter buttons and mouse handling (select, extend, resize).
import { useEffect, useLayoutEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import type { Workbook, Range, Addr } from '../engine/index.ts';
import { colName, addrName, isErr, MAX_ROW, MAX_COL } from '../engine/index.ts';
import { Geometry, cssFamily, textWidth, cellFontPx, type Band } from './geometry';
import { Icon } from './icons';
import { BRAND } from '../util/brand';
import { wheelNotches } from '../util/wheel';

export type SelKind = 'cells' | 'cols' | 'rows' | 'all';
export interface Sel { active: Addr; anchor: Addr; range: Range; kind: SelKind }
export interface EditorState { r: number; c: number; text: string; mode: 'enter' | 'edit'; caret?: number }

export interface GridProps {
  wb: Workbook;
  si: number;
  zoom: number;
  version: number;
  top: number;
  left: number;
  setScroll(top: number, left: number): void;
  sel: Sel;
  onSelect(sel: Sel): void;
  editor: EditorState | null;
  onEditorText(text: string, caret: number): void;
  onEditorKey(e: React.KeyboardEvent<HTMLTextAreaElement>): void;
  editorRef: React.MutableRefObject<HTMLTextAreaElement | null>;
  showGrid: boolean;
  showHeadings: boolean;
  showFormulas: boolean;
  ants: Range | null;
  filter: Range | null;
  onFilterMenu(c: number, el: HTMLElement): void;
  onDoubleClick(r: number, c: number): void;
  onContext(kind: 'cell' | 'col' | 'row' | 'all', x: number, y: number): void;
  onResizeCol(c: number, px: number): void;
  onResizeRow(r: number, px: number): void;
  onAutofitCol(c: number): void;
  onPointRef(ref: string, extend: boolean): void;
  pointing: boolean;
  onViewport(rows: number, cols: number, thumbs?: ScrollThumbs): void;
  /** show the Quick Analysis button under a multi-cell selection */
  quickAnalysis?: boolean;
  /** floating Copilot button at the bottom right of the grid */
  fab?: boolean;
}

/** Excel draws digits on whole-pixel advances (Calibri 11: the 7 px "maximum digit width" its column widths
 *  are based on); our substitute font's digits are ~7.6 px at 15 px, so digit runs get a small negative tracking. */
const dlsCache = new Map<string, number>();
function digitSpacing(css: string, truePx: number, family: string, font: { b?: boolean; i?: boolean }): number {
  const k = `${css}|${truePx}`;
  let v = dlsCache.get(k);
  if (v === undefined) {
    const at = textWidth('0123456789', css) / 10;
    const ideal = textWidth('0123456789', `${font.i ? 'italic ' : ''}${font.b ? '700 ' : ''}${truePx}px ${family}`) / 10;
    v = Math.round(ideal) - at;
    if (Math.abs(v) > 1.5) v = 0;
    // canvas measures with a fallback face until the web font has loaded: cache only real measurements
    if (typeof document === 'undefined' || !document.fonts || document.fonts.check(css)) dlsCache.set(k, v);
  }
  return v;
}
function cellTextWidth(text: string, css: string, dls: number): number {
  if (!dls) return textWidth(text, css);
  const n = (text.match(/\d/g) ?? []).length;
  return textWidth(text, css) + n * dls;
}
function digitRuns(text: string, dls: number): ReactNode {
  if (!dls || !/\d/.test(text)) return text;
  const parts = text.split(/(\d+)/);
  // one inline box: the cell is a flex container, which would drop a space-only text node between two digit runs
  return <span>{parts.map((t, i) => (i % 2 ? <span key={i} style={{ letterSpacing: dls }}>{t}</span> : t || null))}</span>;
}

export function mergeRect(rows: Band[], cols: Band[], g: Range, geo: Geometry) {
  const rFirst = rows.find(b => b.i >= g.r1 && b.i <= g.r2);
  const cFirst = cols.find(b => b.i >= g.c1 && b.i <= g.c2);
  if (!rFirst || !cFirst) return null;
  let y = rFirst.pos; for (let r = g.r1; r < rFirst.i; r++) y -= geo.rowH(r);
  let x = cFirst.pos; for (let c = g.c1; c < cFirst.i; c++) x -= geo.colW(c);
  let h = 0; for (let r = g.r1; r <= g.r2; r++) h += geo.rowH(r);
  let w = 0; for (let c = g.c1; c <= g.c2; c++) w += geo.colW(c);
  return { x, y, w, h };
}

export interface ScrollThumbs { h: { at: number; frac: number }; v: { at: number; frac: number } }

export function Grid(p: GridProps) {
  const { wb, si, zoom } = p;
  const host = useRef<HTMLDivElement>(null);
  const [size, setSize] = useState({ w: 1200, h: 600 });
  useLayoutEffect(() => {
    const el = host.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setSize({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    setSize({ w: el.clientWidth, h: el.clientHeight });
    return () => ro.disconnect();
  }, []);
  // text measurements (digit tracking, overflow, ####) are redone once the cell fonts have loaded
  const [fontEpoch, setFontEpoch] = useState(0);
  useEffect(() => {
    const f = typeof document !== 'undefined' ? document.fonts : undefined;
    if (!f) return;
    let live = true;
    const on = () => { if (!live) return; dlsCache.clear(); setFontEpoch(e => e + 1); };
    f.addEventListener('loadingdone', on);
    void f.ready.then(on);
    return () => { live = false; f.removeEventListener('loadingdone', on); };
  }, []);
  const sh = wb.sheets[si];
  const geo = useMemo(() => new Geometry(wb, si, zoom), [wb, si, zoom, p.version]);
  const fz = sh.freeze ?? { rows: 0, cols: 0 };
  const headH = p.showHeadings ? geo.headH() : 0;
  const bottomRowGuess = p.top + Math.ceil(size.h / Math.max(8, geo.rowH(p.top) || 20)) + 2;
  const headW = p.showHeadings ? geo.rowHeadW(bottomRowGuess) : 0;
  const rows = useMemo(() => geo.rows(p.top, size.h - headH + 1, fz.rows), [geo, p.top, size.h, headH, fz.rows]);
  const cols = useMemo(() => geo.cols(p.left, size.w - headW + 1, fz.cols), [geo, p.left, size.w, headW, fz.cols]);
  const gridW = cols.length ? cols[cols.length - 1].pos + cols[cols.length - 1].size : 0;
  const gridH = rows.length ? rows[rows.length - 1].pos + rows[rows.length - 1].size : 0;
  useEffect(() => {
    // fully visible scrolled rows/cols, for PageDown and keeping the active cell in view
    const fr = rows.filter(b => b.i >= fz.rows && b.pos + b.size <= size.h - headH).length;
    const fc = cols.filter(b => b.i >= fz.cols && b.pos + b.size <= size.w - headW).length;
    // Scroll-bar thumbs in pixels, as Excel sizes them: the scrollable extent runs to the end of the column
    // (row) after the last used one, or to the last partly visible one if that is further; the thumb is the
    // visible share of that extent.
    const sum = (n: number, f: (i: number) => number) => { let s = 0; const lim = Math.min(n, 4000); for (let i = 0; i < lim; i++) s += f(i); if (n > lim) s += (n - lim) * (f(lim - 1) || 1); return s; };
    const lastC = cols.length ? cols[cols.length - 1].i + 1 : 1, lastR = rows.length ? rows[rows.length - 1].i + 1 : 1;
    const totW = Math.max(1, sum(Math.max(sh.maxCol + 1, lastC), c => geo.colW(c)));
    const totH = Math.max(1, sum(Math.max(sh.maxRow + 1, lastR), r => geo.rowH(r)));
    const offW = Math.max(0, sum(p.left, c => geo.colW(c)) - sum(fz.cols, c => geo.colW(c)));
    const offH = Math.max(0, sum(p.top, r => geo.rowH(r)) - sum(fz.rows, r => geo.rowH(r)));
    const hFrac = Math.min(1, (size.w - headW) / totW), vFrac = Math.min(1, (size.h - headH) / totH);
    p.onViewport(Math.max(1, fr), Math.max(1, fc), { h: { at: Math.min(1 - hFrac, offW / totW), frac: hFrac }, v: { at: Math.min(1 - vFrac, offH / totH), frac: vFrac } });
  }, [rows, cols, size.h, size.w]); // eslint-disable-line react-hooks/exhaustive-deps

  const rowBand = new Map(rows.map(b => [b.i, b]));
  const colBand = new Map(cols.map(b => [b.i, b]));
  const sel = p.sel;
  const g = sel.range;

  // ---------------------------------------------------------------- hit testing
  const hit = (clientX: number, clientY: number) => {
    const r0 = host.current!.getBoundingClientRect();
    const x = clientX - r0.left - headW, y = clientY - r0.top - headH;
    const find = (bands: Band[], v: number, max: number) => {
      for (const b of bands) if (b.size > 0 && v >= b.pos && v < b.pos + b.size) return b.i;
      if (v < 0) return bands.find(b => b.size > 0)?.i ?? 0;
      return Math.min(max, (bands[bands.length - 1]?.i ?? 0) + 1);
    };
    return { x, y, r: find(rows, y, MAX_ROW), c: find(cols, x, MAX_COL) };
  };

  const colEdge = (x: number) => { for (const b of cols) if (Math.abs(b.pos + b.size - x) <= 3) return b.i; return -1; };
  const rowEdge = (y: number) => { for (const b of rows) if (Math.abs(b.pos + b.size - y) <= 2) return b.i; return -1; };

  const [resizing, setResizing] = useState<'col' | 'row' | null>(null);
  const [hoverEdge, setHoverEdge] = useState<'col' | 'row' | null>(null);
  const [guide, setGuide] = useState<{ kind: 'col' | 'row'; pos: number; label: string; at: number } | null>(null);

  const extendSel = (a: Addr, b: Addr, kind: SelKind): Sel => {
    let range: Range = { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
    if (kind === 'cols') range = { ...range, r1: 0, r2: MAX_ROW };
    if (kind === 'rows') range = { ...range, c1: 0, c2: MAX_COL };
    if (kind === 'cells') {
      // a selection that touches a merged area grows to include all of it
      for (let pass = 0; pass < 3; pass++) for (const m of sh.merges) {
        if (m.r1 <= range.r2 && m.r2 >= range.r1 && m.c1 <= range.c2 && m.c2 >= range.c1) range = { r1: Math.min(range.r1, m.r1), c1: Math.min(range.c1, m.c1), r2: Math.max(range.r2, m.r2), c2: Math.max(range.c2, m.c2) };
      }
    }
    return { active: a, anchor: a, range, kind };
  };

  const onMouseDown = (e: React.MouseEvent) => {
    if (e.button === 1) return;
    const h = hit(e.clientX, e.clientY);
    const inColHead = p.showHeadings && h.y < 0 && h.x >= 0;
    const inRowHead = p.showHeadings && h.x < 0 && h.y >= 0;
    const inCorner = p.showHeadings && h.x < 0 && h.y < 0;
    if (e.button === 2) {
      // right-click keeps a selection that contains the point
      const inside = h.r >= g.r1 && h.r <= g.r2 && h.c >= g.c1 && h.c <= g.c2;
      if (inCorner) p.onSelect({ active: { r: 0, c: 0 }, anchor: { r: 0, c: 0 }, range: { r1: 0, c1: 0, r2: MAX_ROW, c2: MAX_COL }, kind: 'all' });
      else if (inColHead && !(inside && sel.kind === 'cols')) p.onSelect(extendSel({ r: p.top, c: h.c }, { r: p.top, c: h.c }, 'cols'));
      else if (inRowHead && !(inside && sel.kind === 'rows')) p.onSelect(extendSel({ r: h.r, c: p.left }, { r: h.r, c: p.left }, 'rows'));
      else if (!inColHead && !inRowHead && !inCorner && !inside) p.onSelect(extendSel({ r: h.r, c: h.c }, { r: h.r, c: h.c }, 'cells'));
      return;
    }
    e.preventDefault();
    host.current?.focus({ preventScroll: true });
    if (inCorner) { p.onSelect({ active: { r: Math.max(fz.rows, 0), c: 0 }, anchor: { r: 0, c: 0 }, range: { r1: 0, c1: 0, r2: MAX_ROW, c2: MAX_COL }, kind: 'all' }); return; }
    if (inColHead) {
      const edge = colEdge(h.x);
      if (edge >= 0) {
        if (e.detail >= 2) { p.onAutofitCol(edge); return; }
        const band = colBand.get(edge)!;
        setResizing('col');
        const startX = e.clientX;
        const move = (ev: MouseEvent) => { const w = Math.max(0, band.size + ev.clientX - startX); setGuide({ kind: 'col', pos: band.pos + w, label: `Width: ${(w / zoom / 7 - 5 / 7).toFixed(2)} (${Math.round(w / zoom)} pixels)`, at: band.pos }); };
        const up = (ev: MouseEvent) => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); setResizing(null); setGuide(null); p.onResizeCol(edge, Math.max(0, band.size + ev.clientX - startX) / zoom); };
        window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
        return;
      }
      const anchor = e.shiftKey && sel.kind === 'cols' ? sel.anchor : { r: Math.max(p.top, fz.rows), c: h.c };
      p.onSelect({ ...extendSel(anchor, { r: anchor.r, c: h.c }, 'cols'), active: { r: anchor.r, c: anchor.c } });
      const move = (ev: MouseEvent) => { const hh = hit(ev.clientX, ev.clientY); p.onSelect({ ...extendSel(anchor, { r: anchor.r, c: hh.c }, 'cols'), active: anchor }); };
      const up = () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
      window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
      return;
    }
    if (inRowHead) {
      const edge = rowEdge(h.y);
      if (edge >= 0) {
        const band = rowBand.get(edge)!;
        setResizing('row');
        const startY = e.clientY;
        const move = (ev: MouseEvent) => { const hh = Math.max(0, band.size + ev.clientY - startY); setGuide({ kind: 'row', pos: band.pos + hh, label: `Height: ${(hh / zoom * 0.75).toFixed(2)} (${Math.round(hh / zoom)} pixels)`, at: band.pos }); };
        const up = (ev: MouseEvent) => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); setResizing(null); setGuide(null); p.onResizeRow(edge, Math.max(0, band.size + ev.clientY - startY) / zoom); };
        window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
        return;
      }
      const anchor = e.shiftKey && sel.kind === 'rows' ? sel.anchor : { r: h.r, c: Math.max(p.left, fz.cols) };
      p.onSelect({ ...extendSel(anchor, { r: h.r, c: anchor.c }, 'rows'), active: anchor });
      const move = (ev: MouseEvent) => { const hh = hit(ev.clientX, ev.clientY); p.onSelect({ ...extendSel(anchor, { r: hh.r, c: anchor.c }, 'rows'), active: anchor }); };
      const up = () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
      window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
      return;
    }
    if (p.pointing) {
      // formula point mode: clicking a cell inserts its reference
      const a = { r: h.r, c: h.c };
      p.onPointRef(addrName(a.r, a.c), false);
      const move = (ev: MouseEvent) => { const hh = hit(ev.clientX, ev.clientY); const b = { r: hh.r, c: hh.c }; const ref = (a.r === b.r && a.c === b.c) ? addrName(a.r, a.c) : `${addrName(Math.min(a.r, b.r), Math.min(a.c, b.c))}:${addrName(Math.max(a.r, b.r), Math.max(a.c, b.c))}`; p.onPointRef(ref, true); };
      const up = () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
      window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
      return;
    }
    if (e.detail >= 2) { p.onDoubleClick(h.r, h.c); return; }
    const anchor = e.shiftKey ? sel.anchor : { r: h.r, c: h.c };
    const s0 = extendSel(anchor, { r: h.r, c: h.c }, 'cells');
    p.onSelect({ ...s0, active: e.shiftKey ? sel.active : anchor, anchor });
    const move = (ev: MouseEvent) => {
      const hh = hit(ev.clientX, ev.clientY);
      // dragging past the edge scrolls
      if (hh.y > size.h - headH - 4) p.setScroll(p.top + 1, p.left);
      const s = extendSel(anchor, { r: hh.r, c: hh.c }, 'cells');
      p.onSelect({ ...s, active: e.shiftKey ? sel.active : anchor, anchor });
    };
    const up = () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
    window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
  };

  const onMouseMove = (e: React.MouseEvent) => {
    if (resizing) return;
    const h = hit(e.clientX, e.clientY);
    if (p.showHeadings && h.y < 0 && h.x >= 0 && colEdge(h.x) >= 0) setHoverEdge('col');
    else if (p.showHeadings && h.x < 0 && h.y >= 0 && rowEdge(h.y) >= 0) setHoverEdge('row');
    else if (hoverEdge) setHoverEdge(null);
  };

  const onContextMenu = (e: React.MouseEvent) => {
    e.preventDefault();
    const h = hit(e.clientX, e.clientY);
    const kind = p.showHeadings && h.x < 0 && h.y < 0 ? 'all' : p.showHeadings && h.y < 0 ? 'col' : p.showHeadings && h.x < 0 ? 'row' : 'cell';
    p.onContext(kind, e.clientX, e.clientY);
  };

  const onWheel = (e: React.WheelEvent) => {
    const delta = e.deltaY || e.deltaX;
    if (!delta) return;
    if (e.ctrlKey) return;
    // three rows (one column) per notch; Chromium coalesces wheel events while the main thread is busy, so the
    // notch count comes from the delta (util/wheel.ts), at least one per event (touchpad-sized deltas)
    const n = wheelNotches(delta, e.deltaMode) || Math.sign(delta);
    if (e.shiftKey || Math.abs(e.deltaX) > Math.abs(e.deltaY)) p.setScroll(p.top, Math.max(fz.cols, p.left + n));
    else p.setScroll(Math.max(fz.rows, p.top + 3 * n), p.left);
  };

  // ---------------------------------------------------------------- cells
  const layers = useMemo(() => {
    const fills: ReactNode[] = [];
    const borders: ReactNode[] = [];
    const texts: ReactNode[] = [];
    const covered = new Set<string>(); // cells hidden under merges
    const mergesVisible = sh.merges.filter(m => rows.some(b => b.i >= m.r1 && b.i <= m.r2) && cols.some(b => b.i >= m.c1 && b.i <= m.c2));
    for (const m of mergesVisible) for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(`${r},${c}`);
    const lastCol = cols[cols.length - 1];
    for (const rb of rows) {
      if (rb.size === 0) continue;
      for (let ci = 0; ci < cols.length; ci++) {
        const cb = cols[ci];
        if (cb.size === 0) continue;
        const r = rb.i, c = cb.i;
        if (covered.has(`${r},${c}`)) continue;
        const merge = mergesVisible.find(m => m.r1 === r && m.c1 === c);
        const xf = wb.xfAt(si, r, c);
        const cell = wb.cell(si, r, c);
        if (!cell && xf === 0) continue;
        const st = wb.style(xf);
        let x = cb.pos, y = rb.pos, w = cb.size, h = rb.size;
        if (merge) { const mr = mergeRect(rows, cols, merge, geo); if (mr) ({ x, y, w, h } = mr); }
        const fill = st.fill.pattern !== 'none' && st.fill.pattern !== 'gray125' ? (st.fill.fg ?? st.fill.bg) : undefined;
        if (fill || merge) fills.push(<div key={`f${r},${c}`} className="xl-cellbg" style={{ left: x, top: y, width: w, height: h, background: fill ?? '#fff' }} />);
        const bd = st.border;
        const line = (side: 'top' | 'bottom' | 'left' | 'right') => {
          const b = bd[side];
          if (!b) return;
          const t = b.style === 'medium' || b.style === 'mediumDashed' ? 2 : b.style === 'thick' || b.style === 'double' ? 3 : 1;
          const color = b.color ?? '#000';
          const dashed = /dash|dot/i.test(b.style);
          const common = { position: 'absolute' as const, background: dashed ? undefined : color, borderTop: dashed && (side === 'top' || side === 'bottom') ? `${t}px dashed ${color}` : undefined, borderLeft: dashed && (side === 'left' || side === 'right') ? `${t}px dashed ${color}` : undefined };
          if (side === 'bottom') borders.push(<div key={`b${side}${r},${c}`} style={{ ...common, left: x - 1, top: y + h - Math.ceil(t / 2) - (t === 1 ? 0 : 0) - 1 + (t > 1 ? 1 : 0), width: w + 1, height: dashed ? 0 : t }} />);
          if (side === 'top') borders.push(<div key={`b${side}${r},${c}`} style={{ ...common, left: x - 1, top: y - Math.ceil(t / 2), width: w + 1, height: dashed ? 0 : t }} />);
          if (side === 'left') borders.push(<div key={`b${side}${r},${c}`} style={{ ...common, left: x - Math.ceil(t / 2), top: y - 1, width: dashed ? 0 : t, height: h + 1 }} />);
          if (side === 'right') borders.push(<div key={`b${side}${r},${c}`} style={{ ...common, left: x + w - 1 - Math.floor((t - 1) / 2), top: y - 1, width: dashed ? 0 : t, height: h + 1 }} />);
        };
        line('top'); line('bottom'); line('left'); line('right');
        if (!cell) continue;
        let text: string;
        let color: string | undefined;
        let numeric = false;
        let fillSplit: { left: string; right: string } | undefined;
        if (p.showFormulas && cell.f) text = '=' + cell.f;
        else {
          const d = wb.display(si, r, c);
          text = d.text; color = d.color; fillSplit = d.fill;
          numeric = typeof d.value === 'number';
          if (isErr(d.value) || typeof d.value === 'boolean') numeric = false;
        }
        if (!text) continue;
        const font = st.font;
        const px = cellFontPx(font.sz || 11, zoom);
        const css = `${font.i ? 'italic ' : ''}${font.b ? '700 ' : ''}${px}px ${cssFamily(font.name)}`;
        const val = wb.value(si, r, c);
        const halign = st.align.h && st.align.h !== 'general' ? st.align.h : (p.showFormulas && cell.f ? 'left' : numeric ? 'right' : (typeof val === 'boolean' || isErr(val)) ? 'center' : 'left');
        const wrap = !!st.align.wrap;
        let bx = x, bw = w;
        const dls = digitSpacing(css, (font.sz || 11) * 96 / 72 * zoom, cssFamily(font.name), font);
        const tw = cellTextWidth(text, css, dls) + 5 * zoom;
        if (numeric && !fillSplit && tw > w && !wrap) {
          text = '#'.repeat(Math.max(1, Math.floor((w - 4) / Math.max(1, textWidth('#', css)))));
        } else if (!numeric && !wrap && !merge && tw > w) {
          // overflow into empty neighbours
          const empty = (cc: number) => { const o = wb.cell(si, r, cc); return !o || (o.v === null && !o.f) ; };
          if (halign === 'left' || halign === 'center' || halign === 'fill') {
            let k = ci + 1, extra = 0;
            const need = halign === 'center' ? (tw - w) / 2 : tw - w;
            while (k < cols.length && extra < need && empty(cols[k].i) && !covered.has(`${r},${cols[k].i}`)) { extra += cols[k].size; k++; }
            if (k >= cols.length && extra < need && lastCol) extra = need;
            bw += extra;
            if (halign === 'center') { let k2 = ci - 1, extraL = 0; while (k2 >= 0 && extraL < need && empty(cols[k2].i)) { extraL += cols[k2].size; k2--; } bx -= extraL; bw += extraL; }
          } else if (halign === 'right') {
            let k = ci - 1, extra = 0;
            while (k >= 0 && extra < tw - w && empty(cols[k].i)) { extra += cols[k].size; k--; }
            bx -= extra; bw += extra;
          }
        }
        const va = st.align.v ?? 'bottom';
        const deco = [font.u ? 'underline' : '', font.strike ? 'line-through' : ''].filter(Boolean).join(' ') || undefined;
        const indent = (st.align.indent ?? 0) * 9 * zoom;
        texts.push(
          <div key={`t${r},${c}`} className={`xl-cell${wrap ? ' wrap' : ''}`} data-control-item={addrName(r, c)}
            style={{
              left: bx, top: y, width: bw, height: h - 1, font: css, color: color ?? font.color ?? '#000', textDecoration: deco,
              justifyContent: fillSplit ? 'space-between' : halign === 'right' ? 'flex-end' : halign === 'center' || halign === 'centerContinuous' ? 'center' : 'flex-start',
              alignItems: va === 'top' ? 'flex-start' : va === 'center' ? 'center' : 'flex-end',
              textAlign: halign === 'right' ? 'right' : halign === 'center' ? 'center' : 'left',
              paddingLeft: halign === 'left' ? 2 + indent : undefined, paddingRight: halign === 'right' ? 3 + indent : undefined,
              background: 'transparent',
            }}>
            {fillSplit ? <><span>{fillSplit.left}</span><span>{fillSplit.right}</span></> : digitRuns(text, dls)}
          </div>,
        );
      }
    }
    return { fills, borders, texts };
  }, [rows, cols, geo, p.version, p.showFormulas, si, wb, zoom, sh.merges, fontEpoch]); // eslint-disable-line react-hooks/exhaustive-deps

  // ---------------------------------------------------------------- overlays
  const rectOf = (rg: Range) => {
    const rVis = rows.filter(b => b.i >= rg.r1 && b.i <= rg.r2);
    const cVis = cols.filter(b => b.i >= rg.c1 && b.i <= rg.c2);
    if (!rVis.length || !cVis.length) {
      // partly off screen: clamp to the edges so the border still shows where it continues
      return null;
    }
    const top = rVis[0], bot = rVis[rVis.length - 1], lft = cVis[0], rgt = cVis[cVis.length - 1];
    return { x: lft.pos, y: top.pos, w: rgt.pos + rgt.size - lft.pos, h: bot.pos + bot.size - top.pos, clippedRight: rg.c2 > rgt.i, clippedBottom: rg.r2 > bot.i };
  };
  const selRect = rectOf(g);
  const activeMerge = sh.merges.find(m => m.r1 <= sel.active.r && m.r2 >= sel.active.r && m.c1 <= sel.active.c && m.c2 >= sel.active.c);
  const activeRect = rectOf(activeMerge ?? { r1: sel.active.r, c1: sel.active.c, r2: sel.active.r, c2: sel.active.c });
  const antsRect = p.ants ? rectOf(p.ants) : null;
  const single = g.r1 === g.r2 && g.c1 === g.c2 || (activeMerge && g.r1 === activeMerge.r1 && g.r2 === activeMerge.r2 && g.c1 === activeMerge.c1 && g.c2 === activeMerge.c2);

  // editor
  const ed = p.editor;
  let editorBox: { x: number; y: number; w: number; h: number; font: string; align: string } | null = null;
  if (ed) {
    const m = sh.merges.find(mm => mm.r1 === ed.r && mm.c1 === ed.c);
    const r = m ? mergeRect(rows, cols, m, geo) : (rowBand.get(ed.r) && colBand.get(ed.c) ? { x: colBand.get(ed.c)!.pos, y: rowBand.get(ed.r)!.pos, w: colBand.get(ed.c)!.size, h: rowBand.get(ed.r)!.size } : null);
    if (r) {
      const st = wb.style(wb.xfAt(si, ed.r, ed.c));
      const px = cellFontPx(st.font.sz || 11, zoom);
      const font = `${st.font.i ? 'italic ' : ''}${st.font.b ? '700 ' : ''}${px}px ${cssFamily(st.font.name)}`;
      const lines = ed.text.split('\n');
      const tw = Math.max(...lines.map(l => textWidth(l, font))) + 12;
      editorBox = { x: r.x, y: r.y, w: Math.min(Math.max(r.w + 1, tw), size.w - headW - r.x), h: Math.max(r.h + 1, lines.length * px * 1.2 + 4), font, align: 'left' };
    }
  }
  useEffect(() => {
    const t = p.editorRef.current;
    if (t && ed && document.activeElement !== t && !(document.activeElement as HTMLElement | null)?.classList.contains('xl-formula')) {
      t.focus({ preventScroll: true });
      const at = ed.caret ?? ed.text.length;
      t.setSelectionRange(at, at);
    }
    // also when the editor first becomes visible: an edit started on a cell that a scroll is still bringing into
    // view has no textarea on its first render, and keys typed meanwhile must still reach it
  }, [ed?.r, ed?.c, !!ed, !!editorBox]); // eslint-disable-line react-hooks/exhaustive-deps

  const selColsFull = sel.kind === 'cols' || sel.kind === 'all';
  const selRowsFull = sel.kind === 'rows' || sel.kind === 'all';

  return (
    <div ref={host} className={`xl-grid${resizing === 'col' || hoverEdge === 'col' ? ' resizing-col' : ''}${resizing === 'row' || hoverEdge === 'row' ? ' resizing-row' : ''}`}
      data-control-id="excel.grid_cells" tabIndex={-1}
      onMouseDown={onMouseDown} onMouseMove={onMouseMove} onContextMenu={onContextMenu} onWheel={onWheel}>
      {/* cells */}
      <div className="xl-layer" style={{ left: headW, top: headH, width: size.w - headW, height: size.h - headH, overflow: 'hidden' }}>
        {p.showGrid && rows.map(b => b.size > 0 && <div key={`h${b.i}`} className="xl-hline" style={{ left: 0, top: b.pos + b.size - 1, width: gridW }} />)}
        {p.showGrid && cols.map(b => b.size > 0 && <div key={`v${b.i}`} className="xl-vline" style={{ top: 0, left: b.pos + b.size - 1, height: gridH }} />)}
        {layers.fills}
        {layers.borders}
        <div data-control-id="excel.freeze_panes_merged_cells_rendering" style={{ position: 'absolute', inset: 0 }}>
          {layers.texts}
        </div>
        {p.filter && rows.some(b => b.i === p.filter!.r1) && cols.filter(b => b.i >= p.filter!.c1 && b.i <= p.filter!.c2).map(b => {
          const rb = rowBand.get(p.filter!.r1)!;
          return <button key={`fb${b.i}`} className="xl-filter-btn" data-control-id="excel.data_sort_and_filter" data-control-item={`Filter ${colName(b.i)}`} aria-label={`${colName(b.i)} filter`}
            style={{ left: b.pos + b.size - 18, top: rb.pos + rb.size - 18 }}
            onMouseDown={e => { e.stopPropagation(); e.preventDefault(); p.onFilterMenu(b.i, e.currentTarget); }}><Icon name="dropdown" size={12} /></button>;
        })}
        {selRect && !single && (
          <>
            <div className="xl-sel-fill" style={{
              left: selRect.x, top: selRect.y, width: selRect.w, height: selRect.h,
              // the active cell stays white (a hole in the tint)
              clipPath: activeRect ? `path(evenodd, 'M0 0H${selRect.w}V${selRect.h}H0Z M${activeRect.x - selRect.x} ${activeRect.y - selRect.y}h${activeRect.w - 1}v${activeRect.h - 1}h${-(activeRect.w - 1)}Z')` : undefined,
            }} />
          </>
        )}
        {selRect && <div className="xl-sel" data-control-id="excel.grid_cells" data-control-item="selection" style={{ left: selRect.x - 2, top: selRect.y - 2, width: selRect.w + 3, height: selRect.h + 3 }} />}
        {selRect && !selRect.clippedRight && !selRect.clippedBottom && <div className="xl-handle" style={{ left: selRect.x + selRect.w - 4, top: selRect.y + selRect.h - 4 }} />}
        {selRect && !single && p.quickAnalysis && !selRect.clippedRight && !selRect.clippedBottom && <div className="xl-quick" aria-label="Quick Analysis" title="Quick Analysis (Ctrl+Q)" style={{ left: selRect.x + selRect.w + 3, top: selRect.y + selRect.h + 3 }}><Icon name="quickAnalysis" size={16} /></div>}
        {antsRect && <div className="xl-ants" data-control-id="excel.copy_ctrl_c_marching_ants" style={{ left: antsRect.x - 1, top: antsRect.y - 1, width: antsRect.w + 2, height: antsRect.h + 2 }} />}
        {fz.rows > 0 && <div className="xl-freeze-h" style={{ left: 0, top: (rowBand.get(fz.rows - 1)?.pos ?? 0) + (rowBand.get(fz.rows - 1)?.size ?? 0) - 1, width: gridW }} />}
        {fz.cols > 0 && <div className="xl-freeze-v" style={{ top: 0, left: (colBand.get(fz.cols - 1)?.pos ?? 0) + (colBand.get(fz.cols - 1)?.size ?? 0) - 1, height: gridH }} />}
        {guide && guide.kind === 'col' && <div style={{ position: 'absolute', left: guide.pos - 1, top: 0, width: 1, height: gridH, background: '#444' }} />}
        {guide && guide.kind === 'row' && <div style={{ position: 'absolute', top: guide.pos - 1, left: 0, height: 1, width: gridW, background: '#444' }} />}
      </div>
      {editorBox && ed && (
        <textarea ref={p.editorRef} className="xl-editor" spellCheck={false} value={ed.text} data-control-id="excel.grid_cells" data-control-item="cell editor"
          style={{ left: headW + editorBox.x, top: headH + editorBox.y, width: editorBox.w, height: editorBox.h, font: editorBox.font }}
          onChange={e => p.onEditorText(e.target.value, e.target.selectionStart)}
          onKeyDown={p.onEditorKey}
          onSelect={e => p.onEditorText((e.target as HTMLTextAreaElement).value, (e.target as HTMLTextAreaElement).selectionStart)}
          onMouseDown={e => e.stopPropagation()} />
      )}
      {/* headers */}
      {p.showHeadings && (
        <div data-control-id="excel.column_row_headers">
          <div style={{ position: 'absolute', left: headW, top: 0, width: size.w - headW, height: headH, overflow: 'hidden' }}>
            {cols.map(b => {
              if (b.size === 0) return null;
              const inSel = b.i >= g.c1 && b.i <= g.c2;
              return <div key={b.i} className={`xl-colhead${inSel ? ' sel' : ''}${inSel && selColsFull ? ' full' : ''}${b.i === g.c1 - 1 ? ' sel-before' : ''}`} data-control-item={colName(b.i)}
                style={{ left: b.pos - 1, width: b.size + 1, height: headH, fontSize: 13 * Math.min(1, zoom) }}>{colName(b.i)}</div>;
            })}
          </div>
          <div style={{ position: 'absolute', left: 0, top: headH, width: headW, height: size.h - headH, overflow: 'hidden' }}>
            {rows.map(b => {
              if (b.size === 0) return null;
              const inSel = b.i >= g.r1 && b.i <= g.r2;
              return <div key={b.i} className={`xl-rowhead${inSel ? ' sel' : ''}${inSel && selRowsFull ? ' full' : ''}${b.i === g.r1 - 1 ? ' sel-before' : ''}`} data-control-item={String(b.i + 1)}
                style={{ top: b.pos - 1, height: b.size + 1, width: headW, fontSize: 13 * Math.min(1, zoom) }}>{b.i + 1}</div>;
            })}
          </div>
          <div className="xl-corner" style={{ width: headW, height: headH }} data-control-item="Select All" title="Select All" />
        </div>
      )}
      {p.fab && <div className="xl-copilot-fab" aria-label={BRAND.copilot} title={BRAND.copilot} style={{ right: 25, bottom: 25 }} onMouseDown={e => e.stopPropagation()}><Icon name="copilotFab" size={26} /></div>}
      {guide && <div style={{ position: 'absolute', left: headW + guide.at + 4, top: guide.kind === 'col' ? 2 : headH + guide.at - 22, background: '#fff', border: '1px solid #888', padding: '1px 5px', fontSize: 11, zIndex: 9, whiteSpace: 'nowrap' }}>{guide.label}</div>}
    </div>
  );
}
