// Excel's own dialogs (drawn in the window, modeless like Excel's Find and Replace): Find and
// Replace with Options, Go To, Column Width / Row Height, Zoom, Format Cells, Insert Function.
import { useEffect, useRef, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { Icon } from './icons';

export function Dialog({ title, width, children, onClose, id, onKeyDown, initial }: { title: string; width: number; children: ReactNode; onClose: () => void; id?: string; onKeyDown?: (e: React.KeyboardEvent) => void; initial?: { x: number; y: number } }) {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState<{ x: number; y: number } | null>(initial ?? null);
  useEffect(() => {
    if (pos) return;
    const el = ref.current, parent = el?.parentElement;
    if (!el || !parent) return;
    setPos({ x: Math.max(0, (parent.clientWidth - el.offsetWidth) / 2), y: Math.max(40, (parent.clientHeight - el.offsetHeight) / 2 - 40) });
  }, [pos]);
  const drag = (e: React.MouseEvent) => {
    if ((e.target as HTMLElement).closest('button')) return;
    const sx = e.clientX, sy = e.clientY, p0 = pos ?? { x: 0, y: 0 };
    const move = (ev: MouseEvent) => setPos({ x: p0.x + ev.clientX - sx, y: Math.max(0, p0.y + ev.clientY - sy) });
    const up = () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
    window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
  };
  return (
    <div ref={ref} className="xl-dlg" role="dialog" aria-label={title} data-control-id={id} style={{ width, left: pos?.x ?? -9999, top: pos?.y ?? 0 }}
      onKeyDown={e => { e.stopPropagation(); if (e.key === 'Escape') { e.preventDefault(); onClose(); return; } onKeyDown?.(e); }}
      onMouseDown={e => e.stopPropagation()}>
      <div className="xl-dlg-title" onMouseDown={drag}>
        <span>{title}</span>
        <button aria-label="Close" data-control-id={id} data-control-item="Close dialog" onClick={onClose}><Icon name="close" size={10} /></button>
      </div>
      <div className="xl-dlg-body">{children}</div>
    </div>
  );
}

export interface FindOptions { what: string; repl: string; within: 'sheet' | 'workbook'; by: 'rows' | 'cols'; lookIn: 'formulas' | 'values' | 'notes'; matchCase: boolean; entire: boolean }
export interface FindHit { si: number; sheet: string; r: number; c: number; addr: string; value: string; formula: string }

/** Find and Replace: a separate top-level window in Excel, so it is drawn on the desktop layer (portal) and may
 *  overlap the taskbar like the real one. It opens at the bottom right of the screen . */
export function FindReplace({ mode, opts, setOpts, onFindNext, onFindAll, onReplace, onReplaceAll, onClose, onPick, results, status, setMode, canReplace, onFocusChange, hidden }: {
  mode: 'find' | 'replace'; setMode: (m: 'find' | 'replace') => void;
  opts: FindOptions & { expanded: boolean }; setOpts: (o: Partial<FindOptions & { expanded: boolean }>) => void;
  onFindNext: () => void; onFindAll: () => void; onReplace: () => void; onReplaceAll: () => void; onClose: () => void;
  onPick: (h: FindHit) => void; results: FindHit[] | null; status: string; canReplace: boolean;
  onFocusChange?: (focused: boolean) => void; hidden?: boolean;
}) {
  const ID = 'excel.find_and_replace_ctrl_f_h';
  const whatRef = useRef<HTMLInputElement>(null);
  const replRef = useRef<HTMLInputElement>(null);
  const rootRef = useRef<HTMLDivElement>(null);
  const [selRow, setSelRow] = useState(-1);
  const [pos, setPos] = useState<{ x: number; y: number }>(() => ({ x: Math.max(0, window.innerWidth - 450), y: Math.max(0, window.innerHeight - 249) }));
  useEffect(() => { whatRef.current?.focus(); whatRef.current?.select(); }, [mode]);
  const W = 450;
  const optH = opts.expanded ? 75 : 0;
  const replH = mode === 'replace' ? 29 : 0;
  const resH = results ? 150 : 0;
  const H = 168 + optH + replH + resH;
  const drag = (e: React.MouseEvent) => {
    if ((e.target as HTMLElement).closest('button')) return;
    e.preventDefault();
    const sx = e.clientX, sy = e.clientY, p0 = pos;
    const move = (ev: MouseEvent) => setPos({ x: p0.x + ev.clientX - sx, y: Math.max(0, p0.y + ev.clientY - sy) });
    const up = () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
    window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
  };
  const btn = (label: ReactNode, item: string, onClick: () => void, cls: string, opt: { def?: boolean; disabled?: boolean } = {}) => (
    <button className={`xl-fnd-btn ${cls}${opt.def ? ' default' : ''}`} data-control-id={ID} data-control-item={item} disabled={opt.disabled} onClick={onClick}>{label}</button>
  );
  // accelerators (Alt+letter, as the underlined letters show)
  const onKey = (e: React.KeyboardEvent) => {
    e.stopPropagation();
    if (e.key === 'Escape') { e.preventDefault(); onClose(); return; }
    if (e.altKey && !e.ctrlKey && e.key.length === 1) {
      const k = e.key.toLowerCase();
      const focusSel = (n: string) => (rootRef.current?.querySelector(`[data-control-item="${n}"]`) as HTMLElement | null)?.focus();
      const act: Record<string, () => void> = {
        t: () => setOpts({ expanded: !opts.expanded }),
        n: () => { whatRef.current?.focus(); whatRef.current?.select(); },
        d: () => setMode('find'),
        p: () => { if (canReplace) setMode('replace'); },
        i: () => { if (opts.what) onFindAll(); },
        f: () => { if (opts.what) onFindNext(); },
        ...(mode === 'replace' ? { e: () => replRef.current?.focus(), a: () => { if (opts.what) onReplaceAll(); }, r: () => { if (opts.what) onReplace(); } } : {}),
        ...(opts.expanded ? { h: () => focusSel('Within'), s: () => focusSel('Search'), l: () => focusSel('Look in'), c: () => setOpts({ matchCase: !opts.matchCase }), o: () => setOpts({ entire: !opts.entire }) } : {}),
      };
      if (act[k]) { e.preventDefault(); act[k](); }
      return;
    }
    if (e.key === 'Enter' && !(e.target instanceof HTMLButtonElement)) { e.preventDefault(); onFindNext(); }
  };
  const node = (
    <div ref={rootRef} className="xl-fnd" role="dialog" aria-label="Find and Replace" data-control-id={ID}
      style={{ left: pos.x, top: pos.y, width: W, height: H, display: hidden ? 'none' : undefined }}
      onKeyDown={onKey} onMouseDown={e => e.stopPropagation()}
      onFocus={() => onFocusChange?.(true)} onBlur={e => { if (!e.currentTarget.contains(e.relatedTarget as Node | null)) onFocusChange?.(false); }}>
      <div className="xl-fnd-title" onMouseDown={drag}>
        <span>Find and Replace</span>
        <button className="xl-fnd-cap" disabled aria-label="Minimize"><svg width="10" height="10" viewBox="0 0 10 10"><path d="M0 5.5h10" stroke="currentColor" /></svg></button>
        <button className="xl-fnd-cap" aria-label="Maximize" tabIndex={-1}><svg width="10" height="10" viewBox="0 0 10 10"><rect x=".5" y=".5" width="9" height="9" fill="none" stroke="currentColor" /></svg></button>
        <button className="xl-fnd-cap close" aria-label="Close" data-control-id={ID} data-control-item="Close dialog" onClick={onClose}><svg width="10" height="10" viewBox="0 0 10 10"><path d="M.5 .5l9 9M9.5 .5l-9 9" stroke="currentColor" /></svg></button>
      </div>
      <div className="xl-fnd-body">
        <button className={`xl-fnd-tab${mode === 'find' ? ' active' : ''}`} style={{ left: 6 }} data-control-id={ID} data-control-item="Find" onClick={() => setMode('find')}>Fin<u>d</u></button>
        <button className={`xl-fnd-tab${mode === 'replace' ? ' active' : ''}`} style={{ left: 61 }} data-control-id={ID} data-control-item="Replace" onClick={() => setMode('replace')} disabled={!canReplace}>Re<u>p</u>lace</button>
        <div className="xl-fnd-page" style={{ top: 28, height: 69 + optH + replH }}>
          <label className="xl-fnd-label" style={{ top: 9 }}>Fi<u>n</u>d what:</label>
          <div className="xl-fnd-combo" style={{ top: 9, left: 86 }}>
            <input ref={whatRef} type="text" value={opts.what} onChange={e => setOpts({ what: e.target.value })} data-control-id={ID} data-control-item="Find what" spellCheck={false} />
            <span className="xl-fnd-dd" aria-hidden><DdChevron /></span>
          </div>
          {opts.expanded && <><div className="xl-fnd-noformat" style={{ top: 9 }}>No Format Set</div><button className="xl-fnd-btn xl-fnd-format" style={{ top: 10 }} disabled>For<u>m</u>at...<span className="xl-fnd-split" aria-hidden>▾</span></button></>}
          {mode === 'replace' && (
            <>
              <label className="xl-fnd-label" style={{ top: 38 }}>R<u>e</u>place with:</label>
              <div className="xl-fnd-combo" style={{ top: 38, left: 86 }}>
                <input ref={replRef} type="text" value={opts.repl} onChange={e => setOpts({ repl: e.target.value })} data-control-id={ID} data-control-item="Replace with" spellCheck={false} />
                <span className="xl-fnd-dd" aria-hidden><DdChevron /></span>
              </div>
              {opts.expanded && <><div className="xl-fnd-noformat" style={{ top: 38 }}>No Format Set</div><button className="xl-fnd-btn xl-fnd-format" style={{ top: 39 }} disabled>For<u>m</u>at...<span className="xl-fnd-split" aria-hidden>▾</span></button></>}
            </>
          )}
          {opts.expanded && (
            <>
              <label className="xl-fnd-label" style={{ top: 62 + replH }}>Wit<u>h</u>in:</label>
              <label className="xl-fnd-label" style={{ top: 87 + replH }}><u>S</u>earch:</label>
              <label className="xl-fnd-label" style={{ top: 112 + replH }}><u>L</u>ook in:</label>
              <FndSelect top={62 + replH} value={opts.within} item="Within" id={ID} options={[['sheet', 'Sheet'], ['workbook', 'Workbook']]} onChange={v => setOpts({ within: v as 'sheet' })} />
              <FndSelect top={87 + replH} value={opts.by} item="Search" id={ID} options={[['rows', 'By Rows'], ['cols', 'By Columns']]} onChange={v => setOpts({ by: v as 'rows' })} />
              <FndSelect top={112 + replH} value={mode === 'replace' ? 'formulas' : opts.lookIn} item="Look in" id={ID} disabled={mode === 'replace'} options={[['formulas', 'Formulas'], ['values', 'Values'], ['notes', 'Notes'], ['comments', 'Comments']]} onChange={v => setOpts({ lookIn: v as 'values' })} />
              <label className="xl-fnd-check" style={{ top: 63 + replH }}><input type="checkbox" checked={opts.matchCase} onChange={e => setOpts({ matchCase: e.target.checked })} data-control-id={ID} data-control-item="Match case" /><span className="xl-fnd-box" /><span>Match <u>c</u>ase</span></label>
              <label className="xl-fnd-check" style={{ top: 82 + replH }}><input type="checkbox" checked={opts.entire} onChange={e => setOpts({ entire: e.target.checked })} data-control-id={ID} data-control-item="Match entire cell contents" /><span className="xl-fnd-box" /><span>Match entire cell c<u>o</u>ntents</span></label>
            </>
          )}
          {btn(opts.expanded ? <>Op<u>t</u>ions &lt;&lt;</> : <>Op<u>t</u>ions &gt;&gt;</>, opts.expanded ? 'Options <<' : 'Options >>', () => setOpts({ expanded: !opts.expanded }), 'xl-fnd-options')}
        </div>
        <div className="xl-fnd-btns" style={{ top: 106 + optH + replH }}>
          {mode === 'replace' && btn(<>Replace <u>A</u>ll</>, 'Replace All', onReplaceAll, '', { disabled: !opts.what })}
          {mode === 'replace' && btn(<><u>R</u>eplace</>, 'Replace', onReplace, '', { disabled: !opts.what })}
          {btn(<>F<u>i</u>nd All</>, 'Find All', onFindAll, '', { disabled: !opts.what })}
          {btn(<><u>F</u>ind Next</>, 'Find Next', onFindNext, '', { def: true, disabled: !opts.what })}
          {btn('Close', 'Close', onClose, 'xl-fnd-close')}
        </div>
        {results && (
          <div className="xl-dlg-results xl-fnd-results" style={{ top: 136 + optH + replH }}>
            <table>
              <thead><tr><th>Book</th><th>Sheet</th><th>Name</th><th>Cell</th><th>Value</th><th>Formula</th></tr></thead>
              <tbody>
                {results.map((h, i) => (
                  <tr key={i} className={i === selRow ? 'sel' : ''} data-control-id={ID} data-control-item={`${h.sheet}!${h.addr}`} onMouseDown={() => { setSelRow(i); onPick(h); }}>
                    <td>{(document.querySelector('.xl-root')?.getAttribute('data-book') ?? '')}</td><td>{h.sheet}</td><td></td><td>${h.addr.replace(/(\d)/, '$$$1')}</td><td>{h.value}</td><td>{h.formula}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div className="xl-dlg-status">{status}</div>
          </div>
        )}
      </div>
    </div>
  );
  return createPortal(node, document.body);
}

/** the classic combo-box drop arrow: an 8 x 4 px chevron on whole pixels */
function DdChevron() {
  return (
    <svg className="xl-fnd-chev" width="8" height="4" viewBox="0 0 8 4" shapeRendering="crispEdges" aria-hidden>
      {[0, 1, 2, 3].map(i => <g key={i}><rect x={i} y={i} width="1" height="1" /><rect x={7 - i} y={i} width="1" height="1" /></g>)}
    </svg>
  );
}

function FndSelect({ top, value, options, onChange, item, id, disabled }: { top: number; value: string; options: [string, string][]; onChange: (v: string) => void; item: string; id: string; disabled?: boolean }) {
  return (
    <div className="xl-fnd-select" style={{ top }}>
      <select value={value} disabled={disabled} onChange={e => onChange(e.target.value)} data-control-id={id} data-control-item={item}>
        {options.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
      </select>
      <span className="xl-fnd-dd" aria-hidden><DdChevron /></span>
    </div>
  );
}

export function GoTo({ names, onGo, onClose }: { names: string[]; onGo: (ref: string) => void; onClose: () => void }) {
  const ID = 'excel.go_to_ctrl_g_f5';
  const [ref, setRef] = useState('');
  const [sel, setSel] = useState(-1);
  const input = useRef<HTMLInputElement>(null);
  useEffect(() => { input.current?.focus(); }, []);
  return (
    <Dialog title="Go To" width={300} onClose={onClose} id={ID} onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); if (ref.trim()) onGo(ref.trim()); } }}>
      <div>Go to:</div>
      <div className="xl-goto-list" data-control-id={ID} data-control-item="Go to list">
        {names.map((n, i) => <div key={n} className={i === sel ? 'sel' : ''} onMouseDown={() => { setSel(i); setRef(n); }} onDoubleClick={() => onGo(n)}>{n}</div>)}
      </div>
      <div className="xl-dlg-field"><label style={{ width: 64 }}><u>R</u>eference:</label><input ref={input} type="text" value={ref} onChange={e => setRef(e.target.value)} data-control-id={ID} data-control-item="Reference" /></div>
      <div className="xl-dlg-btns" style={{ justifyContent: 'space-between' }}>
        <button className="xl-dlg-btn" disabled>Special...</button>
        <span style={{ display: 'flex', gap: 6 }}>
          <button className="xl-dlg-btn default" data-control-id={ID} data-control-item="OK" disabled={!ref.trim()} onClick={() => onGo(ref.trim())}>OK</button>
          <button className="xl-dlg-btn" data-control-id={ID} data-control-item="Cancel" onClick={onClose}>Cancel</button>
        </span>
      </div>
    </Dialog>
  );
}

export function NumberPrompt({ title, label, value, onOk, onClose }: { title: string; label: string; value: string; onOk: (v: number) => void; onClose: () => void }) {
  const [v, setV] = useState(value);
  const input = useRef<HTMLInputElement>(null);
  useEffect(() => { input.current?.focus(); input.current?.select(); }, []);
  const ok = () => { const n = Number(v); if (Number.isFinite(n) && n >= 0) onOk(n); };
  return (
    <Dialog title={title} width={240} onClose={onClose} onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); ok(); } }}>
      <div className="xl-dlg-field"><label style={{ width: 90 }}>{label}</label><input ref={input} type="text" value={v} onChange={e => setV(e.target.value)} /></div>
      <div className="xl-dlg-btns"><button className="xl-dlg-btn default" onClick={ok}>OK</button><button className="xl-dlg-btn" onClick={onClose}>Cancel</button></div>
    </Dialog>
  );
}

export function ZoomDialog({ zoom, onOk, onClose }: { zoom: number; onOk: (z: number) => void; onClose: () => void }) {
  const opts = [200, 100, 75, 50, 25];
  const [v, setV] = useState(opts.includes(Math.round(zoom * 100)) ? String(Math.round(zoom * 100)) : 'custom');
  const [custom, setCustom] = useState(String(Math.round(zoom * 100)));
  const ok = () => { const n = v === 'custom' ? Number(custom) : Number(v); if (n >= 10 && n <= 400) onOk(n / 100); };
  return (
    <Dialog title="Zoom" width={200} onClose={onClose} onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); ok(); } }}>
      <fieldset style={{ border: '1px solid #d0d0d0', padding: '6px 10px', display: 'flex', flexDirection: 'column', gap: 4 }}>
        <legend>Magnification</legend>
        {opts.map(o => <label key={o} className="xl-check"><input type="radio" checked={v === String(o)} onChange={() => setV(String(o))} />{o}%</label>)}
        <label className="xl-check"><input type="radio" checked={v === 'custom'} onChange={() => setV('custom')} />Custom: <input type="text" value={custom} onChange={e => { setCustom(e.target.value); setV('custom'); }} style={{ width: 40 }} />%</label>
      </fieldset>
      <div className="xl-dlg-btns"><button className="xl-dlg-btn default" onClick={ok}>OK</button><button className="xl-dlg-btn" onClick={onClose}>Cancel</button></div>
    </Dialog>
  );
}

export const NUMBER_CATEGORIES: { label: string; code: (d: number) => string }[] = [
  { label: 'General', code: () => 'General' },
  { label: 'Number', code: d => d ? `0.${'0'.repeat(d)}` : '0' },
  { label: 'Currency', code: d => `"$"#,##0${d ? '.' + '0'.repeat(d) : ''}` },
  { label: 'Accounting', code: d => `_("$"* #,##0${d ? '.' + '0'.repeat(d) : ''}_);_("$"* \\(#,##0${d ? '.' + '0'.repeat(d) : ''}\\);_("$"* "-"??_);_(@_)` },
  { label: 'Date', code: () => 'm/d/yyyy' },
  { label: 'Time', code: () => 'h:mm:ss AM/PM' },
  { label: 'Percentage', code: d => `0${d ? '.' + '0'.repeat(d) : ''}%` },
  { label: 'Scientific', code: d => `0.${'0'.repeat(Math.max(1, d))}E+00` },
  { label: 'Text', code: () => '@' },
];

export function FormatCells({ tab: tab0, sample, onApply, onClose }: { tab: 'number' | 'alignment' | 'font'; sample: string; onApply: (p: { numFmt?: string; h?: string; v?: string; wrap?: boolean; bold?: boolean; italic?: boolean; size?: number }) => void; onClose: () => void }) {
  const [tab, setTab] = useState(tab0);
  const [cat, setCat] = useState(0);
  const [dec, setDec] = useState(2);
  const [h, setH] = useState('general');
  const [v, setV] = useState('bottom');
  const [wrap, setWrap] = useState(false);
  const [bold, setBold] = useState(false);
  const [italic, setItalic] = useState(false);
  const [size, setSize] = useState(11);
  const [touched, setTouched] = useState<Set<string>>(new Set());
  const t = (k: string) => setTouched(s => new Set(s).add(k));
  const ok = () => {
    const p: Parameters<typeof onApply>[0] = {};
    if (touched.has('num')) p.numFmt = NUMBER_CATEGORIES[cat].code(dec);
    if (touched.has('h')) p.h = h;
    if (touched.has('v')) p.v = v;
    if (touched.has('wrap')) p.wrap = wrap;
    if (touched.has('bold')) p.bold = bold;
    if (touched.has('italic')) p.italic = italic;
    if (touched.has('size')) p.size = size;
    onApply(p);
  };
  return (
    <Dialog title="Format Cells" width={500} onClose={onClose} id="excel.home_tab_number_format_font_alignment_borders" onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); ok(); } }}>
      <div>
        <div className="xl-dlg-tabs">
          {(['number', 'alignment', 'font'] as const).map(k => <button key={k} className={`xl-dlg-tab${tab === k ? ' active' : ''}`} onClick={() => setTab(k)}>{k[0].toUpperCase() + k.slice(1)}</button>)}
          <button className="xl-dlg-tab" disabled>Border</button><button className="xl-dlg-tab" disabled>Fill</button><button className="xl-dlg-tab" disabled>Protection</button>
        </div>
        <div className="xl-dlg-page" style={{ height: 230 }}>
          {tab === 'number' && (
            <div style={{ display: 'flex', gap: 12 }}>
              <div><div>Category:</div><div className="xl-goto-list" style={{ width: 120, height: 190 }}>{NUMBER_CATEGORIES.map((c, i) => <div key={c.label} className={i === cat ? 'sel' : ''} onMouseDown={() => { setCat(i); t('num'); }}>{c.label}</div>)}</div></div>
              <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 8 }}>
                <fieldset style={{ border: '1px solid #d0d0d0' }}><legend>Sample</legend>{sample}</fieldset>
                {['Number', 'Currency', 'Accounting', 'Percentage', 'Scientific'].includes(NUMBER_CATEGORIES[cat].label) && (
                  <div className="xl-dlg-field"><label style={{ width: 100 }}>Decimal places:</label><input type="text" value={String(dec)} onChange={e => { setDec(Math.max(0, Math.min(30, Number(e.target.value) || 0))); t('num'); }} style={{ width: 50, flex: 'none' }} /></div>
                )}
              </div>
            </div>
          )}
          {tab === 'alignment' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              <div className="xl-dlg-field"><label>Horizontal:</label><select value={h} onChange={e => { setH(e.target.value); t('h'); }}><option value="general">General</option><option value="left">Left (Indent)</option><option value="center">Center</option><option value="right">Right (Indent)</option><option value="fill">Fill</option><option value="justify">Justify</option></select></div>
              <div className="xl-dlg-field"><label>Vertical:</label><select value={v} onChange={e => { setV(e.target.value); t('v'); }}><option value="top">Top</option><option value="center">Center</option><option value="bottom">Bottom</option></select></div>
              <label className="xl-check"><input type="checkbox" checked={wrap} onChange={e => { setWrap(e.target.checked); t('wrap'); }} />Wrap text</label>
            </div>
          )}
          {tab === 'font' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              <label className="xl-check"><input type="checkbox" checked={bold} onChange={e => { setBold(e.target.checked); t('bold'); }} />Bold</label>
              <label className="xl-check"><input type="checkbox" checked={italic} onChange={e => { setItalic(e.target.checked); t('italic'); }} />Italic</label>
              <div className="xl-dlg-field"><label>Size:</label><input type="text" value={String(size)} onChange={e => { setSize(Number(e.target.value) || 11); t('size'); }} style={{ width: 50, flex: 'none' }} /></div>
            </div>
          )}
        </div>
      </div>
      <div className="xl-dlg-btns"><button className="xl-dlg-btn default" onClick={ok}>OK</button><button className="xl-dlg-btn" onClick={onClose}>Cancel</button></div>
    </Dialog>
  );
}

export function InsertFunction({ names, onOk, onClose }: { names: string[]; onOk: (fn: string) => void; onClose: () => void }) {
  const ID = 'excel.formulas_autosum_function_library';
  const [q, setQ] = useState('');
  const [sel, setSel] = useState(names.indexOf('SUM'));
  const list = q ? names.filter(n => n.includes(q.toUpperCase())) : names;
  return (
    <Dialog title="Insert Function" width={420} onClose={onClose} id={ID} onKeyDown={e => { if (e.key === 'Enter' && list[sel]) { e.preventDefault(); onOk(list[sel]); } }}>
      <div className="xl-dlg-field"><label style={{ width: 120 }}><u>S</u>earch for a function:</label><input type="text" value={q} onChange={e => { setQ(e.target.value); setSel(0); }} autoFocus /></div>
      <div>Select a functio<u>n</u>:</div>
      <div className="xl-goto-list" style={{ height: 160 }}>{list.map((n, i) => <div key={n} className={i === sel ? 'sel' : ''} onMouseDown={() => setSel(i)} onDoubleClick={() => onOk(n)}>{n}</div>)}</div>
      <div className="xl-dlg-btns"><button className="xl-dlg-btn default" disabled={!list[sel]} onClick={() => onOk(list[sel])}>OK</button><button className="xl-dlg-btn" onClick={onClose}>Cancel</button></div>
    </Dialog>
  );
}

// ---------------------------------------------------------------- sheet tab menu dialogs
// Move or Copy, Protect Sheet (with Confirm Password), Unprotect Sheet, and the Macro dialog that View Code
// and Alt+F8 show (VBA is blocked by policy: no macros, nothing that runs, creates or edits one).
const TABS_ID = 'excel.sheet_tabs';

export function MoveCopyDialog({ book, sheets, current, onOk, onClose }: { book: string; sheets: string[]; current: number; onOk: (before: number, copy: boolean) => void; onClose: () => void }) {
  const [sel, setSel] = useState(current);
  const [copy, setCopy] = useState(false);
  const listRef = useRef<HTMLDivElement>(null);
  useEffect(() => { listRef.current?.focus(); }, []);
  const items = [...sheets, '(move to end)'];
  const ok = () => onOk(sel, copy);
  return (
    <Dialog title="Move or Copy" width={290} onClose={onClose} id={TABS_ID}
      onKeyDown={e => {
        if (e.key === 'Enter') { e.preventDefault(); ok(); return; }
        if (e.altKey && e.key.toLowerCase() === 'c') { e.preventDefault(); setCopy(c => !c); return; }
        if (e.altKey && e.key.toLowerCase() === 'b') { e.preventDefault(); listRef.current?.focus(); return; }
      }}>
      <div>Move selected sheets</div>
      <div className="xl-mc-group">
        <label><u>T</u>o book:</label>
        <select value={book} data-control-id={TABS_ID} data-control-item="To book" onChange={() => {}}><option value={book}>{book}</option></select>
      </div>
      <div className="xl-mc-group">
        <label><u>B</u>efore sheet:</label>
        <div ref={listRef} className="xl-goto-list xl-mc-list" tabIndex={0} role="listbox" aria-label="Before sheet" data-control-id={TABS_ID} data-control-item="Before sheet"
          onKeyDown={e => {
            if (e.key === 'ArrowDown') { e.preventDefault(); setSel(s => Math.min(items.length - 1, s + 1)); }
            if (e.key === 'ArrowUp') { e.preventDefault(); setSel(s => Math.max(0, s - 1)); }
            if (e.key === 'Home') { e.preventDefault(); setSel(0); }
            if (e.key === 'End') { e.preventDefault(); setSel(items.length - 1); }
          }}>
          {items.map((n, i) => <div key={`${i}:${n}`} role="option" aria-selected={i === sel} className={i === sel ? 'sel' : ''} data-control-id={TABS_ID} data-control-item={`Before sheet: ${n}`} onMouseDown={() => setSel(i)} onDoubleClick={() => onOk(i, copy)}>{n}</div>)}
        </div>
      </div>
      <label className="xl-check"><input type="checkbox" checked={copy} onChange={e => setCopy(e.target.checked)} data-control-id={TABS_ID} data-control-item="Create a copy" /><span><u>C</u>reate a copy</span></label>
      <div className="xl-dlg-btns">
        <button className="xl-dlg-btn default" data-control-id={TABS_ID} data-control-item="OK" onClick={ok}>OK</button>
        <button className="xl-dlg-btn" data-control-id={TABS_ID} data-control-item="Cancel" onClick={onClose}>Cancel</button>
      </div>
    </Dialog>
  );
}

export interface ProtectFlag { attr: string; label: string; allowedByDefault: boolean }

/** Protect Sheet: protect worksheet and contents of locked cells, optional password, allow-list; a password asks
 *  for confirmation (Confirm Password) before the sheet is protected */
export function ProtectSheetDialog({ flags, onOk, onClose, alert }: { flags: ProtectFlag[]; onOk: (password: string, allow: Record<string, boolean>) => void; onClose: () => void; alert: (text: string) => Promise<unknown> }) {
  const [on, setOn] = useState(true);
  const [pw, setPw] = useState('');
  const [allow, setAllow] = useState<Record<string, boolean>>(() => Object.fromEntries(flags.map(f => [f.attr, f.allowedByDefault])));
  const [focusRow, setFocusRow] = useState(0);
  const [confirm, setConfirm] = useState(false);
  const pwRef = useRef<HTMLInputElement>(null);
  useEffect(() => { pwRef.current?.focus(); }, []);
  const ok = () => { if (!on) return; if (pw) setConfirm(true); else onOk('', allow); };
  if (confirm) {
    return <ConfirmPasswordDialog onClose={() => { setConfirm(false); setTimeout(() => pwRef.current?.focus(), 0); }}
      onOk={async again => { if (again !== pw) { await alert('Confirmation password is not identical.'); return false; } onOk(pw, allow); return true; }} />;
  }
  return (
    <Dialog title="Protect Sheet" width={292} onClose={onClose} id={TABS_ID}
      onKeyDown={e => {
        if (e.key === 'Enter' && !(e.target instanceof HTMLButtonElement)) { e.preventDefault(); ok(); return; }
        if (e.altKey && e.key.toLowerCase() === 'c') { e.preventDefault(); setOn(v => !v); return; }
        if (e.altKey && e.key.toLowerCase() === 'p') { e.preventDefault(); pwRef.current?.focus(); return; }
      }}>
      <label className="xl-check"><input type="checkbox" checked={on} onChange={e => setOn(e.target.checked)} data-control-id={TABS_ID} data-control-item="Protect worksheet and contents of locked cells" /><span>Protect worksheet and <u>c</u>ontents of locked cells</span></label>
      <div className="xl-mc-group">
        <label><u>P</u>assword to unprotect sheet:</label>
        <input ref={pwRef} type="password" className="xl-dlg-input" value={pw} onChange={e => setPw(e.target.value)} maxLength={255} data-control-id={TABS_ID} data-control-item="Password to unprotect sheet" autoComplete="off" />
      </div>
      <div className="xl-mc-group">
        <label>All<u>o</u>w all users of this worksheet to:</label>
        <div className="xl-goto-list xl-prot-list" role="listbox" aria-label="Allow all users of this worksheet to" tabIndex={0} data-control-id={TABS_ID} data-control-item="Allow all users of this worksheet to"
          onKeyDown={e => {
            if (e.key === 'ArrowDown') { e.preventDefault(); setFocusRow(r => Math.min(flags.length - 1, r + 1)); }
            if (e.key === 'ArrowUp') { e.preventDefault(); setFocusRow(r => Math.max(0, r - 1)); }
            if (e.key === ' ') { e.preventDefault(); const f = flags[focusRow]; setAllow(a => ({ ...a, [f.attr]: !a[f.attr] })); }
          }}>
          {flags.map((f, i) => (
            <label key={f.attr} className={`xl-prot-row${i === focusRow ? ' focus' : ''}`} onMouseDown={() => setFocusRow(i)}>
              <input type="checkbox" checked={!!allow[f.attr]} onChange={e => setAllow(a => ({ ...a, [f.attr]: e.target.checked }))} tabIndex={-1} data-control-id={TABS_ID} data-control-item={`Allow: ${f.label}`} />
              <span>{f.label}</span>
            </label>
          ))}
        </div>
      </div>
      <div className="xl-dlg-btns">
        <button className="xl-dlg-btn default" data-control-id={TABS_ID} data-control-item="OK" disabled={!on} onClick={ok}>OK</button>
        <button className="xl-dlg-btn" data-control-id={TABS_ID} data-control-item="Cancel" onClick={onClose}>Cancel</button>
      </div>
    </Dialog>
  );
}

function ConfirmPasswordDialog({ onOk, onClose }: { onOk: (pw: string) => Promise<boolean>; onClose: () => void }) {
  const [v, setV] = useState('');
  const ref = useRef<HTMLInputElement>(null);
  useEffect(() => { ref.current?.focus(); }, []);
  const ok = async () => { if (!(await onOk(v))) { setV(''); setTimeout(() => ref.current?.focus(), 0); } };
  return (
    <Dialog title="Confirm Password" width={292} onClose={onClose} id={TABS_ID} onKeyDown={e => { if (e.key === 'Enter' && !(e.target instanceof HTMLButtonElement)) { e.preventDefault(); void ok(); } }}>
      <div className="xl-mc-group">
        <label><u>R</u>eenter password to proceed.</label>
        <input ref={ref} type="password" className="xl-dlg-input" value={v} onChange={e => setV(e.target.value)} data-control-id={TABS_ID} data-control-item="Reenter password to proceed" autoComplete="off" />
      </div>
      <div className="xl-prot-caution">Caution: If you lose or forget the password, it cannot be recovered. It is advisable to keep a list of passwords and their corresponding workbook and sheet names in a safe place. (Remember that passwords are case-sensitive.)</div>
      <div className="xl-dlg-btns">
        <button className="xl-dlg-btn default" data-control-id={TABS_ID} data-control-item="OK" onClick={() => void ok()}>OK</button>
        <button className="xl-dlg-btn" data-control-id={TABS_ID} data-control-item="Cancel" onClick={onClose}>Cancel</button>
      </div>
    </Dialog>
  );
}

export function UnprotectSheetDialog({ onOk, onClose }: { onOk: (pw: string) => void; onClose: () => void }) {
  const [v, setV] = useState('');
  const ref = useRef<HTMLInputElement>(null);
  useEffect(() => { ref.current?.focus(); }, []);
  return (
    <Dialog title="Unprotect Sheet" width={250} onClose={onClose} id={TABS_ID} onKeyDown={e => { if (e.key === 'Enter' && !(e.target instanceof HTMLButtonElement)) { e.preventDefault(); onOk(v); } }}>
      <div className="xl-dlg-field"><label style={{ width: 60 }}><u>P</u>assword:</label><input ref={ref} type="password" className="xl-dlg-input" value={v} onChange={e => setV(e.target.value)} data-control-id={TABS_ID} data-control-item="Password" autoComplete="off" /></div>
      <div className="xl-dlg-btns">
        <button className="xl-dlg-btn default" data-control-id={TABS_ID} data-control-item="OK" onClick={() => onOk(v)}>OK</button>
        <button className="xl-dlg-btn" data-control-id={TABS_ID} data-control-item="Cancel" onClick={onClose}>Cancel</button>
      </div>
    </Dialog>
  );
}

/** View Code / Alt+F8: Excel's Macro dialog with nothing in it (VBA is disabled by policy on this PC) */
export function MacroDialog({ onClose }: { onClose: () => void }) {
  const ref = useRef<HTMLInputElement>(null);
  useEffect(() => { ref.current?.focus(); }, []);
  return (
    <Dialog title="Macro" width={420} onClose={onClose} id={TABS_ID}>
      <div className="xl-macro">
        <div className="xl-macro-left">
          <label><u>M</u>acro name:</label>
          <input ref={ref} type="text" className="xl-dlg-input" aria-label="Macro name" data-control-id={TABS_ID} data-control-item="Macro name" />
          <div className="xl-goto-list xl-macro-list" role="listbox" aria-label="Macros" />
          <div className="xl-dlg-field"><label style={{ width: 62 }}>M<u>a</u>cros in:</label><select disabled aria-label="Macros in"><option>All Open Workbooks</option></select></div>
          <fieldset className="xl-macro-desc"><legend>Description</legend></fieldset>
        </div>
        <div className="xl-macro-btns">
          <button className="xl-dlg-btn" disabled><u>R</u>un</button>
          <button className="xl-dlg-btn" disabled><u>S</u>tep Into</button>
          <button className="xl-dlg-btn" disabled><u>E</u>dit</button>
          <button className="xl-dlg-btn" disabled>Cr<u>e</u>ate</button>
          <button className="xl-dlg-btn" disabled><u>D</u>elete</button>
          <button className="xl-dlg-btn" disabled><u>O</u>ptions...</button>
          <div style={{ flex: 1 }} />
          <button className="xl-dlg-btn" data-control-id={TABS_ID} data-control-item="Cancel" onClick={onClose}>Cancel</button>
        </div>
      </div>
    </Dialog>
  );
}
