// Office-drawn context menu (Excel draws its own menus: 28 px rows, underlined access keys, icon column).
// Used for the sheet tab menu. Items carry data-control-id / data-control-item like the shell's menus.
import { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { Icon } from './icons';

export interface XlMenuItem {
  label?: string;
  /** index of the access key letter in label (underlined; Alt-free key press picks the item) */
  key?: string;
  icon?: string;
  disabled?: boolean;
  separator?: boolean;
  onClick?: () => void;
  submenu?: XlMenuItem[];
  swatch?: string;
}

export interface XlMenuSpec { x: number; y: number; items: XlMenuItem[]; controlId: string; above?: boolean; onClose?: () => void }

let setSpecGlobal: ((s: XlMenuSpec | null) => void) | null = null;
export const xlMenu = {
  open(spec: XlMenuSpec) { setSpecGlobal?.(spec); },
  close() { setSpecGlobal?.(null); },
};

function underline(label: string, key?: string): ReactNode {
  if (!key) return label;
  const i = label.toLowerCase().indexOf(key.toLowerCase());
  if (i < 0) return label;
  return <>{label.slice(0, i)}<u>{label.slice(i, i + 1)}</u>{label.slice(i + 1)}</>;
}

/** mount once per Excel window */
export function XlMenuHost() {
  const [spec, setSpec] = useState<XlMenuSpec | null>(null);
  useEffect(() => {
    const mine = (s: XlMenuSpec | null) => setSpec(prev => { if (prev && prev !== s) prev.onClose?.(); return s; });
    setSpecGlobal = mine;
    return () => { if (setSpecGlobal === mine) setSpecGlobal = null; };
  }, []);
  useEffect(() => {
    if (!spec) return;
    const down = (e: MouseEvent) => { if (!(e.target as Element).closest?.('.xl-menu')) setSpec(null); };
    const blur = () => setSpec(null);
    window.addEventListener('mousedown', down, true);
    window.addEventListener('blur', blur);
    return () => { window.removeEventListener('mousedown', down, true); window.removeEventListener('blur', blur); };
  }, [spec]);
  if (!spec) return null;
  return createPortal(<Panel spec={spec} items={spec.items} x={spec.x} y={spec.y} above={spec.above} close={() => setSpec(null)} level={0} />, document.body);
}

function Panel({ spec, items, x, y, above, close, level }: { spec: XlMenuSpec; items: XlMenuItem[]; x: number; y: number; above?: boolean; close: () => void; level: number }) {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState<{ x: number; y: number } | null>(null);
  const [hi, setHi] = useState(-1);
  const [sub, setSub] = useState<{ i: number; rect: DOMRect } | null>(null);
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    const w = el.offsetWidth, h = el.offsetHeight;
    let nx = x, ny = above ? y - h : y;
    if (nx + w > window.innerWidth) nx = Math.max(0, window.innerWidth - w);
    if (ny + h > window.innerHeight) ny = Math.max(0, y - h);
    if (ny < 0) ny = 0;
    setPos({ x: nx, y: ny });
  }, [x, y, above]);
  useEffect(() => { if (level === 0) ref.current?.focus({ preventScroll: true }); }, [level]);
  const pick = (it: XlMenuItem, i: number, el?: HTMLElement) => {
    if (it.disabled || it.separator) return;
    if (it.submenu) { const r = (el ?? ref.current?.querySelectorAll<HTMLElement>('.xl-menu-item')[i])?.getBoundingClientRect(); if (r) setSub({ i, rect: r }); return; }
    close();
    it.onClick?.();
  };
  const enabled = items.map((it, i) => (!it.separator && !it.disabled ? i : -1)).filter(i => i >= 0);
  const onKey = (e: React.KeyboardEvent) => {
    e.stopPropagation();
    const k = e.key;
    if (k === 'Escape') { e.preventDefault(); close(); return; }
    if (k === 'ArrowDown' || k === 'ArrowUp') {
      e.preventDefault();
      const cur = enabled.indexOf(hi);
      const n = k === 'ArrowDown' ? enabled[(cur + 1) % enabled.length] : enabled[(cur - 1 + enabled.length) % enabled.length];
      setHi(n ?? -1);
      return;
    }
    if (k === 'Enter' || k === ' ') { e.preventDefault(); if (hi >= 0) pick(items[hi], hi); return; }
    if (k === 'ArrowRight' && hi >= 0 && items[hi].submenu) { e.preventDefault(); pick(items[hi], hi); return; }
    if (k.length === 1) {
      const i = items.findIndex(it => !it.separator && !it.disabled && it.key && it.key.toLowerCase() === k.toLowerCase());
      if (i >= 0) { e.preventDefault(); pick(items[i], i); }
    }
  };
  return (
    <div ref={ref} className="xl-menu" role="menu" tabIndex={-1} data-control-id={level === 0 ? spec.controlId : undefined}
      style={{ left: pos?.x ?? -9999, top: pos?.y ?? 0 }} onKeyDown={onKey} onMouseDown={e => e.preventDefault()} onContextMenu={e => e.preventDefault()}>
      {items.map((it, i) => it.separator ? <div key={i} className="xl-menu-sep" /> : (
        <div key={i} role="menuitem" aria-disabled={it.disabled || undefined} className={`xl-menu-item${hi === i ? ' hi' : ''}${it.disabled ? ' disabled' : ''}`}
          data-control-id={spec.controlId} data-control-item={it.label}
          onMouseEnter={e => { setHi(i); if (it.submenu && !it.disabled) setSub({ i, rect: e.currentTarget.getBoundingClientRect() }); else setSub(null); }}
          onClick={e => pick(it, i, e.currentTarget)}>
          <span className="xl-menu-icon">{it.swatch ? <span className="xl-menu-swatch" style={{ background: it.swatch }} /> : it.icon ? <Icon name={it.icon} size={16} /> : null}</span>
          <span className="xl-menu-label">{underline(it.label ?? '', it.key)}</span>
          {it.submenu && <span className="xl-menu-arrow"><Icon name="navNext" size={12} /></span>}
        </div>
      ))}
      {sub && items[sub.i]?.submenu && <Panel spec={spec} items={items[sub.i].submenu!} x={sub.rect.right - 2} y={sub.rect.top - 2} close={close} level={level + 1} />}
    </div>
  );
}
