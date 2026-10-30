// Popup menus: the modern rounded menu (with an optional icon row) and the classic compact menu.
// One menu stack open at a time.
import { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import { createStore, useStore } from './store';
import { ChevronRight, CheckGlyph, DotGlyph } from './icons';

export interface MenuItem {
  label?: string;
  icon?: ReactNode;
  shortcut?: string;
  onClick?: () => void;
  disabled?: boolean;
  checked?: boolean;
  radio?: boolean;
  bold?: boolean;
  submenu?: MenuItem[] | (() => MenuItem[]);
  separator?: boolean;
  /** data-control-id for this item (defaults to the menu's controlId) */
  controlId?: string;
  /** label used for `<control_id>[<label>]` targets (defaults to label) */
  itemLabel?: string;
  /** keep the menu open after clicking (checkbox toggles in some menus) */
  keepOpen?: boolean;
  /** tooltip */
  tip?: string;
}

export interface MenuSpec {
  x: number;
  y: number;
  items: MenuItem[];
  variant?: 'modern' | 'classic';
  /** icon buttons drawn above the items (Explorer file menu: Cut, Copy, Rename, Share, Delete) */
  iconRow?: MenuItem[];
  controlId?: string;
  onClose?: () => void;
  /** open upward from y (taskbar menus) */
  above?: boolean;
  /** align right edge to x */
  alignRight?: boolean;
  minWidth?: number;
  /** element the menu belongs to; clicking it again closes the menu without reopening */
  anchor?: Element | null;
  /** keyboard: first item highlighted */
  focusFirst?: boolean;
  /** Explorer command-bar flyout (Sort / View / New / See more): tighter separators, centred under x */
  flyout?: boolean;
  /** shell context menu placement (desktop, Explorer items): Windows fits it as if 124 px taller, flips above
   *  the cursor when that fits, else pushes it up against the work area (taskbar 48 px) */
  shellContext?: boolean;
  /** extra class on the root panel (and its submenus) */
  className?: string;
}

const store = createStore<{ menu: MenuSpec | null; seq: number }>({ menu: null, seq: 0 });

export const menus = {
  open(spec: MenuSpec) { const prev = store.get().menu; store.set(s => ({ menu: spec, seq: s.seq + 1 })); prev?.onClose?.(); },
  close() { const m = store.get().menu; if (!m) return; store.set({ menu: null }); m.onClose?.(); },
  isOpen() { return !!store.get().menu; },
  current() { return store.get().menu; },
};

export function MenuHost() {
  const { menu, seq } = useStore(store, s => ({ menu: s.menu, seq: s.seq }));
  useEffect(() => {
    if (!menu) return;
    const down = (e: MouseEvent) => {
      const t = e.target as Element;
      if (t.closest?.('.menu-root')) return;
      if (menu.anchor && menu.anchor.contains(t)) { e.stopPropagation(); e.preventDefault(); menus.close(); return; }
      menus.close();
    };
    const blur = () => menus.close();
    window.addEventListener('mousedown', down, true);
    window.addEventListener('blur', blur);
    return () => { window.removeEventListener('mousedown', down, true); window.removeEventListener('blur', blur); };
  }, [menu]);
  if (!menu) return null;
  return <MenuPanel key={seq} spec={menu} level={0} items={menu.items} x={menu.x} y={menu.y} />;
}

function resolve(sub: MenuItem['submenu']): MenuItem[] { return typeof sub === 'function' ? sub() : sub ?? []; }

function MenuPanel({ spec, items, x, y, level, parentRect, onCloseSub }: { spec: MenuSpec; items: MenuItem[]; x: number; y: number; level: number; parentRect?: DOMRect; onCloseSub?: () => void }) {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState({ x, y, ready: false, flip: false });
  const [hi, setHi] = useState<number>(spec.focusFirst && level === 0 ? firstEnabled(items) : -1);
  const [openSub, setOpenSub] = useState<number>(-1);
  const subTimer = useRef<number>();
  const variant = spec.variant ?? 'modern';
  const flyout = level === 0 && !!spec.flyout;
  const shellCtx = level === 0 && (spec.shellContext ?? /^(shell\.desktop\.desktop_right_click_menu|explorer\.content\.item_right_click_menu)$/.test(spec.controlId ?? ''));

  useLayoutEffect(() => {
    const el = ref.current; if (!el) return;
    const r = el.getBoundingClientRect();
    const W = window.innerWidth, H = window.innerHeight;
    let nx = x, ny = y, flip = false;
    if (level === 0) {
      if (spec.alignRight) nx = x - r.width;
      if (flyout) nx = Math.max(7, x - r.width / 2);
      if (spec.above) ny = y - r.height;
      if (nx + r.width > W - 4) nx = Math.max(4, x - r.width);
      if (shellCtx) {
        // a menu that would come within 124 px of the taskbar opens above the cursor with the
        // icon row moved below it (items end at the cursor); if that does not fit either it is pushed up instead
        const bottom = H - 48;
        const rowH = (el.querySelector('.menu-iconrow') as HTMLElement | null)?.offsetHeight ?? 0;
        const gap = spec.controlId === 'shell.desktop.desktop_right_click_menu' ? 80 : 124;
        if (y + r.height + gap > bottom) {
          if (r.height - rowH <= y) { ny = y - (r.height - rowH); flip = true; } else ny = Math.max(0, bottom - r.height - gap);
        }
      } else if (ny + r.height > H - 4) ny = Math.max(4, spec.above ? 4 : y - r.height);
    } else if (parentRect) {
      nx = parentRect.right - 4; ny = parentRect.top - (variant === 'modern' ? 4 : 3);
      if (nx + r.width > W - 4) nx = parentRect.left - r.width + 4;
      if (ny + r.height > H - 4) ny = Math.max(4, H - 4 - r.height);
    }
    // shortcut column: left-aligned, as wide as the widest shortcut (Windows 11 context menus)
    let sc = 0;
    el.querySelectorAll<HTMLElement>(':scope > .menu-item > .menu-shortcut').forEach(n => { sc = Math.max(sc, n.scrollWidth); });
    el.style.setProperty('--menu-sc-w', `${Math.ceil(sc)}px`);
    setPos({ x: Math.max(0, nx), y: Math.max(0, ny), ready: true, flip });
  }, [x, y, level, parentRect, spec.alignRight, spec.above, variant, flyout, shellCtx]);

  // keyboard (only the deepest open panel listens)
  useEffect(() => {
    if (openSub >= 0) return;
    const key = (e: KeyboardEvent) => {
      if (e.key === 'ArrowDown') { e.preventDefault(); e.stopPropagation(); setHi(h => nextEnabled(items, h, 1)); }
      else if (e.key === 'ArrowUp') { e.preventDefault(); e.stopPropagation(); setHi(h => nextEnabled(items, h, -1)); }
      else if (e.key === 'ArrowRight') { const it = items[hi]; if (it?.submenu && !it.disabled) { e.preventDefault(); e.stopPropagation(); setOpenSub(hi); } }
      else if (e.key === 'ArrowLeft' || e.key === 'Escape') {
        e.preventDefault(); e.stopPropagation();
        if (level > 0) onCloseSub?.(); else menus.close();
      } else if (e.key === 'Enter' || e.key === ' ') {
        const it = items[hi]; if (!it) return;
        e.preventDefault(); e.stopPropagation();
        activate(it, hi);
      } else if (e.key.length === 1 && !e.ctrlKey && !e.altKey && !e.metaKey) {
        const ch = e.key.toLowerCase();
        const idx = items.findIndex((it, i) => i > hi && !it.separator && !it.disabled && it.label?.toLowerCase().startsWith(ch));
        const idx2 = idx >= 0 ? idx : items.findIndex(it => !it.separator && !it.disabled && it.label?.toLowerCase().startsWith(ch));
        if (idx2 >= 0) { e.preventDefault(); e.stopPropagation(); setHi(idx2); }
      }
    };
    window.addEventListener('keydown', key, true);
    return () => window.removeEventListener('keydown', key, true);
  });

  function activate(it: MenuItem, i: number) {
    if (it.disabled || it.separator) return;
    if (it.submenu) { setOpenSub(i); return; }
    if (!it.keepOpen) menus.close();
    it.onClick?.();
  }

  const iconRow = level === 0 && spec.iconRow ? (
          <div className={`menu-iconrow${pos.flip ? ' bottom' : ''}`}>
            {spec.iconRow.map((it, i) => (
              <button key={i} className="menu-iconbtn" disabled={it.disabled} data-tip={it.tip ?? it.label}
                data-control-id={it.controlId ?? spec.controlId} data-control-item={it.itemLabel ?? it.label}
                onClick={() => { if (it.disabled) return; menus.close(); it.onClick?.(); }}>
                {it.icon}
                <span className="menu-iconbtn-label">{it.label}</span>
              </button>
            ))}
          </div>
  ) : null;
  const rowRefs = useRef<(HTMLDivElement | null)[]>([]);
  const subRect = openSub >= 0 ? rowRefs.current[openSub]?.getBoundingClientRect() : undefined;
  const hasIcons = items.some(i => i.icon || i.checked !== undefined);
  // a menu with both radio / check marks and icons (Explorer View) gets a separate icon column after the marks
  const twoCols = variant === 'modern' && items.some(i => i.checked !== undefined) && items.some(i => i.icon);
  const hasShortcuts = items.some(i => i.shortcut);

  return (
    <>
      <div ref={ref} className={`menu-root menu-${variant}${flyout ? ' menu-flyout' : ''}${shellCtx ? ' menu-shellctx' : ''}${hasShortcuts ? ' has-sc' : ''}${spec.className ? ` ${spec.className}` : ''}`} data-control-id={level === 0 ? spec.controlId : undefined}
        style={{ left: pos.x, top: pos.y, visibility: pos.ready ? 'visible' : 'hidden', minWidth: spec.minWidth }}
        onMouseDown={e => e.preventDefault()} onContextMenu={e => e.preventDefault()} role="menu">
        {!pos.flip && iconRow}
        {items.map((it, i) => it.separator
          ? <div key={i} className="menu-sep" />
          : (
            <div key={i} ref={el => { rowRefs.current[i] = el; }}
              className={`menu-item${hi === i ? ' hi' : ''}${it.disabled ? ' disabled' : ''}${it.bold ? ' bold' : ''}`}
              role="menuitem" aria-disabled={it.disabled}
              data-control-id={it.controlId ?? spec.controlId} data-control-item={it.itemLabel ?? it.label}
              data-tip={it.tip}
              onMouseEnter={() => {
                setHi(i);
                window.clearTimeout(subTimer.current);
                subTimer.current = window.setTimeout(() => setOpenSub(it.submenu && !it.disabled ? i : -1), it.submenu ? 250 : 350);
              }}
              onMouseLeave={() => { if (openSub !== i) setHi(-1); }}
              onClick={() => activate(it, i)}>
              {(hasIcons || variant === 'classic') && (
                <span className="menu-icon">
                  {it.checked ? (it.radio ? <DotGlyph size={16} /> : <CheckGlyph size={16} />) : twoCols ? null : it.icon}
                </span>
              )}
              {twoCols && <span className="menu-icon menu-icon2">{it.icon}</span>}
              <span className="menu-label">{it.label}</span>
              {it.shortcut && <span className="menu-shortcut">{it.shortcut}</span>}
              {it.submenu && <span className="menu-arrow"><ChevronRight size={12} /></span>}
            </div>
          ))}
        {pos.flip && iconRow}
      </div>
      {openSub >= 0 && items[openSub]?.submenu && subRect && (
        <MenuPanel spec={spec} items={resolve(items[openSub].submenu)} x={0} y={0} level={level + 1} parentRect={subRect}
          onCloseSub={() => setOpenSub(-1)} />
      )}
    </>
  );
}

function firstEnabled(items: MenuItem[]) { return items.findIndex(i => !i.separator && !i.disabled); }
function nextEnabled(items: MenuItem[], from: number, dir: 1 | -1) {
  const n = items.length;
  for (let k = 1; k <= n; k++) {
    const i = ((from + dir * k) % n + n) % n;
    if (!items[i].separator && !items[i].disabled) return i;
  }
  return from;
}

// ------------------------------------------------------------------ tooltips

const tipStore = createStore<{ text: string | null; x: number; y: number }>({ text: null, x: 0, y: 0 });

export function TooltipHost() {
  const t = useStore(tipStore, s => s);
  useEffect(() => {
    let timer = 0, current: Element | null = null, hideTimer = 0;
    const over = (e: MouseEvent) => {
      const el = (e.target as Element).closest?.('[data-tip]');
      if (el === current) return;
      current = el;
      window.clearTimeout(timer); window.clearTimeout(hideTimer);
      tipStore.set({ text: null });
      if (!el) return;
      const text = el.getAttribute('data-tip');
      if (!text) return;
      const x = e.clientX, y = e.clientY;
      timer = window.setTimeout(() => {
        if (current !== el || menus.isOpen() && !el.closest('.menu-root')) return;
        const r = el.getBoundingClientRect();
        const below = el.closest('.taskbar') ? false : true;
        tipStore.set({ text, x: el.closest('.taskbar') ? r.left + r.width / 2 : x, y: below ? Math.max(r.bottom, y) + (el.closest('.taskbar') ? 0 : 18) : r.top - 8 });
        hideTimer = window.setTimeout(() => tipStore.set({ text: null }), 5000);
      }, 600);
    };
    const hide = () => { window.clearTimeout(timer); current = null; tipStore.set({ text: null }); };
    window.addEventListener('mouseover', over, true);
    window.addEventListener('mousedown', hide, true);
    window.addEventListener('keydown', hide, true);
    return () => { window.removeEventListener('mouseover', over, true); window.removeEventListener('mousedown', hide, true); window.removeEventListener('keydown', hide, true); };
  }, []);
  if (!t.text) return null;
  return <TipBox text={t.text} x={t.x} y={t.y} />;
}

function TipBox({ text, x, y }: { text: string; x: number; y: number }) {
  const ref = useRef<HTMLDivElement>(null);
  const [p, setP] = useState({ x, y, ok: false });
  useLayoutEffect(() => {
    const r = ref.current!.getBoundingClientRect();
    const W = window.innerWidth, H = window.innerHeight;
    const aboveTaskbar = y > H - 48;
    let nx = aboveTaskbar ? x - r.width / 2 : x;
    let ny = aboveTaskbar ? y - r.height : y;
    if (nx + r.width > W - 2) nx = W - 2 - r.width;
    if (ny + r.height > H - 2) ny = y - r.height - 24;
    setP({ x: Math.max(2, nx), y: Math.max(2, ny), ok: true });
  }, [x, y, text]);
  return <div ref={ref} className="tooltip" style={{ left: p.x, top: p.y, visibility: p.ok ? 'visible' : 'hidden' }}>{text.split('\n').map((l, i) => <div key={i}>{l}</div>)}</div>;
}
