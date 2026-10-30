// The in-page clipboard: { text, html? }.
// Native copy/cut events anywhere are mirrored into it, and pastes into plain inputs are served from it.
// The grid calls read()/write() for its own copy and paste.
import { createStore } from './store';

export interface ClipData {
  text: string;
  html?: string;
  /** which app put it there */
  source?: string;
}

const store = createStore<{ data: ClipData | null; seq: number }>({ data: null, seq: 0 });

export const clipboard = {
  read(): ClipData | null { return store.get().data; },
  write(d: ClipData) {
    store.set(s => ({ data: d, seq: s.seq + 1 }));
    // best effort: keep the browser's clipboard in step for native paste paths
    try { if (navigator.clipboard && d.text) void navigator.clipboard.writeText(d.text).then(() => { lastSystemText = d.text; }, () => {}); } catch { /* ignore */ }
  },
  clear() { store.set(s => ({ data: null, seq: s.seq + 1 })); },
  subscribe: store.subscribe,
  store,
};

// Inline formatting and list context kept when a selection is serialised. Range.cloneContents() drops the
// ancestors of a selection that lies wholly inside them (a selected word inside <b> clones as bare text), so
// the formatting elements between the selection and its editing / reading root are re-wrapped around it.
const KEEP_ANCESTORS = new Set(['B', 'STRONG', 'I', 'EM', 'U', 'S', 'STRIKE', 'SUB', 'SUP', 'A', 'FONT', 'SPAN', 'CODE', 'LI', 'UL', 'OL', 'BLOCKQUOTE', 'P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6']);
const BLOCKS = new Set(['LI', 'UL', 'OL', 'BLOCKQUOTE', 'P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'DIV']);
const ROOT_SEL = '[contenteditable="true"], [contenteditable=""], [data-clip-root]';

function wrapRange(r: Range): Node {
  let frag: Node = r.cloneContents();
  const start = r.commonAncestorContainer;
  let el: Element | null = start.nodeType === Node.ELEMENT_NODE ? start as Element : start.parentElement;
  const root = el?.closest(ROOT_SEL) ?? null;
  // Block ancestors (paragraph, heading, list) are kept only when the selection spans several of their children:
  // words copied from inside one paragraph or list item paste as text with its character formatting, not as a
  // new paragraph (and the editor's paste-options markers stay in one run).
  const spansBlocks = start.nodeType === Node.ELEMENT_NODE && BLOCKS.has((start as Element).tagName);
  const chain: Element[] = [];
  for (let e = el; e && e !== root && e !== document.body; e = e.parentElement) {
    if (!KEEP_ANCESTORS.has(e.tagName)) continue;
    if (BLOCKS.has(e.tagName) && !spansBlocks) continue;
    chain.push(e);
  }
  let tags = chain.map(e => e.tagName);
  for (const e of chain) {
    const c = e.cloneNode(false) as Element;
    c.removeAttribute('id'); c.removeAttribute('contenteditable'); c.removeAttribute('data-control-id');
    if (e.tagName === 'A') { const href = e.getAttribute('href'); if (href) c.setAttribute('href', href); }
    c.appendChild(frag);
    frag = c;
  }
  // formatting that comes from a stylesheet rather than an element (classes in received HTML)
  const base = (start.nodeType === Node.ELEMENT_NODE ? start as Element : start.parentElement);
  if (base) {
    const cs = getComputedStyle(base);
    const wrap = (tag: string) => { const w = document.createElement(tag); w.appendChild(frag); frag = w; tags = [...tags, tag.toUpperCase()]; };
    const weight = parseInt(cs.fontWeight, 10) || (cs.fontWeight === 'bold' ? 700 : 400);
    const rootWeight = root ? parseInt(getComputedStyle(root).fontWeight, 10) || 400 : 400;
    if (weight >= 600 && rootWeight < 600 && !tags.some(t => t === 'B' || t === 'STRONG' || /^H\d$/.test(t))) wrap('b');
    if (cs.fontStyle === 'italic' && !tags.some(t => t === 'I' || t === 'EM')) wrap('i');
    if (/underline/.test(cs.textDecorationLine) && !tags.some(t => t === 'U' || t === 'A')) wrap('u');
  }
  return frag;
}

/** The current selection as { text, html }, with the formatting (bold, italic, underline, links, lists) it sits in. */
export function selectionHtml(): { text: string; html: string } | null {
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return null;
  const div = document.createElement('div');
  for (let i = 0; i < sel.rangeCount; i++) div.appendChild(wrapRange(sel.getRangeAt(i)));
  return { text: sel.toString(), html: div.innerHTML };
}

function isTextField(el: EventTarget | null): el is HTMLInputElement | HTMLTextAreaElement {
  if (el instanceof HTMLTextAreaElement) return true;
  return el instanceof HTMLInputElement && /^(text|search|email|url|tel|password|number|)$/i.test(el.type);
}

// The last system-clipboard text the page knows of: seen at a paste, written by a native copy/cut, or written by
// clipboard.write. A paste whose system text differs from both this and the in-page text was put there from outside
// the page (another program, or an automation driver), so it wins over a stale in-page copy;
// known system text is left to the in-page copy, which carries formatting.
let lastSystemText: string | null = null;
const norm = (t: string) => t.replace(/\s+/g, ' ').trim();

/** Adopt the system clipboard into the in-page one when something outside the page changed it. */
export function adoptSystemClipboard(dt: DataTransfer | null) {
  let text = '', html = '';
  try { text = dt?.getData('text/plain') ?? ''; html = dt?.getData('text/html') ?? ''; } catch { /* unreadable */ }
  if (!text) return;
  const cur = clipboard.read();
  // before the page knows any system text (fresh load, writeText refused), an in-page copy wins: the system text
  // may predate it (a browser without clipboard-write access keeps an older clipboard)
  const fresh = text !== lastSystemText && (!cur || (lastSystemText !== null && norm(cur.text) !== norm(text)));
  lastSystemText = text;
  if (fresh) store.set(s => ({ data: { text, html: html || undefined, source: 'system' }, seq: s.seq + 1 }));
}

export function installClipboard() {
  // before any app handler: a paste from outside the page replaces the in-page clipboard
  window.addEventListener('paste', e => adoptSystemClipboard(e.clipboardData), true);
  // components may write through a DOM event as well ('sheet:clipboard-write' { text, html? })
  window.addEventListener('sheet:clipboard-write', e => {
    const d = (e as CustomEvent).detail ?? {};
    clipboard.write({ text: String(d.text ?? ''), html: d.html, source: d.source ?? 'app' });
  });
  const mirror = (e: ClipboardEvent) => {
    // runs after app handlers (bubble phase on window): take what the app put in clipboardData if any
    const dt = e.clipboardData;
    let text = '', html = '';
    try { text = dt?.getData('text/plain') ?? ''; html = dt?.getData('text/html') ?? ''; } catch { /* read-only */ }
    if (!text && !html) {
      const t = e.target;
      if (isTextField(t)) {
        text = t.value.slice(t.selectionStart ?? 0, t.selectionEnd ?? 0);
      } else {
        const s = selectionHtml();
        if (s) { text = s.text; html = s.html; }
      }
    }
    // a native copy/cut puts this text on the system clipboard itself
    if (text) lastSystemText = text;
    if (text || html) clipboard.write({ text, html: html || undefined, source: 'selection' });
  };
  window.addEventListener('copy', mirror);
  window.addEventListener('cut', mirror);
  // paste into plain fields: serve from the in-page clipboard (works without OS clipboard access)
  window.addEventListener('paste', e => {
    if (e.defaultPrevented) return;
    const d = clipboard.read();
    if (!d?.text) return;
    const t = e.target;
    if (isTextField(t) && !t.readOnly && !t.disabled) {
      e.preventDefault();
      const single = t instanceof HTMLInputElement ? d.text.replace(/\r?\n/g, ' ') : d.text;
      document.execCommand('insertText', false, single);
    } else if (t instanceof HTMLElement && t.isContentEditable) {
      const sys = e.clipboardData?.getData('text/plain');
      if (!sys) { e.preventDefault(); document.execCommand(d.html ? 'insertHTML' : 'insertText', false, d.html ?? d.text); }
    }
  });
}
