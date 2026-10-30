// Minimal XML reader that remembers where every element sits in the source text, so the writer can
// splice new markup into the original string and leave every other byte as it was.

export interface XNode {
  name: string; // qualified name as written, e.g. "c" or "x14ac:dyDescent"
  local: string; // name without prefix
  attrs: Record<string, string>; // decoded attribute values, keyed by qualified name
  start: number; // index of '<'
  end: number; // index just after the element's last character
  openEnd: number; // index just after the start tag's '>'
  closeStart: number; // index of the closing tag's '<' (== end for self-closing)
  selfClosing: boolean;
  children: XNode[];
  parent: XNode | null;
}

export interface XDocument { text: string; root: XNode }

const ENT: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

export function decodeEntities(s: string): string {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|\w+);/g, (m, e: string) => {
    if (e[0] === '#') {
      const cp = e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return Number.isFinite(cp) ? String.fromCodePoint(cp) : m;
    }
    return ENT[e] ?? m;
  });
}

export function escText(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
export function escAttr(s: string): string {
  return escText(s).replace(/"/g, '&quot;');
}

export function parseXml(text: string): XDocument {
  const top: XNode = { name: '#doc', local: '#doc', attrs: {}, start: 0, end: text.length, openEnd: 0, closeStart: text.length, selfClosing: false, children: [], parent: null };
  let cur = top;
  let i = 0;
  const n = text.length;
  while (i < n) {
    const lt = text.indexOf('<', i);
    if (lt < 0) break;
    if (text.startsWith('<?', lt)) { i = text.indexOf('?>', lt) + 2; continue; }
    if (text.startsWith('<!--', lt)) { i = text.indexOf('-->', lt) + 3; continue; }
    if (text.startsWith('<![CDATA[', lt)) { i = text.indexOf(']]>', lt) + 3; continue; }
    if (text.startsWith('<!', lt)) { i = text.indexOf('>', lt) + 1; continue; }
    if (text[lt + 1] === '/') {
      const gt = text.indexOf('>', lt);
      cur.closeStart = lt;
      cur.end = gt + 1;
      cur = cur.parent ?? top;
      i = gt + 1;
      continue;
    }
    // start tag: scan to the closing '>' outside quotes
    let j = lt + 1;
    let q = '';
    while (j < n) {
      const ch = text[j];
      if (q) { if (ch === q) q = ''; }
      else if (ch === '"' || ch === "'") q = ch;
      else if (ch === '>') break;
      j++;
    }
    const selfClosing = text[j - 1] === '/';
    const inner = text.slice(lt + 1, selfClosing ? j - 1 : j);
    const m = /^([^\s/>]+)/.exec(inner);
    const name = m ? m[1] : '';
    const attrs: Record<string, string> = {};
    const re = /([^\s=]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
    let a: RegExpExecArray | null;
    const rest = inner.slice(name.length);
    while ((a = re.exec(rest))) attrs[a[1]] = decodeEntities(a[3] ?? a[4] ?? '');
    const node: XNode = {
      name, local: name.includes(':') ? name.slice(name.indexOf(':') + 1) : name, attrs, start: lt, end: j + 1, openEnd: j + 1,
      closeStart: j + 1, selfClosing, children: [], parent: cur,
    };
    cur.children.push(node);
    if (!selfClosing) cur = node;
    i = j + 1;
  }
  const root = top.children[0];
  if (!root) throw new Error('empty XML');
  return { text, root };
}

export function kids(node: XNode | undefined | null, local?: string): XNode[] {
  if (!node) return [];
  return local ? node.children.filter(c => c.local === local) : node.children;
}
export function kid(node: XNode | undefined | null, local: string): XNode | undefined {
  return node?.children.find(c => c.local === local);
}
/** decoded text content of an element (concatenated descendants' text) */
export function textOf(doc: XDocument, node: XNode): string {
  if (node.selfClosing) return '';
  const raw = doc.text.slice(node.openEnd, node.closeStart);
  let s = raw;
  if (s.indexOf('<') >= 0) {
    // strip nested tags but keep CDATA content
    s = s.replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, (_, c: string) => escText(c)).replace(/<[^>]*>/g, '');
  }
  return decodeEntities(s);
}

export function attrNum(node: XNode | undefined, name: string, dflt: number): number {
  const v = node?.attrs[name];
  if (v === undefined || v === '') return dflt;
  const x = Number(v);
  return Number.isFinite(x) ? x : dflt;
}
export function attrBool(node: XNode | undefined, name: string, dflt = false): boolean {
  const v = node?.attrs[name];
  if (v === undefined) return dflt;
  return v === '1' || v === 'true';
}

// ------------------------------------------------------------------ splicing

export interface Splice { start: number; end: number; text: string }

/** apply non-overlapping splices to the source text (order does not matter) */
export function applySplices(text: string, splices: Splice[]): string {
  const sorted = [...splices].sort((a, b) => a.start - b.start || a.end - b.end);
  let out = '';
  let pos = 0;
  for (const s of sorted) {
    if (s.start < pos) throw new Error('overlapping XML edits');
    out += text.slice(pos, s.start) + s.text;
    pos = s.end;
  }
  return out + text.slice(pos);
}

/** a start tag with one attribute set (or removed when value is null), other attributes byte-identical */
export function setAttrSplice(doc: XDocument, node: XNode, name: string, value: string | null): Splice {
  const tagEnd = node.selfClosing ? node.openEnd - 2 : node.openEnd - 1;
  const tag = doc.text.slice(node.start, tagEnd);
  const re = new RegExp(`(\\s)${name.replace(/[:.]/g, '\\$&')}\\s*=\\s*("[^"]*"|'[^']*')`);
  let next: string;
  if (re.test(tag)) next = value === null ? tag.replace(re, '') : tag.replace(re, `$1${name}="${escAttr(value)}"`);
  else next = value === null ? tag : `${tag} ${name}="${escAttr(value)}"`;
  return { start: node.start, end: tagEnd, text: next };
}

/** start tag text with several attributes changed (null removes) */
export function retag(doc: XDocument, node: XNode, changes: Record<string, string | null>): string {
  const tagEnd = node.selfClosing ? node.openEnd - 2 : node.openEnd - 1;
  let tag = doc.text.slice(node.start, tagEnd).replace(/\s+$/, '');
  for (const [name, value] of Object.entries(changes)) {
    const re = new RegExp(`(\\s)${name.replace(/[:.]/g, '\\$&')}\\s*=\\s*("[^"]*"|'[^']*')`);
    if (re.test(tag)) tag = value === null ? tag.replace(re, '') : tag.replace(re, `$1${name}="${escAttr(value)}"`);
    else if (value !== null) tag = `${tag} ${name}="${escAttr(value)}"`;
  }
  return tag;
}
