// Excel number formats: built-in ids and a renderer for the format codes business workbooks use
// (General, fixed decimals, thousands separators, percent, currency/accounting, dates and times,
// text, sections with colours and conditions, literal text, scientific).

export const BUILTIN_FORMATS: Record<number, string> = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
  44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
  5: '"$"#,##0_);\\("$"#,##0\\)', 6: '"$"#,##0_);[Red]\\("$"#,##0\\)', 7: '"$"#,##0.00_);\\("$"#,##0.00\\)', 8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)', 42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)', 43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
};

export interface Formatted { text: string; color?: string; align?: 'left' | 'right'; /** accounting: text split around the fill char */ fill?: { left: string; right: string } }

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const COLORS: Record<string, string> = { red: '#ff0000', blue: '#0000ff', green: '#00ff00', black: '#000000', white: '#ffffff', magenta: '#ff00ff', yellow: '#ffff00', cyan: '#00ffff' };

/** serial date -> parts (1900 system with the Lotus leap-year bug, or 1904) */
export function serialToParts(serial: number, date1904 = false) {
  let days = Math.floor(serial);
  const frac = serial - days;
  if (!date1904) { if (days > 59) days -= 1; days -= 1; } // serial 1 = 1900-01-01; 60 = fictitious 1900-02-29
  const base = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1900, 0, 1);
  const d = new Date(base + days * 86400000);
  let secs = Math.round(frac * 86400);
  if (secs >= 86400) secs = 86399;
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), dow: d.getUTCDay(), H: Math.floor(secs / 3600), M: Math.floor((secs % 3600) / 60), S: secs % 60, frac };
}
export function partsToSerial(y: number, m: number, d: number, date1904 = false): number {
  const ms = Date.UTC(y, m - 1, d);
  if (date1904) return Math.round((ms - Date.UTC(1904, 0, 1)) / 86400000);
  let days = Math.round((ms - Date.UTC(1900, 0, 1)) / 86400000) + 1;
  if (days >= 60) days += 1;
  return days;
}

export function isDateFormat(code: string): boolean {
  const c = code.replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/\[[^\]]*\]/g, '');
  return /[dmyhs]/i.test(c) && !/^General$/i.test(code);
}

interface Section { code: string; color?: string; cond?: { op: string; v: number } }

function splitSections(code: string): Section[] {
  const out: string[] = [];
  let cur = '', q = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') q = !q;
    if (ch === '\\' && !q) { cur += ch + (code[i + 1] ?? ''); i++; continue; }
    if (ch === ';' && !q) { out.push(cur); cur = ''; continue; }
    cur += ch;
  }
  out.push(cur);
  return out.map(s => {
    const sec: Section = { code: s };
    sec.code = s.replace(/\[([^\]]*)\]/g, (m, inner: string) => {
      const lc = inner.toLowerCase();
      if (COLORS[lc]) { sec.color = COLORS[lc]; return ''; }
      if (/^color\d+$/.test(lc)) return '';
      const cm = /^(<=|>=|<>|<|>|=)(-?[\d.]+)$/.exec(inner);
      if (cm) { sec.cond = { op: cm[1], v: Number(cm[2]) }; return ''; }
      if (/^\$/.test(inner)) { const sym = inner.slice(1).split('-')[0]; return `"${sym}"`; }
      return m; // [h] [mm] [ss] elapsed time
    });
    return sec;
  });
}

function testCond(c: { op: string; v: number }, x: number) {
  switch (c.op) { case '<': return x < c.v; case '<=': return x <= c.v; case '>': return x > c.v; case '>=': return x >= c.v; case '=': return x === c.v; case '<>': return x !== c.v; }
  return false;
}

export function formatGeneral(v: number): string {
  if (v === 0) return '0';
  if (!Number.isFinite(v)) return '#NUM!';
  const abs = Math.abs(v);
  if (abs >= 1e11 || abs < 1e-9) {
    let s = v.toExponential(5).replace(/\.?0+e/, 'E').replace('e', 'E');
    s = s.replace(/E([+-])(\d)$/, 'E$10$2');
    return s;
  }
  // up to 11 characters, like a standard-width column
  const intDigits = Math.floor(Math.log10(abs)) + 1;
  const decimals = Math.max(0, Math.min(10 - Math.max(intDigits, 1) - (v < 0 ? 1 : 0), 15));
  let s = v.toFixed(decimals);
  if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
  return s;
}

/** format a cell value with a format code */
export function formatValue(v: unknown, code: string, date1904 = false): Formatted {
  if (v === null || v === undefined || v === '') return { text: '' };
  if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE', align: undefined };
  if (typeof v === 'object' && v && 'error' in (v as object)) return { text: (v as { error: string }).error };
  const sections = splitSections(code || 'General');
  if (typeof v === 'string') {
    const textSec = sections.length >= 4 ? sections[3] : sections.find(s => s.code.includes('@'));
    if (!textSec) return { text: v };
    return { text: renderLiteral(textSec.code, v), color: textSec.color };
  }
  if (typeof v !== 'number') return { text: String(v) };
  let sec = sections[0];
  let x = v;
  let useAbs = false;
  const hasConds = sections.some(s => s.cond);
  if (hasConds) {
    const hit = sections.find(s => s.cond && testCond(s.cond, v));
    sec = hit ?? sections.find(s => !s.cond && !s.code.includes('@')) ?? sections[sections.length - 1];
    useAbs = !!hit && v < 0;
  } else if (sections.length >= 2 && v < 0 && !sections[1].code.includes('@')) { sec = sections[1]; useAbs = true; }
  else if (sections.length >= 3 && v === 0 && !sections[2].code.includes('@')) sec = sections[2];
  if (useAbs) x = Math.abs(v);
  const c = sec.code;
  if (/^\s*General\s*$/i.test(c) || c.trim() === '') {
    const t = formatGeneral(x);
    return { text: (v < 0 && useAbs ? '-' : '') + (useAbs ? t.replace(/^-/, '') : t), color: sec.color };
  }
  if (/^\s*General\s*$/i.test(c.replace(/"[^"]*"/g, ''))) return { text: renderLiteral(c.replace(/General/i, formatGeneral(x)), ''), color: sec.color };
  if (isDateFormat(c)) return { text: formatDate(x, c, date1904), color: sec.color };
  const r = formatNumber(x, c);
  return { ...r, color: sec.color };
}

function renderLiteral(code: string, value: string): string {
  let out = '';
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') { const j = code.indexOf('"', i + 1); out += code.slice(i + 1, j < 0 ? undefined : j); i = j < 0 ? code.length : j; }
    else if (ch === '\\') { out += code[i + 1] ?? ''; i++; }
    else if (ch === '@') out += value;
    else if (ch === '_') { out += ' '; i++; }
    else if (ch === '*') { i++; }
    else out += ch;
  }
  return out;
}

function formatNumber(x: number, code: string): Formatted {
  // tokens: literals, digit placeholders, decimal point, thousands, percent, exponent, fill (*x), padding (_x)
  let pct = 0;
  let q = false;
  let fillAt = -1;
  const lits: string[] = [];
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') q = !q;
    else if (!q && ch === '\\') i++;
    else if (!q && ch === '%') pct++;
  }
  let v = x * Math.pow(100, pct);
  // find the numeric part: from first placeholder to last placeholder
  const isPh = (ch: string) => ch === '0' || ch === '#' || ch === '?';
  const sciM = /[eE][+-]/.exec(code.replace(/"[^"]*"/g, m => ' '.repeat(m.length)));
  let intPh = '', decPh = '', seenDot = false, thousands = false, scaleCommas = 0;
  let prefix = '', suffix = '', expPh = '';
  let state: 'pre' | 'num' | 'post' = 'pre';
  q = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    const target = () => (state === 'pre' ? 'pre' : 'post');
    if (ch === '"') {
      const j = code.indexOf('"', i + 1);
      const lit = code.slice(i + 1, j < 0 ? undefined : j);
      if (target() === 'pre' && state === 'pre') prefix += lit; else suffix += lit;
      if (state === 'num') state = 'post';
      i = j < 0 ? code.length : j; continue;
    }
    if (ch === '\\') { const lit = code[i + 1] ?? ''; if (state === 'pre') prefix += lit; else { suffix += lit; state = 'post'; } i++; continue; }
    if (ch === '_') { if (state === 'pre') prefix += ' '; else { suffix += ' '; if (state === 'num') state = 'post'; } i++; continue; }
    if (ch === '*') { fillAt = state === 'pre' ? prefix.length : -2; lits.push(prefix); i++; continue; }
    if (sciM && i === sciM.index) {
      i += 1;
      while (i + 1 < code.length && isPh(code[i + 1])) { expPh += code[i + 1]; i++; }
      state = 'post';
      continue;
    }
    if (isPh(ch) && state !== 'post') { state = 'num'; if (seenDot) decPh += ch; else intPh += ch; continue; }
    if (ch === '.' && state !== 'post') { state = 'num'; seenDot = true; continue; }
    if (ch === ',' && state === 'num') {
      // comma between digit placeholders = thousands; trailing commas scale by 1000
      const restPh = /[0#?]/.test(code.slice(i + 1).replace(/"[^"]*"/g, '').split(/[.;]/)[0] ?? '');
      if (restPh && !seenDot) thousands = true; else scaleCommas++;
      continue;
    }
    if (ch === '%') { if (state === 'pre') prefix += '%'; else { suffix += '%'; state = 'post'; } continue; }
    if (ch === '@') continue;
    if (state === 'pre') prefix += ch; else { suffix += ch; if (state === 'num') state = 'post'; }
  }
  v = v / Math.pow(1000, scaleCommas);
  const neg = v < 0;
  v = Math.abs(v);
  let body: string;
  if (sciM) {
    const dec = decPh.length;
    let exp = v === 0 ? 0 : Math.floor(Math.log10(v));
    let mant = v / Math.pow(10, exp);
    if (Number(mant.toFixed(dec)) >= 10) { mant /= 10; exp += 1; }
    const sign = exp < 0 ? '-' : (code[sciM.index + 1] === '+' ? '+' : '');
    body = mant.toFixed(dec) + 'E' + sign + String(Math.abs(exp)).padStart(Math.max(expPh.length, 1), '0');
  } else {
    const dec = decPh.length;
    const s = roundFixed(v, dec);
    let [ip, dp = ''] = s.split('.');
    // integer placeholders: zeros force digits
    const minInt = (intPh.match(/0/g) ?? []).length;
    if (ip === '0' && minInt === 0) ip = '';
    while (ip.length < minInt) ip = '0' + ip;
    if (thousands) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    // decimal placeholders: '#' drops trailing zeros, '?' pads with spaces
    if (dec) {
      const chars = dp.split('');
      for (let k = dec - 1; k >= 0; k--) {
        if (chars[k] === '0' && decPh[k] === '#') chars.pop();
        else if (chars[k] === '0' && decPh[k] === '?') chars[k] = ' ';
        else break;
      }
      dp = chars.join('');
    }
    body = ip + (seenDot ? '.' + dp : '');
    if (intPh.includes('?') && ip === '') body = ' ' + body;
  }
  const minus = neg && Number(body.replace(/[^\d]/g, '')) !== 0 ? '-' : '';
  if (fillAt >= 0) {
    return { text: minus + prefix + body + suffix, fill: { left: prefix.slice(0, fillAt), right: prefix.slice(fillAt) + minus + body + suffix } };
  }
  return { text: minus + prefix + body + suffix };
}

function roundFixed(v: number, dec: number): string {
  // avoid binary artefacts like 1.005 -> 1.00
  const f = Math.pow(10, dec);
  const r = Math.round((v + Number.EPSILON * Math.max(1, v)) * f) / f;
  return r.toFixed(dec);
}

function formatDate(serial: number, code: string, date1904: boolean): string {
  const p = serialToParts(serial, date1904);
  const ampm = /AM\/PM|A\/P/i.test(code);
  let out = '';
  const lc = code;
  // tokenise
  const toks: { t: string; v: string }[] = [];
  for (let i = 0; i < lc.length;) {
    const ch = lc[i];
    if (ch === '"') { const j = lc.indexOf('"', i + 1); toks.push({ t: 'lit', v: lc.slice(i + 1, j < 0 ? undefined : j) }); i = j < 0 ? lc.length : j + 1; continue; }
    if (ch === '\\') { toks.push({ t: 'lit', v: lc[i + 1] ?? '' }); i += 2; continue; }
    if (/^AM\/PM/i.test(lc.slice(i))) { toks.push({ t: 'ampm', v: 'AM/PM' }); i += 5; continue; }
    if (/^A\/P/i.test(lc.slice(i))) { toks.push({ t: 'ap', v: 'A/P' }); i += 3; continue; }
    if (ch === '[') { const j = lc.indexOf(']', i); toks.push({ t: 'elapsed', v: lc.slice(i + 1, j).toLowerCase() }); i = j + 1; continue; }
    const m = /^(y+|m+|d+|h+|s+)/i.exec(lc.slice(i));
    if (m) { toks.push({ t: m[1][0].toLowerCase(), v: m[1].toLowerCase() }); i += m[1].length; continue; }
    if (ch === '.' && /^\.0+/.test(lc.slice(i))) { const z = /^\.0+/.exec(lc.slice(i))![0]; toks.push({ t: 'frac', v: z }); i += z.length; continue; }
    if (ch === '_' || ch === '*') { i += 2; continue; }
    toks.push({ t: 'lit', v: ch }); i++;
  }
  // m after h or before s means minutes
  for (let k = 0; k < toks.length; k++) {
    if (toks[k].t !== 'm') continue;
    const prev = toks.slice(0, k).reverse().find(t => t.t !== 'lit');
    const next = toks.slice(k + 1).find(t => t.t !== 'lit');
    if ((prev && (prev.t === 'h' || prev.t === 'elapsed')) || (next && next.t === 's')) toks[k].t = 'min';
  }
  for (const t of toks) {
    switch (t.t) {
      case 'lit': out += t.v; break;
      case 'y': out += t.v.length <= 2 ? String(p.y % 100).padStart(2, '0') : String(p.y); break;
      case 'm': out += t.v.length === 1 ? String(p.m) : t.v.length === 2 ? String(p.m).padStart(2, '0') : t.v.length === 3 ? MONTHS[p.m - 1].slice(0, 3) : t.v.length === 5 ? MONTHS[p.m - 1][0] : MONTHS[p.m - 1]; break;
      case 'd': out += t.v.length === 1 ? String(p.d) : t.v.length === 2 ? String(p.d).padStart(2, '0') : t.v.length === 3 ? DAYS[p.dow].slice(0, 3) : DAYS[p.dow]; break;
      case 'h': { const h = ampm ? ((p.H + 11) % 12) + 1 : p.H; out += t.v.length >= 2 ? String(h).padStart(2, '0') : String(h); break; }
      case 'min': out += t.v.length >= 2 ? String(p.M).padStart(2, '0') : String(p.M); break;
      case 's': out += t.v.length >= 2 ? String(p.S).padStart(2, '0') : String(p.S); break;
      case 'ampm': out += p.H < 12 ? 'AM' : 'PM'; break;
      case 'ap': out += p.H < 12 ? 'A' : 'P'; break;
      case 'frac': out += (p.frac * 86400 % 1).toFixed(t.v.length - 1).slice(1); break;
      case 'elapsed': {
        const total = serial * 24;
        if (t.v[0] === 'h') out += String(Math.floor(total));
        else if (t.v[0] === 'm') out += String(Math.floor(total * 60));
        else out += String(Math.floor(total * 3600));
        break;
      }
    }
  }
  return out;
}
