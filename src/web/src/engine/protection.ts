// Sheet protection (<sheetProtection>): the attribute set Excel writes, the "Allow all users of this
// worksheet to" flags, and the password hash. Excel 2013+ writes SHA-512 with a 16-byte random salt and
// 100,000 spins (ISO/IEC 29500 agile hash: H0 = SHA512(salt + UTF-16LE password), Hn = SHA512(Hn-1 + n as
// 32-bit little-endian)); files from older versions carry the 16-bit legacy `password` hash instead.
// Both are verified; new protection always uses SHA-512 like current Excel.

/** the flags of Excel's Protect Sheet list, in the dialog's order, with the attribute each one maps to */
export const PROTECTION_FLAGS: { attr: string; label: string; key: string; allowedByDefault: boolean }[] = [
  { attr: 'selectLockedCells', label: 'Select locked cells', key: '', allowedByDefault: true },
  { attr: 'selectUnlockedCells', label: 'Select unlocked cells', key: '', allowedByDefault: true },
  { attr: 'formatCells', label: 'Format cells', key: '', allowedByDefault: false },
  { attr: 'formatColumns', label: 'Format columns', key: '', allowedByDefault: false },
  { attr: 'formatRows', label: 'Format rows', key: '', allowedByDefault: false },
  { attr: 'insertColumns', label: 'Insert columns', key: '', allowedByDefault: false },
  { attr: 'insertRows', label: 'Insert rows', key: '', allowedByDefault: false },
  { attr: 'insertHyperlinks', label: 'Insert hyperlinks', key: '', allowedByDefault: false },
  { attr: 'deleteColumns', label: 'Delete columns', key: '', allowedByDefault: false },
  { attr: 'deleteRows', label: 'Delete rows', key: '', allowedByDefault: false },
  { attr: 'sort', label: 'Sort', key: '', allowedByDefault: false },
  { attr: 'autoFilter', label: 'Use AutoFilter', key: '', allowedByDefault: false },
  { attr: 'pivotTables', label: 'Use PivotTable & PivotChart', key: '', allowedByDefault: false },
  { attr: 'objects', label: 'Edit objects', key: '', allowedByDefault: false },
  { attr: 'scenarios', label: 'Edit scenarios', key: '', allowedByDefault: false },
];

/** schema defaults: true means the action is protected (not allowed) when the attribute is absent */
const DEFAULT_PROTECTED: Record<string, boolean> = {
  objects: false, scenarios: false, selectLockedCells: false, selectUnlockedCells: false,
  formatCells: true, formatColumns: true, formatRows: true, insertColumns: true, insertRows: true, insertHyperlinks: true,
  deleteColumns: true, deleteRows: true, sort: true, autoFilter: true, pivotTables: true,
};

/** attribute order of CT_SheetProtection (Excel writes them in this order) */
const ATTR_ORDER = ['password', 'algorithmName', 'hashValue', 'saltValue', 'spinCount', 'sheet', 'objects', 'scenarios', 'formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows', 'insertHyperlinks', 'deleteColumns', 'deleteRows', 'selectLockedCells', 'sort', 'autoFilter', 'pivotTables', 'selectUnlockedCells'];

export type SheetProtection = Record<string, string>;

const isTrue = (v: string | undefined) => v === '1' || v === 'true';

/** is the action allowed for users of a protected sheet? */
export function protectionAllows(p: SheetProtection | null | undefined, attr: string): boolean {
  if (!p) return true;
  const v = p[attr];
  const protectedFlag = v === undefined ? DEFAULT_PROTECTED[attr] ?? false : isTrue(v);
  return !protectedFlag;
}

export function hasPassword(p: SheetProtection | null | undefined): boolean {
  return !!p && (!!p.hashValue || (!!p.password && p.password !== '0000'));
}

/** the attribute set Excel writes for Protect Sheet with these choices */
export function buildProtection(password: string, allow: Record<string, boolean>, saltBytes?: Uint8Array): SheetProtection {
  const out: SheetProtection = {};
  if (password) {
    const salt = saltBytes ?? randomBytes(16);
    out.algorithmName = 'SHA-512';
    out.hashValue = b64(hashPassword(password, salt, 100000));
    out.saltValue = b64(salt);
    out.spinCount = '100000';
  }
  out.sheet = '1';
  for (const f of PROTECTION_FLAGS) {
    const allowed = allow[f.attr] ?? f.allowedByDefault;
    const def = DEFAULT_PROTECTED[f.attr];
    // Excel always writes objects/scenarios when they are protected, and any flag that differs from the default
    if (allowed === !def && f.attr !== 'objects' && f.attr !== 'scenarios') continue;
    if (f.attr === 'objects' || f.attr === 'scenarios') { if (!allowed) out[f.attr] = '1'; continue; }
    out[f.attr] = allowed ? '0' : '1';
  }
  return out;
}

export function protectionXml(p: SheetProtection): string {
  const keys = Object.keys(p).sort((a, b) => (ATTR_ORDER.indexOf(a) + 1 || 99) - (ATTR_ORDER.indexOf(b) + 1 || 99));
  return `<sheetProtection ${keys.map(k => `${k}="${p[k].replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')}"`).join(' ')}/>`;
}

/** does `password` unlock this protection? (no password set: any input, including empty, unlocks) */
export function verifyPassword(p: SheetProtection, password: string): boolean {
  if (p.hashValue && p.saltValue) {
    if ((p.algorithmName ?? 'SHA-512').toUpperCase() !== 'SHA-512') return false;
    const salt = unb64(p.saltValue);
    const spin = Number(p.spinCount ?? '100000') || 0;
    return b64(hashPassword(password, salt, spin)) === p.hashValue;
  }
  if (p.password && p.password !== '0000') return legacyHash(password) === p.password.toUpperCase();
  return true;
}

/** ISO/IEC 29500 password hash for sheet protection (SHA-512, hash then iterator) */
export function hashPassword(password: string, salt: Uint8Array, spinCount: number): Uint8Array {
  const pw = new Uint8Array(password.length * 2);
  for (let i = 0; i < password.length; i++) { const c = password.charCodeAt(i); pw[i * 2] = c & 255; pw[i * 2 + 1] = c >> 8; }
  const first = new Uint8Array(salt.length + pw.length);
  first.set(salt); first.set(pw, salt.length);
  let h = sha512(first);
  const buf = new Uint8Array(68);
  for (let i = 0; i < spinCount; i++) {
    buf.set(h);
    buf[64] = i & 255; buf[65] = (i >>> 8) & 255; buf[66] = (i >>> 16) & 255; buf[67] = (i >>> 24) & 255;
    h = sha512(buf);
  }
  return h;
}

/** the 16-bit hash of Excel 97-2010 (`password="CBEB"`) */
export function legacyHash(password: string): string {
  let h = 0;
  for (let i = password.length - 1; i >= 0; i--) {
    h = ((h >> 14) & 1) | ((h << 1) & 0x7fff);
    h ^= password.charCodeAt(i) & 0xff;
  }
  h = ((h >> 14) & 1) | ((h << 1) & 0x7fff);
  h ^= password.length;
  h ^= 0xce4b;
  return h.toString(16).toUpperCase().padStart(4, '0');
}

function randomBytes(n: number): Uint8Array {
  const out = new Uint8Array(n);
  const c = (globalThis as { crypto?: { getRandomValues?: (a: Uint8Array) => Uint8Array } }).crypto;
  if (c?.getRandomValues) c.getRandomValues(out);
  else for (let i = 0; i < n; i++) out[i] = Math.floor(Math.random() * 256);
  return out;
}

const B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
export function b64(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i < bytes.length; i += 3) {
    const a = bytes[i], b = bytes[i + 1], c = bytes[i + 2];
    s += B64[a >> 2] + B64[((a & 3) << 4) | ((b ?? 0) >> 4)] + (b === undefined ? '=' : B64[((b & 15) << 2) | ((c ?? 0) >> 6)]) + (c === undefined ? '=' : B64[c & 63]);
  }
  return s;
}
export function unb64(s: string): Uint8Array {
  const clean = s.replace(/[^A-Za-z0-9+/]/g, '');
  const out: number[] = [];
  for (let i = 0; i < clean.length; i += 4) {
    const n = [0, 1, 2, 3].map(j => (i + j < clean.length ? B64.indexOf(clean[i + j]) : -1));
    out.push((n[0] << 2) | (n[1] >> 4));
    if (n[2] >= 0) out.push(((n[1] & 15) << 4) | (n[2] >> 2));
    if (n[3] >= 0) out.push(((n[2] & 3) << 6) | n[3]);
  }
  return new Uint8Array(out);
}

// ------------------------------------------------------------------ SHA-512 (32-bit halves; sync, browser + node)

const K = [
  0x428a2f98, 0xd728ae22, 0x71374491, 0x23ef65cd, 0xb5c0fbcf, 0xec4d3b2f, 0xe9b5dba5, 0x8189dbbc, 0x3956c25b, 0xf348b538, 0x59f111f1, 0xb605d019, 0x923f82a4, 0xaf194f9b, 0xab1c5ed5, 0xda6d8118,
  0xd807aa98, 0xa3030242, 0x12835b01, 0x45706fbe, 0x243185be, 0x4ee4b28c, 0x550c7dc3, 0xd5ffb4e2, 0x72be5d74, 0xf27b896f, 0x80deb1fe, 0x3b1696b1, 0x9bdc06a7, 0x25c71235, 0xc19bf174, 0xcf692694,
  0xe49b69c1, 0x9ef14ad2, 0xefbe4786, 0x384f25e3, 0x0fc19dc6, 0x8b8cd5b5, 0x240ca1cc, 0x77ac9c65, 0x2de92c6f, 0x592b0275, 0x4a7484aa, 0x6ea6e483, 0x5cb0a9dc, 0xbd41fbd4, 0x76f988da, 0x831153b5,
  0x983e5152, 0xee66dfab, 0xa831c66d, 0x2db43210, 0xb00327c8, 0x98fb213f, 0xbf597fc7, 0xbeef0ee4, 0xc6e00bf3, 0x3da88fc2, 0xd5a79147, 0x930aa725, 0x06ca6351, 0xe003826f, 0x14292967, 0x0a0e6e70,
  0x27b70a85, 0x46d22ffc, 0x2e1b2138, 0x5c26c926, 0x4d2c6dfc, 0x5ac42aed, 0x53380d13, 0x9d95b3df, 0x650a7354, 0x8baf63de, 0x766a0abb, 0x3c77b2a8, 0x81c2c92e, 0x47edaee6, 0x92722c85, 0x1482353b,
  0xa2bfe8a1, 0x4cf10364, 0xa81a664b, 0xbc423001, 0xc24b8b70, 0xd0f89791, 0xc76c51a3, 0x0654be30, 0xd192e819, 0xd6ef5218, 0xd6990624, 0x5565a910, 0xf40e3585, 0x5771202a, 0x106aa070, 0x32bbd1b8,
  0x19a4c116, 0xb8d2d0c8, 0x1e376c08, 0x5141ab53, 0x2748774c, 0xdf8eeb99, 0x34b0bcb5, 0xe19b48a8, 0x391c0cb3, 0xc5c95a63, 0x4ed8aa4a, 0xe3418acb, 0x5b9cca4f, 0x7763e373, 0x682e6ff3, 0xd6b2b8a3,
  0x748f82ee, 0x5defb2fc, 0x78a5636f, 0x43172f60, 0x84c87814, 0xa1f0ab72, 0x8cc70208, 0x1a6439ec, 0x90befffa, 0x23631e28, 0xa4506ceb, 0xde82bde9, 0xbef9a3f7, 0xb2c67915, 0xc67178f2, 0xe372532b,
  0xca273ece, 0xea26619c, 0xd186b8c7, 0x21c0c207, 0xeada7dd6, 0xcde0eb1e, 0xf57d4f7f, 0xee6ed178, 0x06f067aa, 0x72176fba, 0x0a637dc5, 0xa2c898a6, 0x113f9804, 0xbef90dae, 0x1b710b35, 0x131c471b,
  0x28db77f5, 0x23047d84, 0x32caab7b, 0x40c72493, 0x3c9ebe0a, 0x15c9bebc, 0x431d67c4, 0x9c100d4c, 0x4cc5d4be, 0xcb3e42b6, 0x597f299c, 0xfc657e2a, 0x5fcb6fab, 0x3ad6faec, 0x6c44198c, 0x4a475817,
];
const IV = [0x6a09e667, 0xf3bcc908, 0xbb67ae85, 0x84caa73b, 0x3c6ef372, 0xfe94f82b, 0xa54ff53a, 0x5f1d36f1, 0x510e527f, 0xade682d1, 0x9b05688c, 0x2b3e6c1f, 0x1f83d9ab, 0xfb41bd6b, 0x5be0cd19, 0x137e2179];
const W = new Int32Array(160);

export function sha512(msg: Uint8Array): Uint8Array {
  const len = msg.length;
  const blocks = Math.ceil((len + 17) / 128);
  const data = new Uint8Array(blocks * 128);
  data.set(msg);
  data[len] = 0x80;
  const bits = len * 8;
  const dv = new DataView(data.buffer);
  dv.setUint32(data.length - 8, Math.floor(bits / 0x100000000));
  dv.setUint32(data.length - 4, bits >>> 0);
  const H = Int32Array.from(IV);
  for (let b = 0; b < blocks; b++) {
    for (let i = 0; i < 32; i++) W[i] = dv.getInt32(b * 128 + i * 4);
    for (let i = 32; i < 160; i += 2) {
      // s0 = rotr1 ^ rotr8 ^ shr7 of W[i-30]; s1 = rotr19 ^ rotr61 ^ shr6 of W[i-4]
      let xh = W[i - 30], xl = W[i - 29];
      const s0h = ((xh >>> 1) | (xl << 31)) ^ ((xh >>> 8) | (xl << 24)) ^ (xh >>> 7);
      const s0l = ((xl >>> 1) | (xh << 31)) ^ ((xl >>> 8) | (xh << 24)) ^ ((xl >>> 7) | (xh << 25));
      xh = W[i - 4]; xl = W[i - 3];
      const s1h = ((xh >>> 19) | (xl << 13)) ^ ((xl >>> 29) | (xh << 3)) ^ (xh >>> 6);
      const s1l = ((xl >>> 19) | (xh << 13)) ^ ((xh >>> 29) | (xl << 3)) ^ ((xl >>> 6) | (xh << 26));
      let lo = (s0l >>> 0) + (s1l >>> 0) + (W[i - 13] >>> 0) + (W[i - 31] >>> 0);
      const hi = s0h + s1h + W[i - 14] + W[i - 32] + Math.floor(lo / 0x100000000);
      lo >>>= 0;
      W[i] = hi | 0; W[i + 1] = lo | 0;
    }
    let ah = H[0], al = H[1], bh = H[2], bl = H[3], ch = H[4], cl = H[5], dh = H[6], dl = H[7];
    let eh = H[8], el = H[9], fh = H[10], fl = H[11], gh = H[12], gl = H[13], hh = H[14], hl = H[15];
    for (let i = 0; i < 160; i += 2) {
      const S1h = ((eh >>> 14) | (el << 18)) ^ ((eh >>> 18) | (el << 14)) ^ ((el >>> 9) | (eh << 23));
      const S1l = ((el >>> 14) | (eh << 18)) ^ ((el >>> 18) | (eh << 14)) ^ ((eh >>> 9) | (el << 23));
      const chh = (eh & fh) ^ (~eh & gh), chl = (el & fl) ^ (~el & gl);
      let t1l = (hl >>> 0) + (S1l >>> 0) + (chl >>> 0) + (K[i + 1] >>> 0) + (W[i + 1] >>> 0);
      const t1h = hh + S1h + chh + K[i] + W[i] + Math.floor(t1l / 0x100000000);
      t1l >>>= 0;
      const S0h = ((ah >>> 28) | (al << 4)) ^ ((al >>> 2) | (ah << 30)) ^ ((al >>> 7) | (ah << 25));
      const S0l = ((al >>> 28) | (ah << 4)) ^ ((ah >>> 2) | (al << 30)) ^ ((ah >>> 7) | (al << 25));
      const mh = (ah & bh) ^ (ah & ch) ^ (bh & ch), ml = (al & bl) ^ (al & cl) ^ (bl & cl);
      let t2l = (S0l >>> 0) + (ml >>> 0);
      const t2h = S0h + mh + Math.floor(t2l / 0x100000000);
      t2l >>>= 0;
      hh = gh; hl = gl; gh = fh; gl = fl; fh = eh; fl = el;
      let nl = (dl >>> 0) + t1l;
      eh = (dh + t1h + Math.floor(nl / 0x100000000)) | 0; el = nl | 0;
      dh = ch; dl = cl; ch = bh; cl = bl; bh = ah; bl = al;
      nl = t1l + t2l;
      ah = (t1h + t2h + Math.floor(nl / 0x100000000)) | 0; al = nl | 0;
    }
    const add = (j: number, xh: number, xl: number) => { const l = (H[j + 1] >>> 0) + (xl >>> 0); H[j] = (H[j] + xh + Math.floor(l / 0x100000000)) | 0; H[j + 1] = l | 0; };
    add(0, ah, al); add(2, bh, bl); add(4, ch, cl); add(6, dh, dl); add(8, eh, el); add(10, fh, fl); add(12, gh, gl); add(14, hh, hl);
  }
  const out = new Uint8Array(64);
  const odv = new DataView(out.buffer);
  for (let i = 0; i < 16; i++) odv.setInt32(i * 4, H[i]);
  return out;
}
