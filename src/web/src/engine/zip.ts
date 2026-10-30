// Lossless ZIP container for OOXML packages.
//
// Every entry keeps its original local record (header + compressed bytes + data descriptor) and its
// original central-directory record. Saving copies those bytes verbatim and only rewrites entries
// that were replaced, so a package saved without edits is byte-identical to the input, and a package
// with one edited part differs only in that part's record and the shifted offsets.

import { inflateSync, deflateSync } from 'fflate';

const enc = new TextEncoder();
const dec = new TextDecoder('utf-8');

interface Entry {
  name: string;
  /** raw local record bytes (header, name, extra, data, descriptor) or null for new/replaced entries */
  local: Uint8Array | null;
  /** raw central-directory record or null */
  central: Uint8Array | null;
  method: number;
  compSize: number;
  dosTime: number;
  dosDate: number;
  flags: number;
  extAttr: number;
  madeBy: number;
  /** decompressed content, filled lazily */
  data: Uint8Array | null;
  /** compressed bytes as stored in the original */
  stored: Uint8Array | null;
  changed: boolean;
}

function u16(b: Uint8Array, o: number) { return b[o] | (b[o + 1] << 8); }
function u32(b: Uint8Array, o: number) { return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0; }
function w16(b: Uint8Array, o: number, v: number) { b[o] = v & 255; b[o + 1] = (v >>> 8) & 255; }
function w32(b: Uint8Array, o: number, v: number) { b[o] = v & 255; b[o + 1] = (v >>> 8) & 255; b[o + 2] = (v >>> 16) & 255; b[o + 3] = (v >>> 24) & 255; }

let crcTable: Uint32Array | null = null;
export function crc32(data: Uint8Array): number {
  if (!crcTable) {
    crcTable = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTable[n] = c >>> 0;
    }
  }
  let c = 0xffffffff;
  for (let i = 0; i < data.length; i++) c = crcTable[(c ^ data[i]) & 255] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

export class ZipPackage {
  entries: Entry[] = [];
  private eocd: Uint8Array = new Uint8Array(0);
  private original: Uint8Array | null = null;

  static open(bytes: Uint8Array): ZipPackage {
    const z = new ZipPackage();
    z.original = bytes;
    let e = -1;
    for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 22 - 65535); i--) {
      if (u32(bytes, i) === 0x06054b50) { e = i; break; }
    }
    if (e < 0) throw new Error('not a zip file');
    const count = u16(bytes, e + 10);
    const cdSize = u32(bytes, e + 12);
    const cdOffset = u32(bytes, e + 16);
    if (count === 0xffff || cdOffset === 0xffffffff) throw new Error('zip64 is not supported');
    z.eocd = bytes.slice(e);
    // central directory
    const cds: { name: string; off: number; rec: Uint8Array; method: number; compSize: number; time: number; date: number; flags: number; ext: number; madeBy: number }[] = [];
    let p = cdOffset;
    for (let i = 0; i < count; i++) {
      if (u32(bytes, p) !== 0x02014b50) throw new Error('bad central directory');
      const nameLen = u16(bytes, p + 28), extraLen = u16(bytes, p + 30), commentLen = u16(bytes, p + 32);
      const len = 46 + nameLen + extraLen + commentLen;
      cds.push({
        name: dec.decode(bytes.subarray(p + 46, p + 46 + nameLen)), off: u32(bytes, p + 42), rec: bytes.slice(p, p + len),
        method: u16(bytes, p + 10), compSize: u32(bytes, p + 20), time: u16(bytes, p + 12), date: u16(bytes, p + 14),
        flags: u16(bytes, p + 8), ext: u32(bytes, p + 38), madeBy: u16(bytes, p + 4),
      });
      p += len;
    }
    if (p !== cdOffset + cdSize) throw new Error('central directory size mismatch');
    // local records span from their offset to the next record (or the central directory)
    const offsets = cds.map(c => c.off).sort((a, b) => a - b);
    for (const c of cds) {
      const next = offsets.find(o => o > c.off) ?? cdOffset;
      const local = bytes.slice(c.off, next);
      if (u32(local, 0) !== 0x04034b50) throw new Error('bad local header');
      const dataStart = 30 + u16(local, 26) + u16(local, 28);
      z.entries.push({
        name: c.name, local, central: c.rec, method: c.method, compSize: c.compSize, dosTime: c.time, dosDate: c.date,
        flags: c.flags, extAttr: c.ext, madeBy: c.madeBy, data: null, stored: local.subarray(dataStart, dataStart + c.compSize), changed: false,
      });
    }
    // keep the physical order (what a rewrite must reproduce)
    const pos = new Map(cds.map(c => [c.name, c.off]));
    z.entries.sort((a, b) => pos.get(a.name)! - pos.get(b.name)!);
    // central order may differ from physical order: remember it
    z.centralOrder = cds.map(c => c.name);
    return z;
  }

  centralOrder: string[] = [];

  names(): string[] { return this.entries.map(e => e.name); }
  has(name: string): boolean { return this.entries.some(e => e.name === name); }

  read(name: string): Uint8Array {
    const e = this.entries.find(x => x.name === name);
    if (!e) throw new Error(`no entry ${name}`);
    if (!e.data) {
      if (e.method === 0) e.data = e.stored!;
      else if (e.method === 8) e.data = inflateSync(e.stored!);
      else throw new Error(`unsupported compression ${e.method} in ${name}`);
    }
    return e.data;
  }

  readText(name: string): string {
    return new TextDecoder('utf-8', { ignoreBOM: true }).decode(this.read(name));
  }

  /** replace (or add) an entry's content */
  write(name: string, data: Uint8Array | string) {
    const bytes = typeof data === 'string' ? enc.encode(data) : data;
    let e = this.entries.find(x => x.name === name);
    if (e) {
      // writing the same bytes back is not a change
      if (!e.changed && e.data && sameBytes(e.data, bytes)) return;
      if (!e.changed && !e.data) { const cur = this.read(name); if (sameBytes(cur, bytes)) return; }
      e.data = bytes; e.changed = true; e.local = null; e.central = null;
    } else {
      e = { name, local: null, central: null, method: 8, compSize: 0, dosTime: 0, dosDate: 0x5d3e, flags: 0, extAttr: 0, madeBy: 20, data: bytes, stored: null, changed: true };
      // new parts take the date stamp of the package's first entry
      const first = this.entries[0];
      if (first) { e.dosTime = first.dosTime; e.dosDate = first.dosDate; }
      this.entries.push(e);
      this.centralOrder.push(name);
    }
  }

  remove(name: string) {
    this.entries = this.entries.filter(e => e.name !== name);
    this.centralOrder = this.centralOrder.filter(n => n !== name);
  }

  changedNames(): string[] { return this.entries.filter(e => e.changed).map(e => e.name); }

  save(): Uint8Array {
    // no shortcut for the unchanged case: the generic path reproduces the original bytes exactly,
    // which is what the round-trip test checks
    const chunks: Uint8Array[] = [];
    let offset = 0;
    const centralFor = new Map<string, Uint8Array>();
    for (const e of this.entries) {
      let local = e.local;
      let central = e.central;
      if (!local || !central) {
        const nameBytes = enc.encode(e.name);
        const data = e.data ?? new Uint8Array(0);
        const comp = deflateSync(data, { level: 6 });
        const crc = crc32(data);
        const utf8 = /[^\x00-\x7f]/.test(e.name) ? 0x800 : 0;
        local = new Uint8Array(30 + nameBytes.length + comp.length);
        w32(local, 0, 0x04034b50); w16(local, 4, 20); w16(local, 6, utf8); w16(local, 8, 8);
        w16(local, 10, e.dosTime); w16(local, 12, e.dosDate); w32(local, 14, crc);
        w32(local, 18, comp.length); w32(local, 22, data.length); w16(local, 26, nameBytes.length); w16(local, 28, 0);
        local.set(nameBytes, 30); local.set(comp, 30 + nameBytes.length);
        central = new Uint8Array(46 + nameBytes.length);
        w32(central, 0, 0x02014b50); w16(central, 4, e.madeBy || 20); w16(central, 6, 20); w16(central, 8, utf8); w16(central, 10, 8);
        w16(central, 12, e.dosTime); w16(central, 14, e.dosDate); w32(central, 16, crc); w32(central, 20, comp.length);
        w32(central, 24, data.length); w16(central, 28, nameBytes.length); w16(central, 30, 0); w16(central, 32, 0);
        w16(central, 34, 0); w16(central, 36, 0); w32(central, 38, e.extAttr); w32(central, 42, 0);
        central.set(nameBytes, 46);
      } else {
        central = central.slice();
      }
      w32(central, 42, offset);
      centralFor.set(e.name, central);
      chunks.push(local);
      offset += local.length;
    }
    const cdStart = offset;
    for (const n of this.centralOrder) {
      const c = centralFor.get(n);
      if (c) { chunks.push(c); offset += c.length; }
    }
    const eocd = this.eocd.length ? this.eocd.slice() : new Uint8Array(22);
    if (!this.eocd.length) w32(eocd, 0, 0x06054b50);
    w16(eocd, 8, this.entries.length); w16(eocd, 10, this.entries.length);
    w32(eocd, 12, offset - cdStart); w32(eocd, 16, cdStart);
    chunks.push(eocd);
    const out = new Uint8Array(offset + eocd.length);
    let p = 0;
    for (const c of chunks) { out.set(c, p); p += c.length; }
    return out;
  }
}

export function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}
