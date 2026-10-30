// Client for the workbooks service (same origin). Paths are relative to the user's Documents root with
// forward slashes, e.g. "Capital Calls/Capital Call 4 Notice.xlsx".

export interface FileEntry { path: string; size: number; modifiedMs: number; version: number }
export interface VersionEntry { version: number; modifiedMs: number; size: number }
export interface TaskCheck { sheet: string; cell: string; equals?: number | string; approx?: number; tolerance?: number }
export interface Task { id: string; title: string; instructions: string; file: string; checks?: TaskCheck[] }

export type PutKind = 'save' | 'save_as' | 'upload';

async function ok(res: Response): Promise<Response> {
  if (!res.ok) {
    let msg = `${res.status} ${res.statusText}`;
    try { const t = await res.text(); if (t) msg = t; } catch { /* keep status */ }
    throw new Error(msg);
  }
  return res;
}

const q = (path: string) => `path=${encodeURIComponent(path)}`;

export const api = {
  async listFiles(): Promise<FileEntry[]> {
    const res = await ok(await fetch('/api/files', { cache: 'no-store' }));
    const body = await res.json() as { files?: FileEntry[] };
    return Array.isArray(body.files) ? body.files : [];
  },
  rawUrl(path: string) { return `/api/files/raw?${q(path)}`; },
  async getRaw(path: string): Promise<Uint8Array> {
    const res = await ok(await fetch(`/api/files/raw?${q(path)}`, { cache: 'no-store' }));
    return new Uint8Array(await res.arrayBuffer());
  },
  /** save / Save As / upload: POST /api/files/save with the bytes as base64 (the service limits a request to ~2 MB) */
  async saveFile(path: string, bytes: Uint8Array, kind: PutKind): Promise<FileEntry> {
    const res = await ok(await fetch('/api/files/save', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ path, kind, contentBase64: toBase64(bytes) }),
    }));
    return await res.json() as FileEntry;
  },
  async deleteRaw(path: string): Promise<void> {
    await ok(await fetch(`/api/files/raw?${q(path)}`, { method: 'DELETE' }));
  },
  async versions(path: string): Promise<VersionEntry[]> {
    const res = await ok(await fetch(`/api/files/versions?${q(path)}`, { cache: 'no-store' }));
    const body = await res.json() as { versions?: VersionEntry[] };
    return Array.isArray(body.versions) ? body.versions : [];
  },
  async tasks(): Promise<Task[]> {
    const res = await ok(await fetch('/api/env/tasks', { cache: 'no-store' }));
    const body = await res.json() as { tasks?: Task[] };
    return Array.isArray(body.tasks) ? body.tasks : [];
  },
  async reset(scenario = 'fund-admin'): Promise<void> {
    await ok(await fetch('/api/env/reset', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ scenario }) }));
  },
};

/** largest workbook the service accepts (2 MB request limit after base64) */
export const MAX_WORKBOOK_BYTES = 1_500_000;

export function toBase64(bytes: Uint8Array): string {
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}

export const isXlsx = (p: string) => /\.xlsx$/i.test(p);
export const baseName = (p: string) => p.slice(p.lastIndexOf('/') + 1);
export const folderOf = (p: string) => { const i = p.lastIndexOf('/'); return i < 0 ? '' : p.slice(0, i); };
export const stemOf = (n: string) => { const i = n.lastIndexOf('.'); return i > 0 ? n.slice(0, i) : n; };
export const joinPath = (dir: string, name: string) => (dir ? `${dir.replace(/\/+$/, '')}/${name}` : name);

export function fmtModified(ms: number): string {
  if (!ms) return '';
  const d = new Date(ms), now = new Date();
  const sameDay = d.toDateString() === now.toDateString();
  if (sameDay) return d.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
  const y = new Date(now); y.setDate(now.getDate() - 1);
  if (d.toDateString() === y.toDateString()) return 'Yesterday';
  return d.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: d.getFullYear() === now.getFullYear() ? undefined : 'numeric' });
}

export function fmtSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

export function downloadBytes(bytes: Uint8Array, name: string) {
  const blob = new Blob([bytes as unknown as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
