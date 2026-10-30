// Home screen: new blank workbook, the user's workbooks from the files API, upload.
import { useEffect, useMemo, useRef, useState } from 'react';
import { api, MAX_WORKBOOK_BYTES, baseName, downloadBytes, fmtModified, fmtSize, folderOf, isXlsx, stemOf, type FileEntry } from '../api';
import { BRAND } from '../util/brand';
import { ExcelAppIcon } from '../excel/icons';
import { BookGlyph } from './modals';

function greeting() {
  const h = new Date().getHours();
  return h < 12 ? 'Good morning' : h < 18 ? 'Good afternoon' : 'Good evening';
}

export function Home({ onOpen, onBlank, ask, notify, refreshKey }: {
  onOpen: (path: string) => void;
  onBlank: () => void;
  ask: (text: string, buttons: string[], title?: string) => Promise<string>;
  notify: (text: string, icon?: 'info' | 'error') => Promise<void>;
  refreshKey: number;
}) {
  const [files, setFiles] = useState<FileEntry[] | null>(null);
  const [err, setErr] = useState<string | null>(null);
  const [filter, setFilter] = useState('');
  const [busy, setBusy] = useState<string | null>(null);
  const upload = useRef<HTMLInputElement>(null);

  const load = () => { setErr(null); api.listFiles().then(setFiles, e => { setErr(String(e?.message ?? e)); setFiles([]); }); };
  useEffect(load, [refreshKey]);
  useEffect(() => { document.title = BRAND.excel; }, []);

  const books = useMemo(() => (files ?? []).filter(f => isXlsx(f.path) && f.path.toLowerCase().includes(filter.toLowerCase()))
    .sort((a, b) => b.modifiedMs - a.modifiedMs || a.path.localeCompare(b.path)), [files, filter]);

  const onUpload = async (list: FileList | null) => {
    const file = list?.[0];
    if (upload.current) upload.current.value = '';
    if (!file) return;
    if (!isXlsx(file.name)) { await notify('Only .xlsx workbooks can be uploaded.', 'error'); return; }
    if (file.size > MAX_WORKBOOK_BYTES) { await notify(`${file.name} is too large to upload (the limit is 1.5 MB).`, 'error'); return; }
    let path = file.name;
    if ((files ?? []).some(f => f.path.toLowerCase() === path.toLowerCase())) {
      const b = await ask(`${file.name} already exists in Documents.`, ['Replace', 'Keep Both', 'Cancel'], 'Upload');
      if (b === 'Cancel') return;
      if (b === 'Keep Both') {
        const taken = new Set((files ?? []).map(f => f.path.toLowerCase()));
        let n = 2;
        while (taken.has(`${stemOf(file.name)} (${n}).xlsx`.toLowerCase())) n++;
        path = `${stemOf(file.name)} (${n}).xlsx`;
      }
    }
    setBusy(`Uploading ${file.name}…`);
    try {
      await api.saveFile(path, new Uint8Array(await file.arrayBuffer()), 'upload');
      load();
    } catch (e) {
      await notify(`Couldn't upload ${file.name}. ${e instanceof Error ? e.message : ''}`, 'error');
    } finally { setBusy(null); }
  };

  const download = async (f: FileEntry) => {
    try { downloadBytes(await api.getRaw(f.path), baseName(f.path)); }
    catch (e) { await notify(`Couldn't download ${baseName(f.path)}. ${e instanceof Error ? e.message : ''}`, 'error'); }
  };

  return (
    <div className="home">
      <header className="home-bar">
        <span className="home-logo"><ExcelAppIcon size={24} /></span>
        <span className="home-product">{BRAND.excel}</span>
        <div className="home-search">
          <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden><circle cx="6.8" cy="6.8" r="4.6" fill="none" stroke="currentColor" /><path d="M10.2 10.2 14 14" stroke="currentColor" /></svg>
          <input placeholder="Search workbooks" aria-label="Search workbooks" value={filter} onChange={e => setFilter(e.target.value)} />
        </div>
      </header>
      <main className="home-main">
        <h1>{greeting()}</h1>
        <section className="home-new" aria-label="New">
          <button className="home-tile" onClick={onBlank} aria-label="Blank workbook">
            <span className="home-tile-art" aria-hidden>
              {Array.from({ length: 24 }, (_, i) => <span key={i} />)}
            </span>
            <span className="home-tile-label">Blank workbook</span>
          </button>
        </section>
        <section className="home-files" aria-label="Workbooks">
          <div className="home-files-head">
            <h2>Workbooks</h2>
            <div style={{ flex: 1 }} />
            {busy && <span className="home-busy">{busy}</span>}
            <input ref={upload} type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" hidden onChange={e => void onUpload(e.target.files)} data-testid="upload-input" />
            <button className="home-btn" onClick={() => upload.current?.click()}>
              <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden><path d="M8 11V2.5M4.5 6 8 2.5 11.5 6M2.5 11v2.5h11V11" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" /></svg>
              Upload
            </button>
            <button className="home-btn" onClick={load} aria-label="Refresh">
              <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden><path d="M13 8a5 5 0 1 1-1.5-3.6M12 1.8v3h-3" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" /></svg>
            </button>
          </div>
          <div className="home-table" role="table" aria-label="Workbooks">
            <div className="home-tr home-th" role="row">
              <span role="columnheader">Name</span><span role="columnheader">Folder</span><span role="columnheader">Modified</span><span role="columnheader">Size</span><span />
            </div>
            {err && <div className="home-empty">Couldn't reach the file service ({err}).</div>}
            {!err && files === null && <div className="home-empty">Loading…</div>}
            {!err && files && !books.length && <div className="home-empty">{filter ? 'No workbooks match your search.' : 'No workbooks yet. Start a blank workbook or upload one.'}</div>}
            {books.map(f => (
              <div key={f.path} className="home-tr" role="row" tabIndex={0} data-path={f.path}
                onClick={() => onOpen(f.path)} onKeyDown={e => { if (e.key === 'Enter') onOpen(f.path); }}>
                <span className="home-name" role="cell"><BookGlyph size={20} />{stemOf(baseName(f.path))}</span>
                <span className="home-folder" role="cell">{folderOf(f.path) ? `Documents › ${folderOf(f.path).split('/').join(' › ')}` : 'Documents'}</span>
                <span role="cell">{fmtModified(f.modifiedMs)}</span>
                <span role="cell">{fmtSize(f.size)}</span>
                <span role="cell" className="home-actions">
                  <button className="home-icon-btn" aria-label={`Download ${baseName(f.path)}`} title="Download" onClick={e => { e.stopPropagation(); void download(f); }}>
                    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden><path d="M8 2.5V11M4.5 7.5 8 11l3.5-3.5M2.5 13.5h11" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" /></svg>
                  </button>
                </span>
              </div>
            ))}
          </div>
        </section>
      </main>
    </div>
  );
}
