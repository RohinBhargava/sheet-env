// In-page modals the page host shows for the workbook view: message boxes, the Open picker and Save As.
import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { api, baseName, folderOf, fmtModified, isXlsx, joinPath, stemOf, type FileEntry } from '../api';
import { BRAND } from '../util/brand';

export type MessageIcon = 'info' | 'warning' | 'error' | 'question' | 'none';

function Frame({ title, width, onClose, children, label }: { title: string; width: number; onClose: () => void; children: ReactNode; label?: string }) {
  return (
    <div className="hm-backdrop" onMouseDown={e => e.stopPropagation()}>
      <div className="hm-dialog" role="dialog" aria-modal="true" aria-label={label ?? title} style={{ width }}
        onKeyDown={e => { e.stopPropagation(); if (e.key === 'Escape') { e.preventDefault(); onClose(); } }}>
        <div className="hm-title">
          <span>{title}</span>
          <button className="hm-x" aria-label="Close" onClick={onClose}>
            <svg width="10" height="10" viewBox="0 0 10 10" aria-hidden><path d="M1 1l8 8M9 1L1 9" stroke="currentColor" strokeWidth="1.1" /></svg>
          </button>
        </div>
        {children}
      </div>
    </div>
  );
}

function MsgIcon({ icon }: { icon: MessageIcon }) {
  if (icon === 'none') return null;
  const color = icon === 'error' ? '#c42b1c' : icon === 'warning' ? '#d18f00' : '#1a7f4b';
  const glyph = icon === 'error' ? 'M11 11l10 10M21 11L11 21' : icon === 'question' ? 'M12.5 12.5a3.5 3.5 0 1 1 5 3.2c-1 .5-1.5 1.2-1.5 2.3v.5M16 22v.5' : icon === 'warning' ? 'M16 9v9M16 22v.5' : 'M16 14v9M16 9.5v.5';
  return (
    <svg className="hm-msg-icon" width="32" height="32" viewBox="0 0 32 32" aria-hidden>
      {icon === 'warning' ? <path d="M16 3 30 28H2z" fill={color} /> : <circle cx="16" cy="16" r="14" fill={color} />}
      <path d={glyph} stroke="#fff" strokeWidth="2.4" strokeLinecap="round" fill="none" />
    </svg>
  );
}

export function MessageBox({ title, text, icon = 'info', buttons = ['OK'], onDone }: { title: string; text: string; icon?: MessageIcon; buttons?: string[]; onDone: (b: string) => void }) {
  const first = useRef<HTMLButtonElement>(null);
  useEffect(() => { first.current?.focus(); }, []);
  const cancel = buttons.includes('Cancel') ? 'Cancel' : buttons[buttons.length - 1];
  return (
    <Frame title={title} width={Math.min(520, Math.max(380, text.length * 2.2))} onClose={() => onDone(cancel)} label={`${title}: ${text.split('\n')[0]}`}>
      <div className="hm-msg">
        <MsgIcon icon={icon} />
        <div className="hm-msg-text">{text.split('\n').map((l, i) => <p key={i}>{l || ' '}</p>)}</div>
      </div>
      <div className="hm-btns">
        {buttons.map((b, i) => <button key={b} ref={i === 0 ? first : undefined} className={`hm-btn${i === 0 ? ' primary' : ''}`} onClick={() => onDone(b)}>{b}</button>)}
      </div>
    </Frame>
  );
}

function useFiles() {
  const [files, setFiles] = useState<FileEntry[] | null>(null);
  const [err, setErr] = useState<string | null>(null);
  useEffect(() => { api.listFiles().then(setFiles, e => setErr(String(e?.message ?? e))); }, []);
  return { files, err };
}

export function OpenPicker({ title = 'Open', okLabel = 'Open', onDone }: { title?: string; okLabel?: string; onDone: (path: string | null) => void }) {
  const { files, err } = useFiles();
  const [sel, setSel] = useState<string | null>(null);
  const [filter, setFilter] = useState('');
  const list = useMemo(() => (files ?? []).filter(f => isXlsx(f.path) && f.path.toLowerCase().includes(filter.toLowerCase()))
    .sort((a, b) => a.path.localeCompare(b.path, undefined, { numeric: true })), [files, filter]);
  return (
    <Frame title={title} width={640} onClose={() => onDone(null)}>
      <div className="hm-body">
        <input className="hm-input" autoFocus placeholder="Filter workbooks" aria-label="Filter workbooks" value={filter} onChange={e => setFilter(e.target.value)}
          onKeyDown={e => { if (e.key === 'Enter' && (sel || list.length === 1)) onDone(sel ?? list[0].path); }} />
        <div className="hm-list" role="listbox" aria-label="Workbooks">
          {err && <div className="hm-empty">Couldn't load the file list: {err}</div>}
          {!err && !files && <div className="hm-empty">Loading…</div>}
          {files && !list.length && <div className="hm-empty">No workbooks found.</div>}
          {list.map(f => (
            <div key={f.path} role="option" aria-selected={sel === f.path} className={`hm-row${sel === f.path ? ' sel' : ''}`}
              onClick={() => setSel(f.path)} onDoubleClick={() => onDone(f.path)}>
              <span className="hm-row-name"><BookGlyph />{stemOf(baseName(f.path))}</span>
              <span className="hm-row-folder">{folderOf(f.path) || 'Documents'}</span>
              <span className="hm-row-date">{fmtModified(f.modifiedMs)}</span>
            </div>
          ))}
        </div>
      </div>
      <div className="hm-btns">
        <button className="hm-btn primary" disabled={!sel} onClick={() => sel && onDone(sel)}>{okLabel}</button>
        <button className="hm-btn" onClick={() => onDone(null)}>Cancel</button>
      </div>
    </Frame>
  );
}

const BAD_NAME = /[\\/:*?"<>|]/;

export function SaveAsDialog({ title = 'Save As', defaultName, defaultFolder = '', onDone, confirm }: { title?: string; defaultName: string; defaultFolder?: string; onDone: (path: string | null) => void; confirm: (text: string) => Promise<boolean> }) {
  const { files } = useFiles();
  const [folder, setFolder] = useState(defaultFolder);
  const [name, setName] = useState(stemOf(defaultName));
  const [problem, setProblem] = useState<string | null>(null);
  const nameRef = useRef<HTMLInputElement>(null);
  useEffect(() => { nameRef.current?.focus(); nameRef.current?.select(); }, []);
  const folders = useMemo(() => [...new Set((files ?? []).map(f => folderOf(f.path)).filter(Boolean))].sort(), [files]);
  const submit = async () => {
    const n = name.trim(), dir = folder.trim().replace(/^\/+|\/+$/g, '');
    if (!n) { setProblem('Enter a file name.'); return; }
    if (BAD_NAME.test(n)) { setProblem('A file name can\'t contain any of these characters: \\ / : * ? " < > |'); return; }
    if (dir.split('/').some(seg => seg === '..' || /[\\:*?"<>|]/.test(seg))) { setProblem('That folder name isn\'t valid.'); return; }
    const path = joinPath(dir, /\.xlsx$/i.test(n) ? n : `${n}.xlsx`);
    if ((files ?? []).some(f => f.path.toLowerCase() === path.toLowerCase())) {
      if (!await confirm(`${baseName(path)} already exists.\nDo you want to replace it?`)) return;
    }
    onDone(path);
  };
  return (
    <Frame title={title} width={520} onClose={() => onDone(null)}>
      <form className="hm-body hm-form" onSubmit={e => { e.preventDefault(); void submit(); }}>
        <label>Folder
          <input className="hm-input" list="hm-folders" value={folder} placeholder="Documents (top level)" aria-label="Folder" onChange={e => { setFolder(e.target.value); setProblem(null); }} />
          <datalist id="hm-folders">{folders.map(f => <option key={f} value={f} />)}</datalist>
        </label>
        <label>File name
          <span className="hm-name-row">
            <input ref={nameRef} className="hm-input" value={name} aria-label="File name" onChange={e => { setName(e.target.value); setProblem(null); }} />
            <span className="hm-ext">.xlsx</span>
          </span>
        </label>
        <div className="hm-type">Save as type: {BRAND.excel} Workbook (*.xlsx)</div>
        {problem && <div className="hm-problem" role="alert">{problem}</div>}
        <button type="submit" hidden />
      </form>
      <div className="hm-btns">
        <button className="hm-btn primary" onClick={() => void submit()}>Save</button>
        <button className="hm-btn" onClick={() => onDone(null)}>Cancel</button>
      </div>
    </Frame>
  );
}

export function BookGlyph({ size = 16 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" aria-hidden className="hm-book">
      <rect x="1.5" y="1.5" width="13" height="13" rx="2.5" fill="#1a7f4b" />
      <rect x="3.5" y="4" width="9" height="8.5" rx=".8" fill="#fff" />
      <path d="M3.5 6.5h9M3.5 9h9M6.5 4v8.5M9.5 4v8.5" stroke="#1a7f4b" strokeWidth=".8" />
    </svg>
  );
}
