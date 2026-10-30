// Page host: Home screen, one workbook view at a time, and the WindowApi the workbook view expects
// (message boxes, Open / Save As pickers over the files API, title, close, open another workbook).
//
// URL parameters:
//   ?file=<path>   open that workbook (the URL follows the open workbook)
//   ?embed=1       no Home screen: open ?file directly (or a blank workbook); Close shows a Reopen page
//   ?panel=1       show the Tasks side panel (scenario tasks and Reset)
import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { ExcelApp } from './excel/ExcelApp';
import type { HostHandle, OpenRequest, WindowApi } from './shell/types';
import { api, baseName, downloadBytes, folderOf } from './api';
import { BRAND, MS_EXCEL } from './util/brand';
import { Home } from './host/Home';
import { TasksPanel } from './host/TasksPanel';
import { MessageBox, OpenPicker, SaveAsDialog, type MessageIcon } from './host/modals';

interface BookState { id: number; file?: string; args?: Record<string, unknown> }
type CloseHandler = () => boolean | Promise<boolean>;

let bookSeq = 0;
let modalSeq = 0;

function focusWorkbook() {
  const el = document.querySelector<HTMLElement>('.xl-root[tabindex]');
  if (el && !el.contains(document.activeElement)) el.focus({ preventScroll: true });
}

export function App() {
  const params = useMemo(() => new URLSearchParams(location.search), []);
  const embed = params.get('embed') === '1';
  const panel = params.get('panel') === '1';
  const [book, setBook] = useState<BookState | null>(() => {
    const f = params.get('file');
    return f ? { id: ++bookSeq, file: f } : embed ? { id: ++bookSeq } : null;
  });
  const [closed, setClosed] = useState(false);
  const [refreshKey, setRefreshKey] = useState(0);
  const [modals, setModals] = useState<{ id: number; node: ReactNode }[]>([]);
  const modalCount = useRef(0);
  modalCount.current = modals.length;
  const closeHandlers = useRef(new Set<CloseHandler>());
  const hostHandle = useRef<{ readonly current: HostHandle } | null>(null);
  const bookRef = useRef(book);
  bookRef.current = book;

  // ------------------------------------------------------------ modals
  const present = useCallback(<T,>(render: (done: (v: T) => void) => ReactNode) => new Promise<T>(resolve => {
    const id = ++modalSeq;
    const done = (v: T) => {
      setModals(m => m.filter(x => x.id !== id));
      resolve(v);
      setTimeout(() => { if (modalCount.current === 0) focusWorkbook(); }, 0);
    };
    setModals(m => [...m, { id, node: render(done) }]);
  }), []);
  const message = useCallback((opts: { title: string; text: string; icon?: MessageIcon; buttons?: string[] }) =>
    present<string>(done => <MessageBox title={opts.title} text={opts.text} icon={opts.icon} buttons={opts.buttons} onDone={done} />), [present]);
  const ask = useCallback((text: string, buttons: string[], title = BRAND.excel) => message({ title, text, icon: 'question', buttons }), [message]);
  const notify = useCallback(async (text: string, icon: 'info' | 'error' = 'info') => { await message({ title: BRAND.excel, text, icon }); }, [message]);

  // ------------------------------------------------------------ navigation
  const urlFor = useCallback((file?: string) => {
    const p = new URLSearchParams();
    if (embed) p.set('embed', '1');
    if (panel) p.set('panel', '1');
    if (file) p.set('file', file);
    const s = p.toString();
    return `${location.pathname}${s ? `?${s}` : ''}`;
  }, [embed, panel]);

  const leaveBook = useCallback(async () => {
    for (const h of [...closeHandlers.current]) if (!(await h())) return false;
    return true;
  }, []);

  const openBook = useCallback(async (file?: string, args?: Record<string, unknown>, opts: { force?: boolean; push?: boolean } = {}) => {
    if (bookRef.current && !opts.force && !(await leaveBook())) return false;
    setClosed(false);
    setBook({ id: ++bookSeq, file, args });
    if (opts.push !== false) history.pushState(null, '', urlFor(file));
    return true;
  }, [leaveBook, urlFor]);

  const goHome = useCallback(async (opts: { push?: boolean } = {}) => {
    if (bookRef.current && !(await leaveBook())) return false;
    setBook(null);
    if (embed) setClosed(true);
    document.title = BRAND.excel;
    if (opts.push !== false) history.pushState(null, '', urlFor());
    setRefreshKey(k => k + 1);
    return true;
  }, [embed, leaveBook, urlFor]);

  useEffect(() => {
    const onPop = async () => {
      const f = new URLSearchParams(location.search).get('file') ?? undefined;
      const cur = bookRef.current;
      if (f && cur?.file === f) return;
      const ok = f ? await openBook(f, undefined, { push: false }) : await goHome({ push: false });
      if (!ok) history.pushState(null, '', urlFor(cur?.file));
    };
    window.addEventListener('popstate', onPop);
    return () => window.removeEventListener('popstate', onPop);
  }, [goHome, openBook, urlFor]);

  const actions = useRef({ openBook, goHome });
  actions.current = { openBook, goHome };

  // ------------------------------------------------------------ the workbook view's window
  const win = useMemo<WindowApi | null>(() => {
    if (!book) return null;
    const id = `book-${book.id}`;
    const confirmReplace = (text: string) => present<string>(d => <MessageBox title="Confirm Save As" text={text} icon="warning" buttons={['Yes', 'No']} onDone={d} />).then(b => b === 'Yes');
    const currentFolder = () => folderOf(hostHandle.current?.current.path ?? book.file ?? '');
    const pickOpen = (title?: string) => present<string | null>(done => <OpenPicker title={title ?? 'Open'} onDone={done} />);
    const pickSave = (title?: string, defaultName?: string) => present<string | null>(done => (
      <SaveAsDialog title={title ?? 'Save As'} defaultName={defaultName ?? 'Book1.xlsx'} defaultFolder={currentFolder()} confirm={confirmReplace} onDone={done} />
    ));
    const w: WindowApi = {
      id,
      setTitle(title) {
        document.title = title;
        // Save As moves the workbook to a new path: keep the address bar on it
        const p = hostHandle.current?.current.path;
        if (p && bookRef.current?.id === book.id && new URLSearchParams(location.search).get('file') !== p) history.replaceState(null, '', urlFor(p));
      },
      close() { void actions.current.goHome(); },
      minimize() { /* single page */ },
      maximize() { /* single page */ },
      restore() { /* single page */ },
      focus() { focusWorkbook(); },
      isFocused: () => modalCount.current === 0,
      isMaximized: () => true,
      open(app, req?: OpenRequest) {
        if (app !== 'excel') return;
        void actions.current.openBook(req?.file, req?.args);
      },
      dialog: {
        message: opts => message({ title: opts.title, text: opts.text, icon: (opts.icon ?? 'info') as MessageIcon, buttons: opts.buttons }),
        openFile: opts => pickOpen(opts.title),
        openFileEx: async opts => { const p = await pickOpen(opts.title); return p ? { path: p, paths: [p], mode: 'open' as const } : null; },
        saveFile: opts => pickSave(opts.title, opts.defaultName),
        saveFileEx: async opts => { const p = await pickSave(opts.title, opts.defaultName); return p ? { path: p, filterIndex: 0, ext: 'xlsx' } : null; },
      },
      onBeforeClose(h) { closeHandlers.current.add(h); return () => { closeHandlers.current.delete(h); }; },
      onRequest() { return () => {}; },
      windows: () => [{ id, appId: 'excel', title: document.title, minimized: false, maximized: true, focused: true }],
      focusWindow() { focusWorkbook(); },
      closeWindow() { void actions.current.goHome(); },
      bindHost(handle) { hostHandle.current = handle; return () => { if (hostHandle.current === handle) hostHandle.current = null; }; },
    };
    return w;
  }, [book?.id]); // eslint-disable-line react-hooks/exhaustive-deps

  // focus the grid once the workbook view is up
  useEffect(() => {
    if (!book) return;
    let n = 0;
    const t = setInterval(() => {
      if (document.querySelector('.xl-root[tabindex]') || ++n > 100) { clearInterval(t); if (modalCount.current === 0 && (document.activeElement === document.body || !document.activeElement)) focusWorkbook(); }
    }, 50);
    return () => clearInterval(t);
  }, [book?.id]); // eslint-disable-line react-hooks/exhaustive-deps

  // Ctrl+S anywhere on the page saves the open workbook (never the browser's Save Page)
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey) || e.key.toLowerCase() !== 's' || e.altKey) return;
      if (e.defaultPrevented) return;   // the workbook view handled it
      e.preventDefault();
      if (modalCount.current > 0 || !bookRef.current) return;
      void hostHandle.current?.current.save();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);

  const download = async () => {
    const h = hostHandle.current?.current;
    if (!h) return;
    if (!h.path || h.dirty) {
      const b = await message({
        title: MS_EXCEL, icon: 'question',
        text: !h.path ? 'This workbook hasn\'t been saved yet. Save it before downloading?' : 'Download includes saved changes only. Save your changes first?',
        buttons: h.path ? ['Save and Download', 'Download Saved Copy', 'Cancel'] : ['Save and Download', 'Cancel'],
      });
      if (b === 'Cancel') return;
      if (b === 'Save and Download' && !(await h.save())) return;
    }
    const path = hostHandle.current?.current.path;
    if (!path) return;
    try { downloadBytes(await api.getRaw(path), baseName(path)); }
    catch (e) { await notify(`Couldn't download ${baseName(path)}. ${e instanceof Error ? e.message : ''}`, 'error'); }
  };

  const onReset = () => {
    setRefreshKey(k => k + 1);
    const cur = bookRef.current;
    if (cur?.file) void openBook(cur.file, undefined, { force: true, push: false });
  };

  let main: ReactNode;
  if (book && win) {
    main = (
      <div className="book">
        <ExcelApp key={book.id} win={win} req={{ file: book.file, args: book.args }} />
        <div className="book-controls">
          <button className="book-ctl" onClick={() => void download()} title="Download (.xlsx)" aria-label="Download">
            <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden><path d="M8 2.5V11M4.5 7.5 8 11l3.5-3.5M2.5 13.5h11" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" /></svg>
            <span>Download</span>
          </button>
          {!embed && (
            <button className="book-ctl" onClick={() => void goHome()} title="Close and go to Home" aria-label="Home">
              <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden><path d="M2.5 7.5 8 2.5l5.5 5M4 6.5v7h3v-4h2v4h3v-7" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" /></svg>
              <span>Home</span>
            </button>
          )}
        </div>
      </div>
    );
  } else if (embed || closed) {
    main = (
      <div className="closed-page">
        <p>The workbook is closed.</p>
        <button className="home-btn" onClick={() => void openBook(params.get('file') ?? undefined)}>Reopen</button>
      </div>
    );
  } else {
    main = <Home onOpen={p => void openBook(p)} onBlank={() => void openBook(undefined, { name: 'Book1' })} ask={ask} notify={notify} refreshKey={refreshKey} />;
  }

  return (
    <div className={`page${panel ? ' with-panel' : ''}`}>
      <div className="page-main">{main}</div>
      {panel && <TasksPanel onOpen={p => void openBook(p)} onReset={onReset} ask={ask} />}
      {modals.map(m => <div key={m.id} className="hm-layer">{m.node}</div>)}
    </div>
  );
}
