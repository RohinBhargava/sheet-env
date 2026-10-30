// The contract between the page host (src/App.tsx) and the workbook view. The workbook view was written
// against a window manager; in the browser app there is one "window" (the page) and the host implements
// this interface with in-page modals and pickers over the files API.

export type AppId = 'excel' | string;

export interface OpenRequest {
  /** a path under the user's Documents root, forward slashes, e.g. "Capital Calls/Notice.xlsx" */
  file?: string;
  /** view arguments, e.g. { protectedView: true } or { name: 'Book2' } */
  args?: Record<string, unknown>;
  newWindow?: boolean;
}

export interface WindowInfo {
  id: string;
  appId: AppId;
  title: string;
  minimized: boolean;
  maximized: boolean;
  focused: boolean;
}

export interface WindowApi {
  id: string;
  setTitle(title: string): void;
  close(): void;
  minimize(): void;
  maximize(): void;
  restore(): void;
  focus(): void;
  isFocused(): boolean;
  /** open another workbook (open('excel', { file })) */
  open(app: AppId, req?: OpenRequest): void;
  /** modal dialogs owned by this window; resolves with the button pressed */
  dialog: DialogApi;
  isMaximized?(): boolean;
  /** called before the view closes; return false (or resolve false) to keep it open */
  onBeforeClose?(handler: () => boolean | Promise<boolean>): () => void;
  onRequest?(handler: (req: OpenRequest) => void): () => void;
  windows?(): WindowInfo[];
  focusWindow?(id: string): void;
  closeWindow?(id: string): void;
  /** the workbook view hands the host a live handle (Save, dirty flag, path); returns an unbind function */
  bindHost?(handle: { readonly current: HostHandle }): () => void;
}

export interface HostHandle {
  save(): Promise<boolean>;
  dirty: boolean;
  path: string | undefined;
  name: string;
}

export interface FileFilter { label: string; ext: string[] }

export interface DialogApi {
  message(opts: { title: string; text: string; icon?: 'info' | 'warning' | 'error' | 'question'; buttons?: string[]; variant?: 'office' | 'classic'; width?: number; height?: number; at?: { x: number; y: number }; className?: string }): Promise<string>;
  openFile(opts: { title?: string; filters?: FileFilter[]; startIn?: string }): Promise<string | null>;
  saveFile(opts: { title?: string; filters?: FileFilter[]; startIn?: string; defaultName?: string }): Promise<string | null>;
  openFileEx?(opts: OpenFileOptions): Promise<OpenFileResult | null>;
  saveFileEx?(opts: SaveFileOptions): Promise<SaveFileResult | null>;
}

export interface OpenFileOptions {
  title?: string;
  filters?: FileFilter[];
  startIn?: string;
  okLabel?: string;
  splitMenu?: boolean;
  multiple?: boolean;
  controlId?: string;
}
export interface OpenFileResult {
  path: string;
  paths: string[];
  mode: 'open' | 'readOnly' | 'copy' | 'protectedView' | 'repair';
}

export interface SaveFileOptions {
  title?: string;
  filters?: FileFilter[];
  filterIndex?: number;
  startIn?: string;
  defaultName?: string;
  authors?: string;
  officeFields?: boolean;
  controlId?: string;
}
export interface SaveFileResult {
  path: string;
  filterIndex: number;
  ext: string;
  authors?: string;
  tags?: string;
  title?: string;
}

export interface AppProps {
  win: WindowApi;
  req: OpenRequest;
}
