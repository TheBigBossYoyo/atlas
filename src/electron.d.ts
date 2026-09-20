import type { Theme } from './types';

interface ElectronFileData {
  content: string;
  name: string;
  path: string;
}

interface SaveFileResult {
  saved: boolean;
  path?: string;
  name?: string;
  /** Present on failure when a specific, user-friendly reason is known (e.g. the file is locked by another program). English fallback text — prefer `errorCode` (see `src/i18n`) for a translated message. */
  error?: string;
  /** Stable key for `error` (`errors.write.*` in `src/i18n`), set whenever `classifyWriteError`/`FileLockedError` recognized the failure. Absent for an unclassified error, where `error`'s English text is the only signal available. */
  errorCode?: string;
}

interface SaveFileRequest {
  content: string;
  suggestedName: string;
  /** Optional dialog filters; defaults to Markdown */
  filters?: Array<{ name: string; extensions: string[] }>;
  /** If provided, save there silently (no dialog) */
  existingPath?: string;
}

interface BinarySaveFileRequest {
  content: Uint8Array;
  suggestedName: string;
  /** Optional dialog filters; defaults to Word Documents */
  filters?: Array<{ name: string; extensions: string[] }>;
  /** If provided, save there silently (no dialog) */
  existingPath?: string;
}

interface SpellCheckOperationResult {
  readonly added?: boolean;
  readonly replaced?: boolean;
}

interface RegisterPathResult {
  readonly ok: boolean;
}

/** NEW-01 — the fixed set of document types the toolbar's "New" action can create; kept in sync with `electron/lib/newDocumentTemplates.cjs`'s `NEW_DOCUMENT_FORMATS`. */
export type NewDocumentFormat = 'markdown' | 'docx' | 'xlsx' | 'ods' | 'pptx' | 'odp';

/** Result of `document:new` — a native Save dialog was shown and either a blank document was written at the chosen path, or the user cancelled. */
type NewDocumentResult =
  | { readonly created: true; readonly path: string }
  | { readonly created: false; readonly error?: string; readonly errorCode?: string };

/** Result of `export:printToPdf` — X1's real per-format PDF export. */
type PrintToPdfResult =
  | { readonly ok: true; readonly bytes: Uint8Array }
  | { readonly ok: false; readonly error: string };

interface ElectronAPI {
  getInitialFile: () => Promise<ElectronFileData | { path: string } | null>;
  openFileDialog: () => Promise<ElectronFileData | null>;
  openFileByPath: (path: string) => Promise<ElectronFileData | null>;
  saveFile: (req: SaveFileRequest) => Promise<SaveFileResult>;
  saveBinaryFile: (req: BinarySaveFileRequest) => Promise<SaveFileResult>;
  /**
   * X1 — renders a self-contained HTML document (the caller must sanitize it
   * first — see `src/utils/export/sanitizeExportHtml.ts`) to a vector PDF via
   * a hidden, script-disabled `BrowserWindow`. Optional so existing test
   * doubles that only implement the file-open/save surface keep compiling.
   */
  printToPdf?: (html: string) => Promise<PrintToPdfResult>;
  onFileOpened: (callback: (data: ElectronFileData) => void) => () => void;
  setTheme: (theme: Theme) => void;
  /** i18n — backs the renderer's "System" language option; see `src/i18n`. Optional so existing test doubles keep compiling. */
  getLocale?: () => Promise<string>;
  /**
   * P2.5/SHELL-02/ELEC-06 — pushes the renderer's combined dirty state to
   * main's window `close` handler. Optional (like `image`/`spellcheck`
   * below) so existing test doubles that only implement the file-open/save
   * surface keep compiling; every real call site invokes it defensively
   * (`window.electronAPI?.notifyDirtyState?.(...)`).
   */
  notifyDirtyState?: (dirty: boolean) => void;
  /** Main asks the renderer to save (the user chose "Save" in the native close-confirmation prompt). */
  onRequestSaveBeforeClose?: (callback: () => void) => () => void;
  /** The renderer reports whether that save succeeded so main knows whether to actually close the window. */
  reportSaveBeforeCloseResult?: (result: { saved: boolean }) => void;
  openFileBinary: () => Promise<{ canceled: boolean; path: string; buffer: ArrayBuffer }>;
  /** NEW-01 — the toolbar's "New" action / Ctrl+N: shows a native Save dialog, writes a blank template there, and reports the resulting path. */
  newDocument: (formatId: NewDocumentFormat) => Promise<NewDocumentResult>;
  readBinaryByPath: (path: string) => Promise<{ path: string; buffer: ArrayBuffer }>;
  onFileOpenedPath: (callback: (path: string) => void) => () => void;
  /** Resolves a dropped `File` to its absolute path (Electron 32+ removed `File.path`). */
  getPathForFile: (file: File) => string;
  /** Registers a drag-dropped path into the main process's read/write allowlist. */
  registerDroppedPath: (path: string) => Promise<RegisterPathResult>;
  /** Re-validates and registers a recent-file path before it is reopened. */
  requestOpenRecent: (path: string) => Promise<RegisterPathResult>;
  /** Reveals an already-allowlisted path in the OS file manager. */
  revealInFolder: (path: string) => Promise<{ ok: boolean }>;
  image?: {
    pick: () => Promise<ImagePickResult>;
  };
  /** USR-19 — explicit, confirmed run of an opened source file. */
  codeRun?: {
    start: (path: string) => Promise<{ ok: boolean; runId?: number; error?: string; cancelled?: boolean }>;
    stop: (runId: number) => Promise<boolean>;
    onOutput: (
      callback: (payload: { runId: number; stream: "stdout" | "stderr" | "system"; text: string }) => void,
    ) => () => void;
    onExit: (
      callback: (payload: {
        runId: number;
        code: number | null;
        timedOut: boolean;
        stopped: boolean;
        error?: string;
      }) => void,
    ) => () => void;
  };
  /** USR-11 — installed font family names (empty outside Windows). */
  fonts?: {
    list: () => Promise<string[]>;
  };
  spellcheck: {
    onContextMenu: (
      callback: (payload: SpellCheckContextMenuPayload) => void,
    ) => () => void;
    replaceMisspelling: (word: string) => Promise<SpellCheckOperationResult>;
    addWord: (word: string) => Promise<{ added: boolean }>;
    getLanguages: () => Promise<{ available: ReadonlyArray<string>; enabled: ReadonlyArray<string> }>;
    setLanguages: (
      languages: ReadonlyArray<string>,
    ) => Promise<{ ok: boolean; enabled?: ReadonlyArray<string> }>;
  };
}

export interface SpellCheckContextMenuPayload {
  readonly word: string;
  readonly suggestions: ReadonlyArray<string>;
  readonly x: number;
  readonly y: number;
}

export type ImagePickResult =
  | { readonly cancelled: true; readonly error?: string }
  | {
      readonly cancelled: false;
      readonly bytes: Uint8Array;
      readonly mime: 'image/png' | 'image/jpeg' | 'image/gif';
      readonly suggestedName: string;
    };

declare global {
  interface Window {
    electronAPI?: ElectronAPI;
  }
}

export {};
