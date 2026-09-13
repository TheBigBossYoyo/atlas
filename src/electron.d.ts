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
  /** Present on failure when a specific, user-friendly reason is known (e.g. the file is locked by another program). */
  error?: string;
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

interface ElectronAPI {
  getInitialFile: () => Promise<ElectronFileData | { path: string } | null>;
  openFileDialog: () => Promise<ElectronFileData | null>;
  openFileByPath: (path: string) => Promise<ElectronFileData | null>;
  saveFile: (req: SaveFileRequest) => Promise<SaveFileResult>;
  saveBinaryFile: (req: BinarySaveFileRequest) => Promise<SaveFileResult>;
  onFileOpened: (callback: (data: ElectronFileData) => void) => () => void;
  setTheme: (theme: Theme) => void;
  openFileBinary: () => Promise<{ canceled: boolean; path: string; buffer: ArrayBuffer }>;
  readBinaryByPath: (path: string) => Promise<{ path: string; buffer: ArrayBuffer }>;
  onFileOpenedPath: (callback: (path: string) => void) => () => void;
  /** Resolves a dropped `File` to its absolute path (Electron 32+ removed `File.path`). */
  getPathForFile: (file: File) => string;
  /** Registers a drag-dropped path into the main process's read/write allowlist. */
  registerDroppedPath: (path: string) => Promise<RegisterPathResult>;
  /** Re-validates and registers a recent-file path before it is reopened. */
  requestOpenRecent: (path: string) => Promise<RegisterPathResult>;
  image?: {
    pick: () => Promise<ImagePickResult>;
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
