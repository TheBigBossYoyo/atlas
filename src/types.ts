export type Theme = 'light' | 'dark' | 'sepia' | 'nord' | 'dracula';
export type ViewMode = 'preview' | 'editor' | 'split';
// UX-12 — 'csv' added for spreadsheet/csv/tsv's real delimited-text export.
// X1 — 'copy' is a format-preserving "Save a copy" passthrough (PDF/XLSX/ODS
// keep their own original bytes; which extension/MIME to use comes from the
// active file's own `format`, not from this value).
export type ExportFormat = 'html' | 'pdf' | 'docx' | 'md' | 'csv' | 'copy';

export interface ThemeMeta {
  id: Theme;
  label: string;
  /** Used for the system titlebar overlay in Electron */
  overlayBg: string;
  overlayFg: string;
}

export const THEMES: readonly ThemeMeta[] = [
  { id: 'light',   label: 'Light',           overlayBg: '#f6f8fa', overlayFg: '#1f2328' },
  { id: 'dark',    label: 'Dark',            overlayBg: '#161b22', overlayFg: '#e6edf3' },
  { id: 'sepia',   label: 'Sepia',           overlayBg: '#f4ecd8', overlayFg: '#5b4636' },
  { id: 'nord',    label: 'Nord',            overlayBg: '#2e3440', overlayFg: '#eceff4' },
  { id: 'dracula', label: 'Dracula',         overlayBg: '#282a36', overlayFg: '#f8f8f2' },
] as const;

export interface TocItem {
  id: string;
  text: string;
  level: number;
}

export interface SearchMatch {
  index: number;
  text: string;
  line: number;
}

export interface RecentFile {
  /** Absolute path (Electron only) — empty in browser */
  path: string;
  name: string;
  /** Epoch ms */
  openedAt: number;
}

export interface DocStats {
  characters: number;
  charactersNoSpaces: number;
  words: number;
  lines: number;
  readingMinutes: number;
}
