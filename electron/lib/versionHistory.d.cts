export interface VersionEntry {
  readonly id: string;
  readonly at: number;
  readonly bytes: number;
  readonly label: string | null;
}

export type SnapshotResult =
  | { readonly stored: true; readonly id: string; readonly at: number }
  | { readonly stored: false; readonly reason: 'unchanged'; readonly id: string }
  | { readonly stored: false; readonly reason: 'invalidPath'; readonly id: '' };

export interface VersionHistory {
  snapshot(documentPath: string, bytes: Buffer, label?: string | null): SnapshotResult;
  list(documentPath: string): VersionEntry[];
  read(documentPath: string, id: string): Buffer | null;
  clear(documentPath: string): void;
}

export function createVersionHistory(rootDir: string): VersionHistory;
export const INDEX_FILE_NAME: string;
export const MAX_VERSIONS_PER_DOCUMENT: number;
export const MAX_BYTES_PER_DOCUMENT: number;
