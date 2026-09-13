export interface RecentFilesStore {
  record(rawPath: unknown): void;
  has(rawPath: unknown): boolean;
}

export function createRecentFilesStore(storeDir: string): RecentFilesStore;
export const STORE_FILE_NAME: string;
