export declare class FileLockedError extends Error {
  constructor(message?: string);
}

export const LOCK_ERROR_MESSAGE: string;

export interface AtomicWriteResult {
  readonly fallbackUsed: boolean;
}

export function atomicWriteFile(
  targetPath: string,
  data: string | Uint8Array,
): AtomicWriteResult;
