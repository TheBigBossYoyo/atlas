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

/**
 * X5/save-error-classification — maps a raw Node `fs` error's `.code` to a
 * friendly message, or `undefined` when there's no specific mapping for it.
 * (Declaration was missing this export entirely — found during wave3/export
 * review: any `.ts` consumer importing it, e.g. a unit test, failed to
 * compile with "has no exported member 'classifyWriteError'" even though the
 * runtime export has existed since this file's own X5 commit.)
 */
export function classifyWriteError(err: unknown): string | undefined;

/**
 * i18n — the stable-key counterpart to `classifyWriteError`, one of
 * `'fileLocked' | 'permissionDenied' | 'diskFull' | 'isDirectory' |
 * 'destinationMissing' | 'readOnly' | 'nameTooLong'`, or `undefined` for an
 * unclassified error. The renderer maps this to `src/i18n`'s `errors.write.*`
 * catalogue keys instead of displaying `classifyWriteError`'s English text
 * directly, so the message is translated in a French UI.
 */
export function classifyWriteErrorCode(err: unknown): string | undefined;
