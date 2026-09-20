import type { TranslateFn } from './types';

/** The shape both `SaveFileResult` and the failure branch of `NewDocumentResult` share. */
interface WriteErrorLike {
  readonly error?: string;
  readonly errorCode?: string;
}

const KNOWN_WRITE_ERROR_CODES: ReadonlySet<string> = new Set([
  'permissionDenied',
  'diskFull',
  'isDirectory',
  'destinationMissing',
  'readOnly',
  'nameTooLong',
  'fileLocked',
  'unknownNewDocument',
]);

/**
 * Resolves a failed save/new-document IPC result to a translated message.
 * `main.cjs` (via `electron/lib/atomicWrite.cjs`'s `classifyWriteErrorCode`)
 * attaches a stable `errorCode` for every failure class it recognizes —
 * this is the ONLY place that code is turned into user-facing text, so the
 * message is always in the current UI language rather than main's English
 * fallback. Falls back to `result.error` (English) when there's no
 * recognized code, which only happens for a genuinely unclassified
 * filesystem error.
 */
export function translateWriteError(t: TranslateFn, result: WriteErrorLike): string | null {
  if (result.errorCode && KNOWN_WRITE_ERROR_CODES.has(result.errorCode)) {
    return t(`errors.write.${result.errorCode}`);
  }
  return result.error ?? null;
}
