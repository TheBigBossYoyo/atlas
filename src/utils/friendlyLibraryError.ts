/**
 * RUN-14 — wraps a raw error thrown by a third-party parsing/generation
 * library (JSZip, docx, pdfjs-dist, ...) into a friendly, format-specific
 * message. These libraries throw developer-facing strings ("Can't find end
 * of central directory", "Invalid PDF structure", ...) that mean nothing to
 * an end user; this maps the common ones to plain language while still
 * preserving the original message for diagnostics (console/log only).
 *
 * i18n — this file lives in the renderer (unlike
 * `electron/lib/atomicWrite.cjs`, which can't reach the translation
 * catalogue at all), so the friendly text is looked up through `t()`
 * directly here, in whatever language is active when the error is built,
 * rather than embedding English (or hard-coding French) in the pattern
 * table itself.
 */
import { t } from '../i18n';

/** [pattern to match against the raw message, catalogue key for the plain-language explanation] */
const KNOWN_LIBRARY_ERRORS: ReadonlyArray<readonly [RegExp, string]> = [
  [/central directory|not a valid zip|corrupted zip/i, 'errors.library.corruptZip'],
  [/invalid signature|unsupported compression/i, 'errors.library.unsupportedFormat'],
  [/out of memory|allocation failed|maximum call stack/i, 'errors.library.tooLarge'],
  [/invalid pdf structure|bad xref|invalid.*pdf/i, 'errors.library.corruptPdf'],
  [/password|encrypted/i, 'errors.library.passwordProtected'],
];

function friendlyLibraryMessage(rawMessage: string): string | null {
  for (const [pattern, key] of KNOWN_LIBRARY_ERRORS) {
    if (pattern.test(rawMessage)) return t(key);
  }
  return null;
}

/**
 * Turns `err` into a user-friendly `Error` prefixed with `context`
 * (e.g. "DOCX export", "Opening this PDF"). Falls back to the raw message
 * when it doesn't match a known library failure signature, so nothing is
 * ever silently hidden — just made readable when we can recognize it.
 */
export function toFriendlyError(err: unknown, context: string): Error {
  const rawMessage = err instanceof Error ? err.message : String(err);
  const friendly = friendlyLibraryMessage(rawMessage);
  return new Error(friendly ? `${context}: ${friendly}` : `${context}: ${rawMessage}`);
}
