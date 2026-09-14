/**
 * RUN-14 — wraps a raw error thrown by a third-party parsing/generation
 * library (JSZip, docx, pdfjs-dist, ...) into a friendly, format-specific
 * message. These libraries throw developer-facing strings ("Can't find end
 * of central directory", "Invalid PDF structure", ...) that mean nothing to
 * an end user; this maps the common ones to plain language while still
 * preserving the original message for diagnostics (console/log only).
 */

/** [pattern to match against the raw message, plain-language explanation] */
const KNOWN_LIBRARY_ERRORS: ReadonlyArray<readonly [RegExp, string]> = [
  [/central directory|not a valid zip|corrupted zip/i, 'the file could not be read as a valid Office document (it may be corrupted or not a real Office file)'],
  [/invalid signature|unsupported compression/i, 'the file is not in a format this app recognizes'],
  [/out of memory|allocation failed|maximum call stack/i, 'the document is too large to process'],
  [/invalid pdf structure|bad xref|invalid.*pdf/i, 'the PDF file appears to be corrupted or malformed'],
  [/password|encrypted/i, 'the file is password-protected and cannot be opened'],
];

function friendlyLibraryMessage(rawMessage: string): string | null {
  for (const [pattern, friendly] of KNOWN_LIBRARY_ERRORS) {
    if (pattern.test(rawMessage)) return friendly;
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
