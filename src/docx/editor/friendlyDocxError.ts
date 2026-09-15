/**
 * RUN-14 — friendly, format-specific DOCX load/save error messages.
 *
 * `loadDocx`/`saveDocx` surface failures as plain `Error`s (frequently a
 * `DocxParseError`/`DocxSaveError`, but sometimes a raw exception from a
 * dependency — JSZip's "Can't find end of central directory", or a
 * `fast-xml-parser` parse failure — that leaked past a call site D28 hasn't
 * hardened yet). Showing that verbatim to an end user ("Failed to render
 * DOCX: Corrupted zip: can't find end of central directory") reads as a
 * crash report, not guidance. This module classifies the message by a small
 * set of known substrings and returns a short, actionable sentence instead,
 * while still appending the original detail in parentheses so the message
 * remains useful for bug reports/support.
 */

const ZIP_CORRUPTION_PATTERN =
  /central directory|not a valid zip|corrupted zip|invalid signature|unexpected signature|wrong (local file header|password)|invalid crc32/i

const ZIP_BOMB_PATTERN = /zip bomb|uncompressed size|per-entry limit|total limit/i

const XML_CORRUPTION_PATTERN =
  /unclosed tag|invalid char|unexpected end of (data|input|file)|xml.*(pars|declaration)|not well-formed|mismatched tag/i

function extractDetail(error: unknown): string {
  if (error instanceof Error) {
    return error.message
  }

  return String(error)
}

/**
 * @param error   The value caught from `loadDocx`/`saveDocx` (or anything
 *                further down the DOCX pipeline).
 * @param action  Which operation failed, so the leading sentence reads
 *                naturally either way.
 */
export function friendlyDocxErrorMessage(error: unknown, action: 'open' | 'save'): string {
  const detail = extractDetail(error)
  const name = error instanceof Error ? error.name : undefined

  if (ZIP_BOMB_PATTERN.test(detail)) {
    // Already a clear, specific message (DocxParseError's own zip-bomb
    // guard) — just frame which operation it interrupted.
    return action === 'open'
      ? `This file was not opened: ${detail}`
      : `This file was not saved: ${detail}`
  }

  if (ZIP_CORRUPTION_PATTERN.test(detail)) {
    return action === 'open'
      ? `This file doesn't appear to be a valid Word document (.docx). It may be corrupted, password-protected, or a different file type entirely. (${detail})`
      : `The document could not be saved because its file data appears corrupted. Try again, or save a copy under a new name. (${detail})`
  }

  if (XML_CORRUPTION_PATTERN.test(detail)) {
    return action === 'open'
      ? `This document's internal structure appears to be corrupted and could not be read. It may have been damaged by another application. (${detail})`
      : `The document could not be saved because Atlas generated XML that failed its own validity check. Your changes were not written to disk — please try again or report this issue. (${detail})`
  }

  if (name === 'DocxSaveError') {
    return `Atlas could not verify the saved file was valid, so nothing was written to disk. Please try again or report this issue. (${detail})`
  }

  if (name === 'DocxParseError') {
    return action === 'open'
      ? `This Word document could not be opened. (${detail})`
      : `This Word document could not be saved. (${detail})`
  }

  return action === 'open'
    ? `Couldn't open this Word document. (${detail})`
    : `Couldn't save this Word document. Your changes have not been written to disk. (${detail})`
}
