import { t } from '../../i18n'

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
      ? t('docx.error.notOpenedDetail', { detail })
      : t('docx.error.notSavedDetail', { detail })
  }

  if (ZIP_CORRUPTION_PATTERN.test(detail)) {
    return action === 'open'
      ? t('docx.error.invalidWordFile', { detail })
      : t('docx.error.saveCorruptedData', { detail })
  }

  if (XML_CORRUPTION_PATTERN.test(detail)) {
    return action === 'open'
      ? t('docx.error.openXmlCorrupted', { detail })
      : t('docx.error.saveXmlInvalid', { detail })
  }

  if (name === 'DocxSaveError') {
    return t('docx.error.saveVerifyFailed', { detail })
  }

  if (name === 'DocxParseError') {
    return action === 'open'
      ? t('docx.error.openParseFailed', { detail })
      : t('docx.error.saveParseFailed', { detail })
  }

  return action === 'open'
    ? t('docx.error.openGenericFailed', { detail })
    : t('docx.error.saveGenericFailed', { detail })
}
