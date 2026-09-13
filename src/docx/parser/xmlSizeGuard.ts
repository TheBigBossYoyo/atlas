/**
 * Atlas — defense-in-depth size guard for untrusted OOXML XML parts (D21 / DXP-20)
 *
 * `unzip.ts`'s per-entry/running-total budget (DXP-15) rejects a classic zip
 * bomb — a small compressed archive that decompresses to a huge total — but
 * it says nothing about a single XML part that is merely large, or one that
 * is pathologically deep/wide (thousands of nested elements packed into a
 * modest byte count). Either shape can make `fast-xml-parser` build an
 * enormous or extremely slow-to-construct result tree from a `.docx` whose
 * *compressed and uncompressed* sizes both look innocuous.
 *
 * This guard runs immediately before every `XMLParser.parse()` call in the
 * DOCX parser pipeline (`document.ts`, `styles.ts`, `numbering.ts`,
 * `headers.ts`, `footers.ts`, `footnotes.ts`, `endnotes.ts`, `comments.ts`,
 * `relationships.ts`, `contentTypes.ts`, `theme.ts`) and rejects any part
 * above a conservative ceiling with a clear `DocxParseError`, before the
 * parser ever sees it.
 */

import { DocxParseError } from './unzip'

/**
 * Conservative ceiling for a single OOXML XML part, in UTF-16 code units
 * (`string.length`, an O(1) check). Far larger than any legitimate
 * `document.xml`/`styles.xml`/etc. produced by Word or any other real
 * authoring tool — this exists to catch a maliciously-crafted part, not to
 * constrain normal documents.
 */
export const MAX_XML_PART_LENGTH = 20 * 1024 * 1024

/**
 * Throws `DocxParseError` when `xml` exceeds {@link MAX_XML_PART_LENGTH}.
 *
 * @param xml - Raw XML text about to be handed to `XMLParser.parse()`.
 * @param partLabel - Human-readable identifier for the offending part (e.g.
 *   `"word/document.xml"`), used in the error message and as the thrown
 *   error's `path`.
 */
export function assertXmlPartSizeWithinLimit(xml: string, partLabel: string): void {
  if (xml.length <= MAX_XML_PART_LENGTH) {
    return
  }

  throw new DocxParseError(
    `"${partLabel}" is ${xml.length.toLocaleString()} characters, over the ` +
      `${MAX_XML_PART_LENGTH.toLocaleString()}-character safety limit for a single DOCX XML part; ` +
      'refusing to parse it.',
    partLabel,
  )
}
