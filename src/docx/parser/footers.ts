/**
 * Atlas — DOCX footer parser (Wave A.5, hardened in P1.3)
 *
 * Parses the body content of a `word/footer*.xml` part into real
 * `Paragraph[]` blocks by porting `comments.ts`'s synthetic-wrapper
 * technique (see `partBody.ts`), replacing the earlier "Option B" stub that
 * stored the whole part as a single opaque `UnknownNode` and made the
 * serializer throw on every save that included a footer (DXP-04/DXS-01).
 *
 * The `id` parameter is the relationship-derived file identifier (e.g. "rId2")
 * that the orchestrator must supply; the footer XML itself carries no self-id.
 */

import { XMLParser } from 'fast-xml-parser'

import { parseParagraphsFromRawNodes } from './partBody'
import { DocxParseError } from './unzip'
import { assertXmlPartSizeWithinLimit } from './xmlSizeGuard'
import type { Footer } from '../model/document'

// ---------------------------------------------------------------------------
// Parser instance — same config as Wave A.1
// ---------------------------------------------------------------------------

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

// ---------------------------------------------------------------------------
// Internal XML shapes
// ---------------------------------------------------------------------------

interface RawFooterDoc {
  'w:ftr'?: {
    'w:p'?: unknown
    [key: string]: unknown
  }
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Parse the XML content of a `word/footer*.xml` part.
 *
 * @param xml - UTF-8 text of the footer XML file.
 * @param id  - Relationship id that the orchestrator resolved for this footer.
 *              Defaults to `''`; callers should always supply the real id.
 */
export function parseFooter(xml: string, id: string = ''): Footer {
  assertXmlPartSizeWithinLimit(xml, id === '' ? 'word/footer*.xml' : `word/footer*.xml (${id})`)
  let parsed: RawFooterDoc
  try {
    parsed = xmlParser.parse(xml) as RawFooterDoc
  } catch (cause) {
    const msg = cause instanceof Error ? cause.message : String(cause)
    throw new DocxParseError(`Failed to parse footer XML: ${msg}`, id)
  }

  const blocks = parseParagraphsFromRawNodes(parsed?.['w:ftr']?.['w:p'])

  return { kind: 'footer', id, blocks }
}
