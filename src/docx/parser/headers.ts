/**
 * Atlas — DOCX header parser (Wave A.5, hardened in P1.3)
 *
 * Parses the body content of a `word/header*.xml` part into real
 * `Paragraph[]` blocks by porting `comments.ts`'s synthetic-wrapper
 * technique (see `partBody.ts`), replacing the earlier "Option B" stub that
 * stored the whole part as a single opaque `UnknownNode` and made the
 * serializer throw on every save that included a header (DXP-04/DXS-01).
 *
 * The `id` parameter is the relationship-derived file identifier (e.g. "rId1")
 * that the orchestrator must supply; the header XML itself carries no self-id.
 */

import { XMLParser } from 'fast-xml-parser'

import { parseParagraphsFromRawNodes } from './partBody'
import { DocxParseError } from './unzip'
import type { Header } from '../model/document'

// ---------------------------------------------------------------------------
// Parser instance — same config as Wave A.1
// ---------------------------------------------------------------------------

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

// ---------------------------------------------------------------------------
// Internal XML shapes
// ---------------------------------------------------------------------------

interface RawHeaderDoc {
  'w:hdr'?: {
    'w:p'?: unknown
    [key: string]: unknown
  }
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Parse the XML content of a `word/header*.xml` part.
 *
 * @param xml - UTF-8 text of the header XML file.
 * @param id  - Relationship id that the orchestrator resolved for this header.
 *              Defaults to `''`; callers should always supply the real id.
 */
export function parseHeader(xml: string, id: string = ''): Header {
  let parsed: RawHeaderDoc
  try {
    parsed = xmlParser.parse(xml) as RawHeaderDoc
  } catch (cause) {
    const msg = cause instanceof Error ? cause.message : String(cause)
    throw new DocxParseError(`Failed to parse header XML: ${msg}`, id)
  }

  const blocks = parseParagraphsFromRawNodes(parsed?.['w:hdr']?.['w:p'])

  return { kind: 'header', id, blocks }
}
