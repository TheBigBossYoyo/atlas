/**
 * Atlas — DOCX header parser (Wave A.5)
 *
 * OPTION B: stores body content as a single UnknownNode containing the raw
 * inner XML of <w:hdr>.  Wave A.6/B.x will consolidate via a shared body
 * parser once document.ts (A.3) ships the paragraph/run/table logic.
 *
 * The `id` parameter is the relationship-derived file identifier (e.g. "rId1")
 * that the orchestrator must supply; the header XML itself carries no self-id.
 */

import { XMLParser } from 'fast-xml-parser'

import { DocxParseError } from './unzip'
import type { Header, UnknownNode, Block } from '../model/document'

// ---------------------------------------------------------------------------
// Parser instance — same config as Wave A.1
// ---------------------------------------------------------------------------

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

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
  try {
    // Validate that the XML is at least parseable; we do not use the result
    // here because Option B defers content parsing to a future shared parser.
    xmlParser.parse(xml)
  } catch (cause) {
    const msg = cause instanceof Error ? cause.message : String(cause)
    throw new DocxParseError(`Failed to parse header XML: ${msg}`, id)
  }

  const inner: UnknownNode = { kind: 'unknown', xml }
  const blocks: ReadonlyArray<Block> = [inner]

  return { kind: 'header', id, blocks }
}
