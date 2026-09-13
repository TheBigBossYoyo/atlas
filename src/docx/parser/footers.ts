/**
 * Atlas — DOCX footer parser (Wave A.5)
 *
 * OPTION B: stores body content as a single UnknownNode containing the raw
 * inner XML of <w:ftr>.  Wave A.6/B.x will consolidate via a shared body
 * parser once document.ts (A.3) ships the paragraph/run/table logic.
 *
 * The `id` parameter is the relationship-derived file identifier (e.g. "rId2")
 * that the orchestrator must supply; the footer XML itself carries no self-id.
 */

import { XMLParser } from 'fast-xml-parser'

import { DocxParseError } from './unzip'
import type { Footer, UnknownNode, Block } from '../model/document'

// ---------------------------------------------------------------------------
// Parser instance — same config as Wave A.1
// ---------------------------------------------------------------------------

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

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
  try {
    xmlParser.parse(xml)
  } catch (cause) {
    const msg = cause instanceof Error ? cause.message : String(cause)
    throw new DocxParseError(`Failed to parse footer XML: ${msg}`, id)
  }

  const inner: UnknownNode = { kind: 'unknown', xml }
  const blocks: ReadonlyArray<Block> = [inner]

  return { kind: 'footer', id, blocks }
}
