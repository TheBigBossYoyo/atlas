/**
 * Atlas — DOCX header parser (Wave A.5, hardened in P1.3, extended for
 * table support in the wave 1 follow-up)
 *
 * Parses the body content of a `word/header*.xml` part into real `Block[]`
 * blocks — paragraphs AND tables — by extracting the `<w:hdr>` wrapper's
 * exact inner XML and re-parsing it via `partBody.ts`'s shared synthetic-
 * wrapper technique, replacing the earlier "Option B" stub that stored the
 * whole part as a single opaque `UnknownNode` and made the serializer throw
 * on every save that included a header (DXP-04/DXS-01), and later a
 * paragraph-only extraction that silently dropped any table a header
 * contained (e.g. a letterhead) instead of throwing.
 *
 * The `id` parameter is the relationship-derived file identifier (e.g. "rId1")
 * that the orchestrator must supply; the header XML itself carries no self-id.
 */

import { XMLParser } from 'fast-xml-parser'

import { extractSingleElementInnerXml, parseBlocksFromXmlFragment } from './partBody'
import { DocxParseError } from './unzip'
import { assertXmlPartSizeWithinLimit } from './xmlSizeGuard'
import type { Header } from '../model/document'

// Used only to validate well-formedness up front (with the same error
// message/behaviour this module always had) — the actual block content
// comes from partBody.ts's substring-based, order-preserving extraction.
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
  assertXmlPartSizeWithinLimit(xml, id === '' ? 'word/header*.xml' : `word/header*.xml (${id})`)

  try {
    xmlParser.parse(xml)
  } catch (cause) {
    const msg = cause instanceof Error ? cause.message : String(cause)
    throw new DocxParseError(`Failed to parse header XML: ${msg}`, id)
  }

  const innerXml = extractSingleElementInnerXml(xml, 'w:hdr')
  const blocks = parseBlocksFromXmlFragment(innerXml)

  return { kind: 'header', id, blocks }
}
