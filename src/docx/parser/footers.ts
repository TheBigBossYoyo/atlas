/**
 * Atlas — DOCX footer parser (Wave A.5, hardened in P1.3, extended for
 * table support in the wave 1 follow-up)
 *
 * Parses the body content of a `word/footer*.xml` part into real `Block[]`
 * blocks — paragraphs AND tables — by extracting the `<w:ftr>` wrapper's
 * exact inner XML and re-parsing it via `partBody.ts`'s shared synthetic-
 * wrapper technique, replacing the earlier "Option B" stub that stored the
 * whole part as a single opaque `UnknownNode` and made the serializer throw
 * on every save that included a footer (DXP-04/DXS-01), and later a
 * paragraph-only extraction that silently dropped any table a footer
 * contained instead of throwing.
 *
 * The `id` parameter is the relationship-derived file identifier (e.g. "rId2")
 * that the orchestrator must supply; the footer XML itself carries no self-id.
 */

import { XMLParser } from 'fast-xml-parser'

import { extractSingleElementInnerXml, parseBlocksFromXmlFragment } from './partBody'
import { DocxParseError } from './unzip'
import { assertXmlPartSizeWithinLimit } from './xmlSizeGuard'
import type { Footer } from '../model/document'

// Used only to validate well-formedness up front (with the same error
// message/behaviour this module always had) — the actual block content
// comes from partBody.ts's substring-based, order-preserving extraction.
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
  assertXmlPartSizeWithinLimit(xml, id === '' ? 'word/footer*.xml' : `word/footer*.xml (${id})`)

  try {
    xmlParser.parse(xml)
  } catch (cause) {
    const msg = cause instanceof Error ? cause.message : String(cause)
    throw new DocxParseError(`Failed to parse footer XML: ${msg}`, id)
  }

  const innerXml = extractSingleElementInnerXml(xml, 'w:ftr')
  const blocks = parseBlocksFromXmlFragment(innerXml)

  return { kind: 'footer', id, blocks }
}
