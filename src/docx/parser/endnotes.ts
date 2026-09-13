/**
 * Atlas — DOCX endnotes parser (Wave A.5, hardened in P1.3)
 *
 * Parses `word/endnotes.xml`.  Each `<w:endnote w:id="...">` becomes an
 * `Endnote` entry in the returned map (keyed by id string), whose body is
 * now real `Paragraph[]` blocks — ported from `comments.ts`'s
 * synthetic-wrapper technique (see `partBody.ts`) — replacing the earlier
 * "Option B" stub that stored each note's body as a single opaque
 * `UnknownNode` and made the serializer throw on every save that included an
 * endnote (DXP-03/DXS-01).
 *
 * Special separator types are included with their noteType set accordingly.
 */

import { XMLParser } from 'fast-xml-parser'

import { parseParagraphsFromRawNodes } from './partBody'
import { DocxParseError } from './unzip'
import { assertXmlPartSizeWithinLimit } from './xmlSizeGuard'
import type { Endnote, NoteType } from '../model/document'

// ---------------------------------------------------------------------------
// Parser instance — same config as Wave A.1
// ---------------------------------------------------------------------------

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

// ---------------------------------------------------------------------------
// Internal XML shapes
// ---------------------------------------------------------------------------

interface RawEndnote {
  '@_w:id'?: string | number
  '@_w:type'?: string
  'w:p'?: unknown
  [key: string]: unknown
}

interface RawEndnotesDoc {
  'w:endnotes'?: {
    'w:endnote'?: RawEndnote | RawEndnote[]
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function toNoteType(raw: string | undefined): NoteType | undefined {
  if (raw === 'separator') return 'separator'
  if (raw === 'continuationSeparator') return 'continuationSeparator'
  if (raw === 'continuationNotice') return 'continuationNotice'
  if (raw === 'normal') return 'normal'
  return undefined
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Parse the XML content of `word/endnotes.xml`.
 *
 * @param xml - UTF-8 text of the endnotes part.
 * @returns   Immutable map: endnote id → Endnote node (includes separators).
 * @throws    DocxParseError on malformed XML or an endnote missing w:id.
 */
export function parseEndnotes(xml: string): ReadonlyMap<string, Endnote> {
  assertXmlPartSizeWithinLimit(xml, 'word/endnotes.xml')
  let parsed: RawEndnotesDoc
  try {
    parsed = xmlParser.parse(xml) as RawEndnotesDoc
  } catch (cause) {
    const msg = cause instanceof Error ? cause.message : String(cause)
    throw new DocxParseError(`Failed to parse endnotes XML: ${msg}`)
  }

  const rawNotes = parsed?.['w:endnotes']?.['w:endnote']
  if (rawNotes === undefined || rawNotes === null) {
    return new Map()
  }

  const items: RawEndnote[] = Array.isArray(rawNotes) ? rawNotes : [rawNotes]
  const result = new Map<string, Endnote>()

  for (const item of items) {
    const rawId = item['@_w:id']
    if (rawId === undefined || rawId === null) {
      throw new DocxParseError('Endnote element is missing required w:id attribute')
    }
    const id = String(rawId)
    const noteType = toNoteType(item['@_w:type'])
    const blocks = parseParagraphsFromRawNodes(item['w:p'])

    const endnote: Endnote = {
      kind: 'endnote',
      id,
      ...(noteType !== undefined ? { noteType } : {}),
      blocks,
    }
    result.set(id, endnote)
  }

  return result
}
