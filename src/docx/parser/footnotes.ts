/**
 * Atlas — DOCX footnotes parser (Wave A.5, hardened in P1.3)
 *
 * Parses `word/footnotes.xml`.  Each `<w:footnote w:id="...">` becomes a
 * `Footnote` entry in the returned map (keyed by id string), whose body is
 * now real `Paragraph[]` blocks — ported from `comments.ts`'s
 * synthetic-wrapper technique (see `partBody.ts`) — replacing the earlier
 * "Option B" stub that stored each note's body as a single opaque
 * `UnknownNode` and made the serializer throw on every save that included a
 * footnote (DXP-03/DXS-01).
 *
 * Special separator types (separator, continuationSeparator, continuationNotice)
 * ARE included in the map but their `noteType` field is set accordingly so
 * downstream consumers can filter them out.
 */

import { XMLParser } from 'fast-xml-parser'

import { parseParagraphsFromRawNodes } from './partBody'
import { DocxParseError } from './unzip'
import { assertXmlPartSizeWithinLimit } from './xmlSizeGuard'
import type { Footnote, NoteType } from '../model/document'

// ---------------------------------------------------------------------------
// Parser instance — same config as Wave A.1
// ---------------------------------------------------------------------------

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

// ---------------------------------------------------------------------------
// Internal XML shapes
// ---------------------------------------------------------------------------

interface RawFootnote {
  '@_w:id'?: string | number
  '@_w:type'?: string
  'w:p'?: unknown
  [key: string]: unknown
}

interface RawFootnotesDoc {
  'w:footnotes'?: {
    'w:footnote'?: RawFootnote | RawFootnote[]
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
 * Parse the XML content of `word/footnotes.xml`.
 *
 * @param xml - UTF-8 text of the footnotes part.
 * @returns   Immutable map: footnote id → Footnote node (includes separators).
 * @throws    DocxParseError on malformed XML or a footnote missing w:id.
 */
export function parseFootnotes(xml: string): ReadonlyMap<string, Footnote> {
  assertXmlPartSizeWithinLimit(xml, 'word/footnotes.xml')
  let parsed: RawFootnotesDoc
  try {
    parsed = xmlParser.parse(xml) as RawFootnotesDoc
  } catch (cause) {
    const msg = cause instanceof Error ? cause.message : String(cause)
    throw new DocxParseError(`Failed to parse footnotes XML: ${msg}`)
  }

  const rawNotes = parsed?.['w:footnotes']?.['w:footnote']
  if (rawNotes === undefined || rawNotes === null) {
    return new Map()
  }

  const items: RawFootnote[] = Array.isArray(rawNotes) ? rawNotes : [rawNotes]
  const result = new Map<string, Footnote>()

  for (const item of items) {
    const rawId = item['@_w:id']
    if (rawId === undefined || rawId === null) {
      throw new DocxParseError('Footnote element is missing required w:id attribute')
    }
    const id = String(rawId)
    const noteType = toNoteType(item['@_w:type'])
    const blocks = parseParagraphsFromRawNodes(item['w:p'])

    const footnote: Footnote = {
      kind: 'footnote',
      id,
      ...(noteType !== undefined ? { noteType } : {}),
      blocks,
    }
    result.set(id, footnote)
  }

  return result
}
