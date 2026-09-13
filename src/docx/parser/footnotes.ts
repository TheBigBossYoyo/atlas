/**
 * Atlas — DOCX footnotes parser (Wave A.5)
 *
 * Parses `word/footnotes.xml`.  Each `<w:footnote w:id="...">` becomes a
 * `Footnote` entry in the returned map (keyed by id string).
 *
 * OPTION B: body content is stored as a single UnknownNode; Wave A.6/B.x
 * will expand it using the shared body parser from document.ts (A.3).
 *
 * Special separator types (separator, continuationSeparator, continuationNotice)
 * ARE included in the map but their `noteType` field is set accordingly so
 * downstream consumers can filter them out.
 */

import { XMLParser } from 'fast-xml-parser'

import { DocxParseError } from './unzip'
import type { Footnote, UnknownNode, Block, NoteType } from '../model/document'

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

function rawXmlForNote(parsed: RawFootnote): string {
  // Serialize back to a minimal XML string so the UnknownNode carries enough
  // context for Option B consumers.  We use JSON as a lossless stand-in; a
  // real serializer will replace this in Wave A.6/B.x.
  return JSON.stringify(parsed)
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

    const inner: UnknownNode = { kind: 'unknown', xml: rawXmlForNote(item) }
    const blocks: ReadonlyArray<Block> = [inner]

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
