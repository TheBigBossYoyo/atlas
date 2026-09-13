/**
 * Atlas — DOCX endnotes parser (Wave A.5)
 *
 * Parses `word/endnotes.xml`.  Each `<w:endnote w:id="...">` becomes an
 * `Endnote` entry in the returned map (keyed by id string).
 *
 * OPTION B: body content is stored as a single UnknownNode; Wave A.6/B.x
 * will expand it using the shared body parser from document.ts (A.3).
 *
 * Special separator types are included with their noteType set accordingly.
 */

import { XMLParser } from 'fast-xml-parser'

import { DocxParseError } from './unzip'
import type { Endnote, UnknownNode, Block, NoteType } from '../model/document'

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

    const inner: UnknownNode = { kind: 'unknown', xml: JSON.stringify(item) }
    const blocks: ReadonlyArray<Block> = [inner]

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
