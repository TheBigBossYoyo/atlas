/**
 * .doc (Word 97-2003) text extraction orchestrator (wave-4 legacy-office).
 *
 * Ties together the CFB container, FIB, piece table, and text/paragraph
 * decoders into one entry point. Deliberately produces a plain
 * `{ paragraphs: string[] }` structure rather than mapping into the app's
 * full DOCX model (`docx/model`) — that model's ~700 lines of run/paragraph
 * property types describe OOXML formatting concepts (styles, numbering,
 * sections, ...) a binary .doc's piece table doesn't expose at all without a
 * much larger reverse-engineering effort (CHPX/PAPX formatted-disk-page
 * parsing). A flat paragraph list is the honest "best-effort, text and
 * structure only" result this wave asks for.
 *
 * What this covers: paragraph breaks, table cell/row marks, line/page
 * breaks (as embedded soft breaks), field codes collapsed to their cached
 * result text, and both compressed (cp1252) and Unicode (UTF-16LE) text
 * pieces.
 *
 * What this does NOT cover (all out of scope for a read-only text preview):
 * character/paragraph formatting (bold, fonts, alignment, ...), tables as
 * structured grids (cell marks just become paragraph breaks), images,
 * headers/footers/footnotes/endnotes/comments (only the main document story
 * — `ccpText` — is extracted), hyperlinks, and any field's *instruction*
 * text (only its cached result survives, see `paragraphs.ts`).
 */
import { findEntry, readCfb, type CfbEntry } from '../cfb'
import { BinaryReader } from '../binaryReader'
import { LegacyFormatError } from '../errors'
import { parseFib } from './fib'
import { parseClx } from './pieceTable'
import { decodeStoryText } from './textDecode'
import { splitIntoParagraphs } from './paragraphs'

export interface LegacyDocResult {
  readonly paragraphs: ReadonlyArray<string>
}

function requireStream(entry: CfbEntry | null, label: string): Uint8Array {
  if (!entry || entry.type !== 'stream' || !entry.content) {
    throw new LegacyFormatError(`This file doesn't look like a valid Word 97-2003 document: missing the "${label}" stream.`)
  }
  return entry.content
}

/** Extracts the main document story's text from a legacy Word 97-2003 (.doc) file, as a flat paragraph list. */
export function extractLegacyDocText(bytes: Uint8Array): LegacyDocResult {
  const cfb = readCfb(bytes)
  const wordDocument = requireStream(findEntry(cfb, 'WordDocument'), 'WordDocument')

  const fib = parseFib(wordDocument)
  const tableStream = requireStream(findEntry(cfb, fib.tableStreamName), fib.tableStreamName)

  const pieces = parseClx(tableStream, fib.fcClx, fib.lcbClx)
  const rawText = decodeStoryText(new BinaryReader(wordDocument), pieces, fib.ccpText)

  return { paragraphs: splitIntoParagraphs(rawText) }
}
