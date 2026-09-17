/**
 * Decodes a .doc's main-document-story text from its pieces (see
 * `pieceTable.ts`) into one raw string — still containing Word's in-band
 * special characters (paragraph marks, field markers, ...), which
 * `paragraphs.ts` interprets afterward.
 */
import type { BinaryReader } from '../binaryReader'
import { LegacyFormatError } from '../errors'
import { decodeWindows1252 } from '../../utils/textDecoding'
import type { Piece } from './pieceTable'

const UTF16LE_DECODER = new TextDecoder('utf-16le')

/**
 * Whole-document safety cap independent of any single piece's own size —
 * bounds the total characters this module will ever decode/concatenate, so
 * a corrupt `ccpText` claiming an enormous story can't turn a text
 * extraction into unbounded work (keeps a large .doc's parse linear and
 * capped, per this wave's performance requirement).
 */
export const MAX_STORY_CHARACTERS = 20_000_000

/**
 * Concatenates every piece's decoded text up to `ccpLimit` characters (the
 * FIB's `ccpText` — the main story's length; later pieces belong to
 * footnotes/headers/etc. sharing the same CP space and are not included).
 * Builds the result via an array + single `join` rather than repeated `+=`,
 * so this stays linear instead of quadratic on a large document.
 */
export function decodeStoryText(wordDocument: BinaryReader, pieces: ReadonlyArray<Piece>, ccpLimit: number): string {
  const parts: string[] = []
  let charsDecoded = 0

  for (const piece of pieces) {
    if (piece.cpStart >= ccpLimit) {
      break // pieces are laid out in CP order; nothing after this belongs to the main story.
    }

    const pieceEnd = Math.min(piece.cpEnd, ccpLimit)
    const numChars = pieceEnd - piece.cpStart
    if (numChars <= 0) {
      continue
    }

    if (charsDecoded + numChars > MAX_STORY_CHARACTERS) {
      throw new LegacyFormatError(
        `Corrupt or unsupported document: main text exceeds the ${MAX_STORY_CHARACTERS.toLocaleString()}-character safety limit.`,
      )
    }

    const byteLength = piece.isCompressed ? numChars : numChars * 2
    const bytes = wordDocument.subarray(piece.byteOffset, byteLength)

    parts.push(piece.isCompressed ? decodeWindows1252(bytes) : UTF16LE_DECODER.decode(bytes))
    charsDecoded += numChars
  }

  return parts.join('')
}
