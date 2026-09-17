/**
 * Slide text-run collection ([MS-PPT] 2.4.14 TextHeaderAtom, 2.9.113
 * TextCharsAtom, 2.9.114 TextBytesAtom).
 *
 * Each text-bearing shape's content lives in a `TextHeaderAtom` (declares
 * what kind of text it is — title, body, other) immediately followed by one
 * `TextCharsAtom` (UTF-16) or `TextBytesAtom` (cp1252) sibling holding the
 * actual characters. Resolving the full OfficeArt shape tree to know exactly
 * *which* shape each pair belongs to is out of scope for a text-only
 * preview (see `slides.ts`'s header) — pairing them up by simple
 * document-order adjacency is sufficient here.
 */
import type { BinaryReader } from '../binaryReader'
import { decodeWindows1252 } from '../../utils/textDecoding'
import { walkPptRecords, type PptRecordHeader } from './records'

const RT_TEXT_HEADER_ATOM = 3999
const RT_TEXT_CHARS_ATOM = 4000
const RT_TEXT_BYTES_ATOM = 4008

const UTF16LE_DECODER = new TextDecoder('utf-16le')

export interface PptTextGroup {
  /** TextHeaderAtom.textType (0/6 = title/center-title); `undefined` when no header preceded this run. */
  readonly textType: number | undefined
  readonly text: string
}

function decodeCharsAtom(reader: BinaryReader, header: PptRecordHeader): string {
  const length = header.dataEnd - header.dataStart
  // A well-formed TextCharsAtom's length is always even (UTF-16 code
  // units); an odd length is corrupt — drop the dangling trailing byte
  // rather than fail the whole slide over one malformed atom.
  const evenLength = length - (length % 2)
  return UTF16LE_DECODER.decode(reader.subarray(header.dataStart, evenLength))
}

function decodeBytesAtom(reader: BinaryReader, header: PptRecordHeader): string {
  return decodeWindows1252(reader.subarray(header.dataStart, header.dataEnd - header.dataStart))
}

/**
 * Collects every text run within one Slide container's record subtree
 * (`[start, end)`), in document order, pairing each `TextCharsAtom`/
 * `TextBytesAtom` with whichever `TextHeaderAtom` most recently preceded it.
 */
export function collectSlideTextGroups(reader: BinaryReader, start: number, end: number): ReadonlyArray<PptTextGroup> {
  const groups: PptTextGroup[] = []
  let pendingTextType: number | undefined

  walkPptRecords(reader, start, end, (header) => {
    if (header.type === RT_TEXT_HEADER_ATOM) {
      pendingTextType = header.dataEnd - header.dataStart >= 4 ? reader.i32(header.dataStart) : undefined
      return
    }
    if (header.type === RT_TEXT_CHARS_ATOM) {
      groups.push({ textType: pendingTextType, text: decodeCharsAtom(reader, header) })
      pendingTextType = undefined
      return
    }
    if (header.type === RT_TEXT_BYTES_ATOM) {
      groups.push({ textType: pendingTextType, text: decodeBytesAtom(reader, header) })
      pendingTextType = undefined
    }
  })

  return groups
}
