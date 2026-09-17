/**
 * FIB (File Information Block) reader — the fixed-layout header at the start
 * of a .doc's "WordDocument" stream ([MS-DOC] 2.5.1).
 *
 * Only the handful of fields needed to locate the piece table are read here:
 * which Table stream holds it, where the CLX sits within that stream, and
 * how many characters belong to the main document story (as opposed to
 * footnotes/headers/annotations, which share the same CP space further
 * along the same piece table). Everything else in the FIB (fonts, styles,
 * bookmarks, ...) is out of scope for a text-only preview.
 */
import { BinaryReader } from '../binaryReader'
import { LegacyFormatError } from '../errors'

const FIB_MAGIC = 0xa5ec

// Fixed byte offsets into the FIB. These are stable across every
// Word-97-through-2003 file (nFib 0x00C1 through 0x0112) — only fields far
// beyond what's read here vary by version:
//   FibBase (32 bytes) -> csw (2) -> fibRgW97 (28) -> cslw (2) ->
//   fibRgLw97 (88, `ccpText` is its 4th 4-byte field) -> cbRgFcLcb (2) ->
//   fibRgFcLcbBlob (`fcClx`/`lcbClx` is pair #33 of that 8-byte-per-pair blob).
const OFFSET_WIDENT = 0
const OFFSET_FLAGS1 = 0x0a // bit 0x0200 = fWhichTblStm
const OFFSET_CCP_TEXT = 76 // 32 + 2 + 28 + 2 + (3 * 4)
const OFFSET_FC_CLX = 418 // 32 + 2 + 28 + 2 + 88 + 2 + (33 * 8)
const OFFSET_LCB_CLX = 422

const F_WHICH_TABLE_STREAM = 0x0200

export interface FibInfo {
  /** [MS-DOC] FibBase.fWhichTblStm selects which of the two Table streams holds the CLX/piece table. */
  readonly tableStreamName: '0Table' | '1Table'
  /** Byte offset of the CLX within the Table stream. */
  readonly fcClx: number
  /** Byte length of the CLX within the Table stream. */
  readonly lcbClx: number
  /** Character count of the main document story — the piece table also covers footnotes/headers/etc. appended after it in the same CP space. */
  readonly ccpText: number
}

/** Reads the FIB fields needed to locate and bound the piece table. */
export function parseFib(wordDocument: Uint8Array): FibInfo {
  const reader = new BinaryReader(wordDocument)

  const magic = reader.u16(OFFSET_WIDENT)
  if (magic !== FIB_MAGIC) {
    throw new LegacyFormatError(
      `This doesn't look like a Word 97-2003 document (bad FIB signature: expected 0x${FIB_MAGIC.toString(16)}, ` +
        `found 0x${magic.toString(16)}).`,
    )
  }

  const flags1 = reader.u16(OFFSET_FLAGS1)
  const tableStreamName = (flags1 & F_WHICH_TABLE_STREAM) !== 0 ? '1Table' : '0Table'

  const ccpText = reader.u32(OFFSET_CCP_TEXT)
  const fcClx = reader.u32(OFFSET_FC_CLX)
  const lcbClx = reader.u32(OFFSET_LCB_CLX)

  return { tableStreamName, fcClx, lcbClx, ccpText }
}
