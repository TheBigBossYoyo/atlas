/**
 * CLX / piece table (PlcPcd) reader ([MS-DOC] 2.9.38, 2.8.35, 2.9.177).
 *
 * A .doc's text is stored as a sequence of "pieces," each a contiguous run
 * of either compressed (cp1252, 1 byte/char) or Unicode (UTF-16LE, 2
 * bytes/char) text somewhere in the WordDocument stream. The CLX (found via
 * the FIB's `fcClx`/`lcbClx`, inside the Table stream) is a small
 * self-describing block sequence that ends in exactly one piece table
 * (Pcdt) describing where every piece lives and how it's encoded.
 */
import { BinaryReader } from '../binaryReader'
import { LegacyFormatError } from '../errors'

export interface Piece {
  /** Character-position range this piece covers, exclusive of `cpEnd`. */
  readonly cpStart: number
  readonly cpEnd: number
  readonly isCompressed: boolean
  /** Byte offset into the WordDocument stream where this piece's text begins. */
  readonly byteOffset: number
}

const CLXT_PROPERTY_RUN = 0x01 // "Prc" block — paragraph/character formatting sprms; not needed for text.
const CLXT_PIECE_TABLE = 0x02 // "Pcdt" block — the piece table itself.

const PCD_SIZE = 8 // 2 bytes flags (unused here) + 4 bytes fc + 2 bytes prm (unused here).
const FC_COMPRESSED_FLAG = 0x40000000 // PCD.fc bit 30 — [MS-DOC] 2.9.73 FcCompressed.

// A hostile/corrupt `lcbClx` could otherwise claim an enormous piece count;
// every real piece needs at least 12 bytes (4 for its CP, 8 for its PCD), so
// this ceiling is already far beyond any legitimate document.
const MAX_PIECES = 1_000_000

/**
 * Walks the CLX's block sequence to find and parse its piece table.
 * `tableStream` is the raw bytes of whichever Table stream the FIB pointed
 * at (`0Table`/`1Table`); `fcClx`/`lcbClx` bound the CLX within it.
 */
export function parseClx(tableStream: Uint8Array, fcClx: number, lcbClx: number): ReadonlyArray<Piece> {
  const reader = new BinaryReader(tableStream)

  if (lcbClx <= 0) {
    throw new LegacyFormatError('Corrupt document: empty piece table (CLX).')
  }
  // Bounds-check the whole CLX region up front so every read below is
  // guaranteed to stay inside it (a corrupt fcClx/lcbClx fails here, once,
  // with one clear message, rather than at some arbitrary later offset).
  reader.subarray(fcClx, lcbClx)

  const clxEnd = fcClx + lcbClx
  let offset = fcClx

  while (offset < clxEnd) {
    const clxt = reader.u8(offset)
    offset += 1

    if (clxt === CLXT_PROPERTY_RUN) {
      const cbGrpprl = reader.u16(offset)
      offset += 2 + cbGrpprl
      continue
    }

    if (clxt === CLXT_PIECE_TABLE) {
      const lcb = reader.u32(offset)
      return parsePlcPcd(reader, offset + 4, lcb)
    }

    throw new LegacyFormatError(`Corrupt document: unrecognized piece-table block type 0x${clxt.toString(16)}.`)
  }

  throw new LegacyFormatError('Corrupt document: piece table (CLX) has no piece-table block.')
}

function parsePlcPcd(reader: BinaryReader, start: number, lcb: number): ReadonlyArray<Piece> {
  if (lcb < 4 || (lcb - 4) % 12 !== 0) {
    throw new LegacyFormatError('Corrupt document: malformed piece table size.')
  }

  const pieceCount = (lcb - 4) / 12
  if (pieceCount > MAX_PIECES) {
    throw new LegacyFormatError(
      `Corrupt document: piece table claims ${pieceCount.toLocaleString()} pieces, over the safety limit.`,
    )
  }

  // Bounds-check the whole PlcPcd region (CP array + PCD array) up front.
  reader.subarray(start, lcb)

  const cpArrayStart = start
  const pcdArrayStart = start + (pieceCount + 1) * 4

  const pieces: Piece[] = []
  for (let i = 0; i < pieceCount; i += 1) {
    const cpStart = reader.u32(cpArrayStart + i * 4)
    const cpEnd = reader.u32(cpArrayStart + (i + 1) * 4)

    const pcdOffset = pcdArrayStart + i * PCD_SIZE
    const fcRaw = reader.u32(pcdOffset + 2) // skip the 2-byte flags field.
    const isCompressed = (fcRaw & FC_COMPRESSED_FLAG) !== 0
    // Clearing bit 30 and forcing an unsigned result guards against a
    // hostile fc value with bit 31 set turning this into a negative offset.
    const fc = (fcRaw & ~FC_COMPRESSED_FLAG) >>> 0
    const byteOffset = isCompressed ? fc >>> 1 : fc

    pieces.push({ cpStart, cpEnd, isCompressed, byteOffset })
  }

  return pieces
}
