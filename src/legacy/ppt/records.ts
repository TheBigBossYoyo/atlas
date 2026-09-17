/**
 * PPT binary record-tree walker ([MS-PPT] 2.3.1 RecordHeader).
 *
 * A .ppt's "PowerPoint Document" stream is one big tree of records: each
 * starts with an 8-byte header (2-byte verInstance, 2-byte recType, 4-byte
 * recLen), and a record whose low 4 bits of verInstance (`recVer`) equal
 * 0xF is a *container* whose `recLen` bytes are themselves a nested run of
 * records — every other `recVer` marks a leaf "atom" holding raw data.
 */
import type { BinaryReader } from '../binaryReader'

export interface PptRecordHeader {
  readonly recVer: number
  readonly recInstance: number
  readonly type: number
  /** Byte offset of this record's payload (just past its 8-byte header). */
  readonly dataStart: number
  /** Exclusive end of this record's payload. */
  readonly dataEnd: number
}

const CONTAINER_REC_VER = 0x0f
const RECORD_HEADER_SIZE = 8

// Guards against a maliciously/corruptly self-referential record tree —
// real PPT shape/text nesting is at most a handful of levels deep.
const MAX_RECORD_DEPTH = 64

/**
 * Depth-first, pre-order walk over a run of PPT records in `[start, end)`.
 * `onRecord` is called for every record (container or atom) in document
 * order — a container's descendants are all visited before its next
 * sibling, so a caller pairing up adjacent sibling atoms (see `text.ts`)
 * sees them in the same order Word/PowerPoint originally wrote them.
 *
 * A record whose declared length would run past `end` stops the walk at
 * that level rather than throwing — one corrupt record loses only its own
 * (and any later siblings') subtree, not the whole presentation, mirroring
 * the ODP/PPTX parsers' per-slide resilience one level up (see
 * `slides.ts`'s per-slide `try/catch`).
 */
export function walkPptRecords(
  reader: BinaryReader,
  start: number,
  end: number,
  onRecord: (header: PptRecordHeader) => void,
  depth: number = 0,
): void {
  if (depth > MAX_RECORD_DEPTH) {
    return
  }

  let offset = start
  while (offset + RECORD_HEADER_SIZE <= end) {
    const verInstance = reader.u16(offset)
    const recVer = verInstance & 0x000f
    const type = reader.u16(offset + 2)
    const length = reader.u32(offset + 4)
    const dataStart = offset + RECORD_HEADER_SIZE
    const dataEnd = dataStart + length

    if (dataEnd > end) {
      break
    }

    onRecord({ recVer, recInstance: (verInstance >> 4) & 0x0fff, type, dataStart, dataEnd })

    if (recVer === CONTAINER_REC_VER) {
      walkPptRecords(reader, dataStart, dataEnd, onRecord, depth + 1)
    }

    offset = dataEnd
  }
}
