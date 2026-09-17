/**
 * Bounds-checked little-endian byte cursor shared by the .doc FIB/piece-table
 * reader and the .ppt record-tree walker (wave-4 legacy-office).
 *
 * Both formats are full of file-provided offsets and lengths (piece byte
 * offsets, record lengths, CLX sizes, ...) that a parser must trust just
 * enough to follow, but never enough to dereference blindly — a crafted or
 * merely corrupt file can claim any offset/length it likes. Every accessor
 * here validates the requested range against the buffer's real length
 * first, throwing a `LegacyFormatError` instead of letting a bad offset
 * produce a raw `RangeError` (DataView) or silent `undefined` (typed array
 * index) somewhere downstream that's much harder to trace back to "this
 * file is corrupt."
 */
import { LegacyFormatError } from './errors'

export class BinaryReader {
  readonly length: number
  private readonly bytes: Uint8Array
  private readonly view: DataView

  constructor(bytes: Uint8Array) {
    this.bytes = bytes
    this.length = bytes.length
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  }

  private assertRange(offset: number, size: number, what: string): void {
    if (offset < 0 || size < 0 || offset + size > this.length) {
      throw new LegacyFormatError(
        `Corrupt document: expected ${what} at byte offset ${offset.toLocaleString()}, but the stream is only ` +
          `${this.length.toLocaleString()} bytes long.`,
      )
    }
  }

  u8(offset: number): number {
    this.assertRange(offset, 1, 'a byte')
    return this.view.getUint8(offset)
  }

  u16(offset: number): number {
    this.assertRange(offset, 2, 'a 16-bit value')
    return this.view.getUint16(offset, true)
  }

  i32(offset: number): number {
    this.assertRange(offset, 4, 'a 32-bit value')
    return this.view.getInt32(offset, true)
  }

  u32(offset: number): number {
    this.assertRange(offset, 4, 'a 32-bit value')
    return this.view.getUint32(offset, true)
  }

  /** A read-only view into the underlying bytes (never copies) — bounds-checked the same as every numeric read. */
  subarray(offset: number, length: number): Uint8Array {
    this.assertRange(offset, length, `a ${length.toLocaleString()}-byte block`)
    return this.bytes.subarray(offset, offset + length)
  }
}
