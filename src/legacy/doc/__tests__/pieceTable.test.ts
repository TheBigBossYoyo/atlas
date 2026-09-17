import { describe, expect, it } from 'vitest'

import { parseClx } from '../pieceTable'
import { LegacyFormatError } from '../../errors'

const FC_COMPRESSED_FLAG = 0x40000000

type PieceSpec = { readonly cpStart: number; readonly cpEnd: number; readonly isCompressed: boolean; readonly byteOffset: number }

/** Builds a byte layout `clxt(1) [Prc blocks...] 0x02 lcb(4) PlcPcd` at `fcClx`, inside a buffer `bufferSize` bytes long. */
function buildClxBytes(
  fcClx: number,
  pieces: ReadonlyArray<PieceSpec>,
  options: { readonly bufferSize?: number; readonly leadingPrc?: ReadonlyArray<number> } = {},
): { readonly bytes: Uint8Array; readonly lcbClx: number } {
  const cpCount = pieces.length + 1
  const plcPcdSize = cpCount * 4 + pieces.length * 8
  const pcdtBlockSize = 1 + 4 + plcPcdSize // clxt byte + lcb(4) + PlcPcd

  const leadingPrc = options.leadingPrc ?? []
  const leadingSize = leadingPrc.length > 0 ? 1 + 2 + leadingPrc.length : 0 // clxt(1) + cbGrpprl(2) + payload

  const lcbClx = leadingSize + pcdtBlockSize
  const bufferSize = options.bufferSize ?? fcClx + lcbClx + 16
  const buffer = new ArrayBuffer(bufferSize)
  const view = new DataView(buffer)

  let offset = fcClx

  if (leadingPrc.length > 0) {
    view.setUint8(offset, 0x01)
    offset += 1
    view.setUint16(offset, leadingPrc.length, true)
    offset += 2
    new Uint8Array(buffer, offset, leadingPrc.length).set(leadingPrc)
    offset += leadingPrc.length
  }

  view.setUint8(offset, 0x02)
  offset += 1
  view.setUint32(offset, plcPcdSize, true)
  offset += 4

  const cpArrayStart = offset
  for (let i = 0; i < pieces.length; i += 1) {
    view.setUint32(cpArrayStart + i * 4, pieces[i].cpStart, true)
  }
  view.setUint32(cpArrayStart + pieces.length * 4, pieces[pieces.length - 1]?.cpEnd ?? 0, true)

  const pcdArrayStart = cpArrayStart + cpCount * 4
  pieces.forEach((piece, i) => {
    const pcdOffset = pcdArrayStart + i * 8
    view.setUint16(pcdOffset, 0, true) // flags — unused by the reader
    const fc = piece.isCompressed ? (piece.byteOffset * 2) | FC_COMPRESSED_FLAG : piece.byteOffset
    view.setUint32(pcdOffset + 2, fc, true)
    view.setUint16(pcdOffset + 6, 0, true) // prm — unused by the reader
  })

  return { bytes: new Uint8Array(buffer), lcbClx }
}

describe('parseClx', () => {
  it('parses a single compressed piece', () => {
    const fcClx = 0
    const { bytes, lcbClx } = buildClxBytes(fcClx, [{ cpStart: 0, cpEnd: 10, isCompressed: true, byteOffset: 500 }])

    const pieces = parseClx(bytes, fcClx, lcbClx)

    expect(pieces).toHaveLength(1)
    expect(pieces[0]).toEqual({ cpStart: 0, cpEnd: 10, isCompressed: true, byteOffset: 500 })
  })

  it('parses a single uncompressed (Unicode) piece', () => {
    const fcClx = 0
    const { bytes, lcbClx } = buildClxBytes(fcClx, [{ cpStart: 0, cpEnd: 5, isCompressed: false, byteOffset: 1000 }])

    const pieces = parseClx(bytes, fcClx, lcbClx)

    expect(pieces[0]).toEqual({ cpStart: 0, cpEnd: 5, isCompressed: false, byteOffset: 1000 })
  })

  it('parses multiple pieces with mixed compression, preserving CP order', () => {
    const fcClx = 0
    const { bytes, lcbClx } = buildClxBytes(fcClx, [
      { cpStart: 0, cpEnd: 10, isCompressed: true, byteOffset: 0 },
      { cpStart: 10, cpEnd: 20, isCompressed: false, byteOffset: 200 },
      { cpStart: 20, cpEnd: 25, isCompressed: true, byteOffset: 50 },
    ])

    const pieces = parseClx(bytes, fcClx, lcbClx)

    expect(pieces).toHaveLength(3)
    expect(pieces.map((p) => [p.cpStart, p.cpEnd])).toEqual([
      [0, 10],
      [10, 20],
      [20, 25],
    ])
    expect(pieces.map((p) => p.isCompressed)).toEqual([true, false, true])
  })

  it('skips a leading property-run (Prc) block before the piece table', () => {
    const fcClx = 0
    const { bytes, lcbClx } = buildClxBytes(fcClx, [{ cpStart: 0, cpEnd: 3, isCompressed: true, byteOffset: 8 }], {
      leadingPrc: [0xaa, 0xbb, 0xcc, 0xdd],
    })

    const pieces = parseClx(bytes, fcClx, lcbClx)
    expect(pieces).toHaveLength(1)
    expect(pieces[0].byteOffset).toBe(8)
  })

  it('honors a non-zero fcClx offset (CLX not at the start of the Table stream)', () => {
    const fcClx = 128
    const { bytes, lcbClx } = buildClxBytes(fcClx, [{ cpStart: 0, cpEnd: 4, isCompressed: false, byteOffset: 16 }])

    const pieces = parseClx(bytes, fcClx, lcbClx)
    expect(pieces[0].cpEnd).toBe(4)
  })

  it('clears the fCompressed flag bit rather than leaving it in the byte offset', () => {
    const fcClx = 0
    const { bytes, lcbClx } = buildClxBytes(fcClx, [{ cpStart: 0, cpEnd: 1, isCompressed: true, byteOffset: 12345 }])

    const pieces = parseClx(bytes, fcClx, lcbClx)
    // If the flag bit leaked into byteOffset this would be enormous instead.
    expect(pieces[0].byteOffset).toBe(12345)
  })

  it('throws LegacyFormatError for a zero-length CLX', () => {
    expect(() => parseClx(new Uint8Array(16), 0, 0)).toThrow(LegacyFormatError)
  })

  it('throws LegacyFormatError when the CLX has no Pcdt block', () => {
    const buffer = new Uint8Array(16)
    buffer[0] = 0x01 // Prc block claiming to run past the declared CLX with nothing after it
    buffer[1] = 0
    buffer[2] = 0
    expect(() => parseClx(buffer, 0, 3)).toThrow(LegacyFormatError)
  })

  it('throws LegacyFormatError for an unrecognized CLX block type', () => {
    const buffer = new Uint8Array(16)
    buffer[0] = 0x99
    expect(() => parseClx(buffer, 0, 4)).toThrow(LegacyFormatError)
  })

  it('throws LegacyFormatError for a malformed PlcPcd size (not 4 + 12n)', () => {
    const buffer = new Uint8Array(32)
    const view = new DataView(buffer.buffer)
    view.setUint8(0, 0x02)
    view.setUint32(1, 11, true) // not of the form 4 + 12n
    expect(() => parseClx(buffer, 0, 9)).toThrow(LegacyFormatError)
  })

  it('throws LegacyFormatError when the declared piece count exceeds the safety limit', () => {
    const buffer = new Uint8Array(16)
    const view = new DataView(buffer.buffer)
    view.setUint8(0, 0x02)
    const hugePieceCount = 1_000_001
    view.setUint32(1, 4 + hugePieceCount * 12, true)
    expect(() => parseClx(buffer, 0, 5)).toThrow(LegacyFormatError)
  })
})
