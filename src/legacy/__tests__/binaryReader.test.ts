import { describe, expect, it } from 'vitest'

import { BinaryReader } from '../binaryReader'
import { LegacyFormatError } from '../errors'

describe('BinaryReader', () => {
  it('reads little-endian u8/u16/u32/i32 values', () => {
    const bytes = new Uint8Array([0x01, 0x02, 0x03, 0x04, 0xff, 0xff, 0xff, 0xff])
    const reader = new BinaryReader(bytes)

    expect(reader.u8(0)).toBe(0x01)
    expect(reader.u16(0)).toBe(0x0201)
    expect(reader.u32(0)).toBe(0x04030201)
    expect(reader.i32(4)).toBe(-1)
  })

  it('reads from a Uint8Array that is a slice of a larger buffer (byteOffset honored)', () => {
    const backing = new Uint8Array([0xaa, 0xaa, 0x2a, 0x00, 0xaa, 0xaa])
    const view = backing.subarray(2, 4) // [0x2a, 0x00]
    const reader = new BinaryReader(view)

    expect(reader.u16(0)).toBe(0x002a)
  })

  it('exposes the byte length', () => {
    expect(new BinaryReader(new Uint8Array(10)).length).toBe(10)
  })

  it('returns a subarray view without copying', () => {
    const bytes = new Uint8Array([1, 2, 3, 4, 5])
    const reader = new BinaryReader(bytes)
    const view = reader.subarray(1, 3)

    expect(Array.from(view)).toEqual([2, 3, 4])
    expect(view.buffer).toBe(bytes.buffer)
  })

  it.each([
    ['u8', (r: BinaryReader) => r.u8(5)],
    ['u16', (r: BinaryReader) => r.u16(4)],
    ['u32', (r: BinaryReader) => r.u32(2)],
    ['i32', (r: BinaryReader) => r.i32(2)],
    ['subarray', (r: BinaryReader) => r.subarray(2, 10)],
  ] as const)('%s throws a LegacyFormatError instead of a raw RangeError when out of bounds', (_label, read) => {
    const reader = new BinaryReader(new Uint8Array(5))
    expect(() => read(reader)).toThrow(LegacyFormatError)
  })

  it('rejects a negative offset', () => {
    const reader = new BinaryReader(new Uint8Array(5))
    expect(() => reader.u8(-1)).toThrow(LegacyFormatError)
  })
})
