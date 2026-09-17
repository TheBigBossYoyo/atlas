import { describe, expect, it } from 'vitest'

import { BinaryReader } from '../../binaryReader'
import { walkPptRecords, type PptRecordHeader } from '../records'

/** Writes one [MS-PPT] RecordHeader (8 bytes) + `payload` at `offset`, returns the offset just past it. */
function writeRecord(
  view: DataView,
  offset: number,
  options: { readonly recVer: number; readonly recInstance?: number; readonly type: number },
  payload: ReadonlyArray<number> = [],
): number {
  const verInstance = (options.recVer & 0x000f) | (((options.recInstance ?? 0) & 0x0fff) << 4)
  view.setUint16(offset, verInstance, true)
  view.setUint16(offset + 2, options.type, true)
  view.setUint32(offset + 4, payload.length, true)
  new Uint8Array(view.buffer, offset + 8, payload.length).set(payload)
  return offset + 8 + payload.length
}

describe('walkPptRecords', () => {
  it('visits a single atom record', () => {
    const buffer = new ArrayBuffer(16)
    const view = new DataView(buffer)
    const end = writeRecord(view, 0, { recVer: 0x0, type: 4000 }, [1, 2, 3, 4])

    const seen: PptRecordHeader[] = []
    walkPptRecords(new BinaryReader(new Uint8Array(buffer)), 0, end, (h) => seen.push(h))

    expect(seen).toHaveLength(1)
    expect(seen[0]).toMatchObject({ recVer: 0, type: 4000, dataStart: 8, dataEnd: 12 })
  })

  it('recurses into a container (recVer 0xF) and visits its children in document order', () => {
    const buffer = new ArrayBuffer(64)
    const view = new DataView(buffer)

    // Container of length 16, holding two 8-byte-payload-less atoms.
    let offset = 0
    const containerStart = offset
    offset += 8 // header written after we know the child bytes' total length
    const childrenStart = offset
    offset = writeRecord(view, offset, { recVer: 0x0, type: 100 })
    offset = writeRecord(view, offset, { recVer: 0x0, type: 200 })
    const childrenLength = offset - childrenStart
    // Back-patch the container's own header now the length is known.
    view.setUint16(containerStart, 0x000f, true) // recVer 0xF
    view.setUint16(containerStart + 2, 999, true)
    view.setUint32(containerStart + 4, childrenLength, true)

    const seen: Array<{ type: number; recVer: number }> = []
    walkPptRecords(new BinaryReader(new Uint8Array(buffer)), 0, offset, (h) => seen.push({ type: h.type, recVer: h.recVer }))

    expect(seen).toEqual([
      { type: 999, recVer: 0x0f },
      { type: 100, recVer: 0x0 },
      { type: 200, recVer: 0x0 },
    ])
  })

  it('decodes recInstance from the high 12 bits of verInstance', () => {
    const buffer = new ArrayBuffer(16)
    const view = new DataView(buffer)
    const end = writeRecord(view, 0, { recVer: 0x0, recInstance: 0x0abc, type: 1 })

    const seen: PptRecordHeader[] = []
    walkPptRecords(new BinaryReader(new Uint8Array(buffer)), 0, end, (h) => seen.push(h))

    expect(seen[0].recInstance).toBe(0x0abc)
  })

  it('stops walking a level when a record length would run past the given end (corrupt length)', () => {
    const buffer = new ArrayBuffer(16)
    const view = new DataView(buffer)
    view.setUint16(0, 0x0000, true)
    view.setUint16(2, 1, true)
    view.setUint32(4, 1000, true) // claims 1000 bytes of payload — far past the buffer

    const seen: PptRecordHeader[] = []
    walkPptRecords(new BinaryReader(new Uint8Array(buffer)), 0, 16, (h) => seen.push(h))

    expect(seen).toHaveLength(0)
  })

  it('visits sibling records after a container one, not just its children', () => {
    const buffer = new ArrayBuffer(64)
    const view = new DataView(buffer)

    let offset = 0
    const containerStart = offset
    offset += 8
    const childStart = offset
    offset = writeRecord(view, offset, { recVer: 0x0, type: 1 })
    view.setUint16(containerStart, 0x000f, true)
    view.setUint16(containerStart + 2, 500, true)
    view.setUint32(containerStart + 4, offset - childStart, true)

    offset = writeRecord(view, offset, { recVer: 0x0, type: 2 }) // sibling after the container

    const seen: number[] = []
    walkPptRecords(new BinaryReader(new Uint8Array(buffer)), 0, offset, (h) => seen.push(h.type))

    expect(seen).toEqual([500, 1, 2])
  })

  it('bounds recursion on a very deeply nested record tree instead of stack-overflowing', () => {
    // DEPTH containers, each one's entire payload being the next nested
    // container's header-and-onward, terminating in a zero-length payload.
    const DEPTH = 200
    const total = 8 * DEPTH
    const buffer = new ArrayBuffer(total)
    const view = new DataView(buffer)

    for (let i = 0; i < DEPTH; i += 1) {
      const offset = i * 8
      const remaining = total - offset - 8
      view.setUint16(offset, 0x000f, true) // recVer 0xF — container
      view.setUint16(offset + 2, i, true) // type == nesting depth, for the assertion below
      view.setUint32(offset + 4, remaining, true)
    }

    const seen: number[] = []
    expect(() => {
      walkPptRecords(new BinaryReader(new Uint8Array(buffer)), 0, total, (h) => seen.push(h.type))
    }).not.toThrow()

    // The depth guard must have cut this off well short of all 200 levels.
    expect(seen.length).toBeGreaterThan(0)
    expect(seen.length).toBeLessThan(DEPTH)
    // Each visited container's own recorded depth (its `type`) increases by
    // exactly one per level, with no gaps, up to wherever the walk stopped.
    expect(seen).toEqual(seen.map((_, i) => i))
  })
})
