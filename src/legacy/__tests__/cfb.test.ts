/**
 * Tests for src/legacy/cfb.ts
 *
 * Builds real (small, in-memory) OLE2/CFB containers (see `./fixtures.ts`)
 * rather than committing binary fixture files — `CFB.write`'s output
 * round-trips through `CFB.read` exactly the way a real .doc/.ppt file would.
 */
import { describe, expect, it } from 'vitest'

import { findEntry, readCfb } from '../cfb'
import { LegacyFormatError } from '../errors'
import { buildCfbBytes } from './fixtures'

describe('readCfb', () => {
  it('parses a real CFB container into typed entries', () => {
    const bytes = buildCfbBytes([['WordDocument', new Uint8Array([1, 2, 3, 4])]])
    const cfb = readCfb(bytes)

    const wordDoc = cfb.entries.find((e) => e.name === 'WordDocument')
    expect(wordDoc).toBeDefined()
    expect(wordDoc?.type).toBe('stream')
    expect(Array.from(wordDoc?.content ?? [])).toEqual([1, 2, 3, 4])
  })

  it('reads a stream past the 4096-byte OLE2 ministream cutoff (RUN-08-style: a real document is almost always this big)', () => {
    // `XLSX.CFB`'s codec hands back a stream at least this size as a plain
    // `Array<number>` (not a `Buffer`/typed array) when read with `{ type:
    // 'array' }` — see `toContent`'s comment. Every OTHER fixture in this
    // suite is a handful of bytes, well under the cutoff, so this never
    // exercised the FAT sector-chain path at all — only the small-stream
    // ministream path, which was never affected.
    const bigContent = new Uint8Array(5000).map((_, i) => i % 256)
    const bytes = buildCfbBytes([['WordDocument', bigContent]])
    const cfb = readCfb(bytes)

    const wordDoc = cfb.entries.find((e) => e.name === 'WordDocument')
    expect(wordDoc?.type).toBe('stream')
    expect(wordDoc?.content).not.toBeNull()
    expect(wordDoc?.content?.length).toBe(5000)
    expect(Array.from(wordDoc?.content?.slice(0, 4) ?? [])).toEqual([0, 1, 2, 3])
  })

  it('includes the root entry', () => {
    const bytes = buildCfbBytes([['WordDocument', new Uint8Array([1])]])
    const cfb = readCfb(bytes)

    expect(cfb.entries.some((e) => e.type === 'root')).toBe(true)
  })

  it('throws LegacyFormatError for a buffer too small to be a CFB file', () => {
    expect(() => readCfb(new Uint8Array([1, 2, 3]))).toThrow(LegacyFormatError)
  })

  it('throws LegacyFormatError for a buffer with the wrong magic entirely', () => {
    const bytes = new Uint8Array(600) // right size, all-zero — not a valid CFB header
    expect(() => readCfb(bytes)).toThrow(LegacyFormatError)
  })
})

describe('findEntry', () => {
  it('finds a stream by its bare name', () => {
    const cfb = readCfb(buildCfbBytes([['1Table', new Uint8Array([9, 9])]]))
    const entry = findEntry(cfb, '1Table')

    expect(entry).not.toBeNull()
    expect(entry?.name).toBe('1Table')
  })

  it('matches case-insensitively', () => {
    const cfb = readCfb(buildCfbBytes([['PowerPoint Document', new Uint8Array([1])]]))
    expect(findEntry(cfb, 'powerpoint document')).not.toBeNull()
  })

  it('returns null for a name that is not present', () => {
    const cfb = readCfb(buildCfbBytes([['WordDocument', new Uint8Array([1])]]))
    expect(findEntry(cfb, 'NotThere')).toBeNull()
  })

  it('distinguishes 0Table from 1Table', () => {
    const cfb = readCfb(
      buildCfbBytes([
        ['0Table', new Uint8Array([0])],
        ['1Table', new Uint8Array([1])],
      ]),
    )

    expect(findEntry(cfb, '0Table')?.content?.[0]).toBe(0)
    expect(findEntry(cfb, '1Table')?.content?.[0]).toBe(1)
  })
})
