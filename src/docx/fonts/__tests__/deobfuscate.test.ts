/**
 * Tests for src/docx/fonts/deobfuscate.ts (DEFER-4 / DXP-13)
 */

import { readFile } from 'node:fs/promises'
import path from 'node:path'

import { describe, expect, it } from 'vitest'

import { FontKeyError, guidToFontKeyBytes, xorObfuscatedFontHeader } from '../deobfuscate'

const FIXTURE_PATH = path.resolve(process.cwd(), 'public/fonts/Carlito-Regular.ttf')

// Memoized: several tests in this file read the same on-disk fixture, and
// under `--maxWorkers` contention repeating that disk I/O per test was
// observed to occasionally exceed vitest's default 5000ms per-test budget
// on a loaded machine — reading it once and reusing the bytes removes the
// redundant I/O rather than just padding the timeout.
let fixtureBytesPromise: Promise<Uint8Array> | undefined

async function readFixtureBytes(): Promise<Uint8Array> {
  fixtureBytesPromise ??= readFile(FIXTURE_PATH).then(
    (buffer) => new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength),
  )
  return fixtureBytesPromise
}

describe('guidToFontKeyBytes', () => {
  it('converts a braced GUID to its 16-byte little-endian/as-is struct layout', () => {
    // Data1=01020304 (LE -> 04 03 02 01), Data2=0506 (LE -> 06 05),
    // Data3=0708 (LE -> 08 07), Data4=090A0B0C0D0E0F10 (as-is).
    const key = guidToFontKeyBytes('{01020304-0506-0708-090A-0B0C0D0E0F10}')
    expect(Array.from(key)).toEqual([
      0x04, 0x03, 0x02, 0x01, 0x06, 0x05, 0x08, 0x07, 0x09, 0x0a, 0x0b, 0x0c, 0x0d, 0x0e, 0x0f, 0x10,
    ])
  })

  it('accepts an unbraced GUID', () => {
    const braced = guidToFontKeyBytes('{01020304-0506-0708-090A-0B0C0D0E0F10}')
    const unbraced = guidToFontKeyBytes('01020304-0506-0708-090A-0B0C0D0E0F10')
    expect(Array.from(unbraced)).toEqual(Array.from(braced))
  })

  it('throws FontKeyError for a malformed GUID', () => {
    expect(() => guidToFontKeyBytes('not-a-guid')).toThrow(FontKeyError)
  })
})

describe('xorObfuscatedFontHeader', () => {
  const GUID = '{3D2F1A00-1111-2222-3333-444455556666}'

  it('is self-inverse: obfuscating then de-obfuscating with the same key restores the original bytes', async () => {
    const original = await readFixtureBytes()

    const obfuscated = xorObfuscatedFontHeader(original, GUID)
    const restored = xorObfuscatedFontHeader(obfuscated, GUID)

    expect(restored).toEqual(original)
  })

  it('only touches the first 32 bytes, leaving the rest of the font untouched', async () => {
    const original = await readFixtureBytes()
    const obfuscated = xorObfuscatedFontHeader(original, GUID)

    expect(obfuscated.slice(32)).toEqual(original.slice(32))
    // Sanity: the header actually changed (extremely unlikely to coincide with a real font's bytes).
    expect(obfuscated.slice(0, 32)).not.toEqual(original.slice(0, 32))
  })

  it('does not mutate the input array', () => {
    const data = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9, 10])
    const copy = Uint8Array.from(data)
    xorObfuscatedFontHeader(data, GUID)
    expect(data).toEqual(copy)
  })

  it('handles buffers shorter than 32 bytes without throwing', () => {
    const data = new Uint8Array([1, 2, 3, 4])
    const result = xorObfuscatedFontHeader(data, GUID)
    expect(result.length).toBe(4)
    // Round-trips too.
    expect(xorObfuscatedFontHeader(result, GUID)).toEqual(data)
  })

  it('throws FontKeyError for an invalid key, matching guidToFontKeyBytes', () => {
    expect(() => xorObfuscatedFontHeader(new Uint8Array([1, 2, 3]), 'bogus')).toThrow(FontKeyError)
  })
})
