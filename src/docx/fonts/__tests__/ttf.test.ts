import { readFile } from 'node:fs/promises'
import path from 'node:path'

import { describe, expect, it } from 'vitest'

import { FontParseError, parseTtf } from '../ttf'

const FIXTURE_PATH = path.resolve(process.cwd(), 'public/fonts/Carlito-Regular.ttf')

async function readFixtureBuffer(): Promise<ArrayBuffer> {
  const bytes = await readFile(FIXTURE_PATH)
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)
}

describe('parseTtf', () => {
  it('parses the required tables from Carlito Regular', async () => {
    const tables = parseTtf(await readFixtureBuffer())

    expect(tables.head.unitsPerEm).toBeGreaterThan(0)
    expect(tables.hhea.numberOfHMetrics).toBeGreaterThan(0)
    expect(tables.hmtx.advanceWidths.length).toBeGreaterThan(0)
    expect(tables.cmap.size).toBeGreaterThan(0)
    expect(tables.name.familyName).toBe('Carlito')
    expect(tables.os2.sTypoAscender).toBeGreaterThan(0)
  })

  it('reads Carlito unitsPerEm as 2048', async () => {
    const tables = parseTtf(await readFixtureBuffer())
    expect(tables.head.unitsPerEm).toBe(2048)
  })

  it('exposes the family name from the name table', async () => {
    const tables = parseTtf(await readFixtureBuffer())
    expect(tables.name.familyName).toBe('Carlito')
  })

  it('contains a glyph mapping for uppercase A', async () => {
    const tables = parseTtf(await readFixtureBuffer())
    const glyphId = tables.cmap.get('A'.codePointAt(0) ?? 0)
    expect(glyphId).toBeTypeOf('number')
    expect(glyphId).toBeGreaterThan(0)
  })

  it('includes horizontal metrics for every glyph referenced by cmap', async () => {
    const tables = parseTtf(await readFixtureBuffer())
    for (const glyphId of tables.cmap.values()) {
      expect(tables.hmtx.advanceWidths[glyphId]).toBeTypeOf('number')
    }
  })

  it('throws FontParseError for unsupported signatures', async () => {
    const buffer = await readFixtureBuffer()
    const bytes = new Uint8Array(buffer.slice(0))
    bytes[0] = 0xff
    expect(() => parseTtf(bytes.buffer)).toThrow(FontParseError)
  })

  it('throws when the table directory is truncated', async () => {
    const buffer = await readFixtureBuffer()
    const truncated = buffer.slice(0, 20)
    expect(() => parseTtf(truncated)).toThrow(FontParseError)
  })

  it('throws when the head table record is removed', async () => {
    const buffer = await readFixtureBuffer()
    const bytes = new Uint8Array(buffer.slice(0))
    bytes[4] = 0x00
    bytes[5] = 0x00
    expect(() => parseTtf(bytes.buffer)).toThrow(/Missing required head table/)
  })
})
