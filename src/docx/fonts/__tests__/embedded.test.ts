/**
 * Tests for src/docx/fonts/embedded.ts (DEFER-4 / DXP-13)
 *
 * Builds a synthetic obfuscated-font fixture in-memory (no real Word-authored
 * embedded-font .docx needed): a bundled OFL TTF is obfuscated with
 * `xorObfuscatedFontHeader` under a fixed test GUID — the same operation
 * Word applies, since XOR is self-inverse — then a minimal fontTable.xml +
 * fontTable.xml.rels + font part are assembled into a raw-archive files map
 * exactly like `unzipDocx` would produce, and `loadEmbeddedFonts` is
 * exercised end to end against it.
 */

import { readFile } from 'node:fs/promises'
import path from 'node:path'

import { describe, expect, it } from 'vitest'

import { xorObfuscatedFontHeader } from '../deobfuscate'
import { loadEmbeddedFonts } from '../embedded'

const FIXTURE_PATH = path.resolve(process.cwd(), 'public/fonts/Carlito-Regular.ttf')
const TEST_GUID = '{5F3759DF-0000-0000-0000-000000000001}'

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

const RELS_XML_TEMPLATE = (target: string) => `<?xml version="1.0"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1"
      Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
      Target="${target}"/>
  </Relationships>`

function fontTableXml(fontKeyAttr: string): string {
  return `<?xml version="1.0"?>
    <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <w:font w:name="Custom Embedded Font">
        <w:embedRegular r:id="rId1"${fontKeyAttr}/>
      </w:font>
    </w:fonts>`
}

function encodeUtf8(xml: string): Uint8Array {
  return new TextEncoder().encode(xml)
}

describe('loadEmbeddedFonts', () => {
  it('returns [] when the archive has no fontTable.xml', () => {
    expect(loadEmbeddedFonts(new Map())).toEqual([])
  })

  it('returns [] when files is undefined', () => {
    expect(loadEmbeddedFonts(undefined)).toEqual([])
  })

  it('returns [] when fontTable.xml lists fonts but none are embedded', () => {
    const files = new Map<string, Uint8Array>([
      [
        'word/fontTable.xml',
        encodeUtf8(
          '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
            '<w:font w:name="Calibri"/></w:fonts>',
        ),
      ],
    ])
    expect(loadEmbeddedFonts(files)).toEqual([])
  })

  it('de-obfuscates and resolves an embedded regular face end to end', async () => {
    const original = await readFixtureBytes()
    const obfuscated = xorObfuscatedFontHeader(original, TEST_GUID)

    const files = new Map<string, Uint8Array>([
      ['word/fontTable.xml', encodeUtf8(fontTableXml(` w:fontKey="${TEST_GUID}"`))],
      ['word/_rels/fontTable.xml.rels', encodeUtf8(RELS_XML_TEMPLATE('fonts/font1.fntdata'))],
      ['word/fonts/font1.fntdata', obfuscated],
    ])

    const families = loadEmbeddedFonts(files)

    expect(families).toHaveLength(1)
    expect(families[0]?.name).toBe('Custom Embedded Font')
    expect(families[0]?.faces.regular).toEqual(original)
    expect(families[0]?.faces.bold).toBeUndefined()
  })

  it('resolves a relationship target with a leading slash as package-absolute', async () => {
    const original = await readFixtureBytes()
    const obfuscated = xorObfuscatedFontHeader(original, TEST_GUID)

    const files = new Map<string, Uint8Array>([
      ['word/fontTable.xml', encodeUtf8(fontTableXml(` w:fontKey="${TEST_GUID}"`))],
      ['word/_rels/fontTable.xml.rels', encodeUtf8(RELS_XML_TEMPLATE('/word/fonts/font1.fntdata'))],
      ['word/fonts/font1.fntdata', obfuscated],
    ])

    const families = loadEmbeddedFonts(files)
    expect(families[0]?.faces.regular).toEqual(original)
  })

  it('skips a face whose font part is missing from the archive, without throwing', () => {
    const files = new Map<string, Uint8Array>([
      ['word/fontTable.xml', encodeUtf8(fontTableXml(` w:fontKey="${TEST_GUID}"`))],
      ['word/_rels/fontTable.xml.rels', encodeUtf8(RELS_XML_TEMPLATE('fonts/font1.fntdata'))],
      // No 'word/fonts/font1.fntdata' entry.
    ])

    expect(loadEmbeddedFonts(files)).toEqual([])
  })

  it('skips a face whose relationship id is not present in fontTable.xml.rels', () => {
    const files = new Map<string, Uint8Array>([
      ['word/fontTable.xml', encodeUtf8(fontTableXml(` w:fontKey="${TEST_GUID}"`))],
      // No .rels part at all.
      ['word/fonts/font1.fntdata', new Uint8Array([1, 2, 3])],
    ])

    expect(loadEmbeddedFonts(files)).toEqual([])
  })

  it('skips a face with a malformed w:fontKey rather than failing the whole document', () => {
    const files = new Map<string, Uint8Array>([
      ['word/fontTable.xml', encodeUtf8(fontTableXml(' w:fontKey="not-a-guid"'))],
      ['word/_rels/fontTable.xml.rels', encodeUtf8(RELS_XML_TEMPLATE('fonts/font1.fntdata'))],
      ['word/fonts/font1.fntdata', new Uint8Array([1, 2, 3, 4])],
    ])

    expect(loadEmbeddedFonts(files)).toEqual([])
  })

  it('uses the raw bytes as-is when no w:fontKey is present (unusual but not fatal)', () => {
    const rawBytes = new Uint8Array([9, 8, 7, 6])
    const files = new Map<string, Uint8Array>([
      ['word/fontTable.xml', encodeUtf8(fontTableXml(''))],
      ['word/_rels/fontTable.xml.rels', encodeUtf8(RELS_XML_TEMPLATE('fonts/font1.fntdata'))],
      ['word/fonts/font1.fntdata', rawBytes],
    ])

    const families = loadEmbeddedFonts(files)
    expect(families[0]?.faces.regular).toEqual(rawBytes)
  })

  it('resolves multiple variants for the same font family independently', async () => {
    const original = await readFixtureBytes()
    const regularObfuscated = xorObfuscatedFontHeader(original, TEST_GUID)
    const boldGuid = '{5F3759DF-0000-0000-0000-000000000002}'
    const boldObfuscated = xorObfuscatedFontHeader(original, boldGuid)

    const xml = `<?xml version="1.0"?>
      <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:font w:name="Custom Embedded Font">
          <w:embedRegular r:id="rId1" w:fontKey="${TEST_GUID}"/>
          <w:embedBold r:id="rId2" w:fontKey="${boldGuid}"/>
        </w:font>
      </w:fonts>`

    const relsXml = `<?xml version="1.0"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
          Target="fonts/font1.fntdata"/>
        <Relationship Id="rId2"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
          Target="fonts/font2.fntdata"/>
      </Relationships>`

    const files = new Map<string, Uint8Array>([
      ['word/fontTable.xml', encodeUtf8(xml)],
      ['word/_rels/fontTable.xml.rels', encodeUtf8(relsXml)],
      ['word/fonts/font1.fntdata', regularObfuscated],
      ['word/fonts/font2.fntdata', boldObfuscated],
    ])

    const families = loadEmbeddedFonts(files)
    expect(families).toHaveLength(1)
    expect(families[0]?.faces.regular).toEqual(original)
    expect(families[0]?.faces.bold).toEqual(original)
  })
})
