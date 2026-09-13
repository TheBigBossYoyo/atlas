/**
 * Atlas — zipPackager tests (Wave D.4)
 */

import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import {
  packDocx,
  sortPartsForWord,
  WORD_PART_ORDER,
  type DocxPart,
} from '../zipPackager'

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const CONTENT_TYPES_XML = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'

const DOCUMENT_XML = '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>'

const MINIMAL_PARTS: ReadonlyArray<DocxPart> = [
  { path: '[Content_Types].xml', content: CONTENT_TYPES_XML },
  { path: 'word/document.xml', content: DOCUMENT_XML },
]

// ---------------------------------------------------------------------------
// packDocx tests
// ---------------------------------------------------------------------------

describe('packDocx', () => {
  it('returns a non-empty Uint8Array for a minimal DOCX', async () => {
    const result = await packDocx(MINIMAL_PARTS)
    expect(result).toBeInstanceOf(Uint8Array)
    expect(result.length).toBeGreaterThan(0)
  })

  it('output starts with PKZip magic bytes 0x50 0x4B 0x03 0x04', async () => {
    const result = await packDocx(MINIMAL_PARTS)
    expect(result[0]).toBe(0x50)
    expect(result[1]).toBe(0x4b)
    expect(result[2]).toBe(0x03)
    expect(result[3]).toBe(0x04)
  })

  it('round-trip: pack → loadAsync → all parts present', async () => {
    const result = await packDocx(MINIMAL_PARTS)
    const zip = await JSZip.loadAsync(result)
    expect(zip.files['[Content_Types].xml']).toBeDefined()
    expect(zip.files['word/document.xml']).toBeDefined()
  })

  it('round-trip: packed content matches original string content', async () => {
    const result = await packDocx(MINIMAL_PARTS)
    const zip = await JSZip.loadAsync(result)
    const docXml = await zip.files['word/document.xml'].async('string')
    expect(docXml).toBe(DOCUMENT_XML)
  })

  it('round-trip: Uint8Array content is preserved correctly', async () => {
    const bytes = new Uint8Array([0x01, 0x02, 0x03, 0x04])
    const parts: ReadonlyArray<DocxPart> = [
      { path: 'word/media/image1.png', content: bytes },
    ]
    const result = await packDocx(parts)
    const zip = await JSZip.loadAsync(result)
    const extracted = await zip.files['word/media/image1.png'].async('uint8array')
    expect(extracted).toEqual(bytes)
  })
})

// ---------------------------------------------------------------------------
// sortPartsForWord tests
// ---------------------------------------------------------------------------

describe('sortPartsForWord', () => {
  const MIXED_PARTS: ReadonlyArray<DocxPart> = [
    { path: 'word/footnotes.xml', content: '' },
    { path: 'word/header1.xml', content: '' },
    { path: '[Content_Types].xml', content: '' },
    { path: '_rels/.rels', content: '' },
    { path: 'word/document.xml', content: '' },
    { path: 'zzz/unknown.xml', content: '' },
    { path: 'aaa/another.xml', content: '' },
  ]

  it('is pure: returns a new array without mutating the input', () => {
    const input = [...MIXED_PARTS]
    const result = sortPartsForWord(input)
    expect(result).not.toBe(input)
    // input still in original order
    expect(input[0].path).toBe('word/footnotes.xml')
  })

  it('puts [Content_Types].xml first', () => {
    const result = sortPartsForWord(MIXED_PARTS)
    expect(result[0].path).toBe('[Content_Types].xml')
  })

  it('puts _rels/.rels second', () => {
    const result = sortPartsForWord(MIXED_PARTS)
    expect(result[1].path).toBe('_rels/.rels')
  })

  it('puts word/document.xml third', () => {
    const result = sortPartsForWord(MIXED_PARTS)
    expect(result[2].path).toBe('word/document.xml')
  })

  it('puts word/header1.xml before word/footnotes.xml (header pattern)', () => {
    const result = sortPartsForWord(MIXED_PARTS)
    const headerIdx = result.findIndex(p => p.path === 'word/header1.xml')
    const footnotesIdx = result.findIndex(p => p.path === 'word/footnotes.xml')
    expect(headerIdx).toBeLessThan(footnotesIdx)
  })

  it('unknown paths are sorted alphabetically at the end', () => {
    const result = sortPartsForWord(MIXED_PARTS)
    const paths = result.map(p => p.path)
    const aaaIdx = paths.indexOf('aaa/another.xml')
    const zzzIdx = paths.indexOf('zzz/unknown.xml')
    // both are after all known patterns
    expect(aaaIdx).toBeLessThan(zzzIdx)
    expect(aaaIdx).toBeGreaterThan(paths.indexOf('word/footnotes.xml'))
  })

  it('handles empty array', () => {
    const result = sortPartsForWord([])
    expect(result).toEqual([])
  })
})

// ---------------------------------------------------------------------------
// WORD_PART_ORDER tests
// ---------------------------------------------------------------------------

describe('WORD_PART_ORDER', () => {
  it('includes [Content_Types].xml', () => {
    expect(WORD_PART_ORDER).toContain('[Content_Types].xml')
  })

  it('includes _rels/.rels', () => {
    expect(WORD_PART_ORDER).toContain('_rels/.rels')
  })

  it('includes word/document.xml', () => {
    expect(WORD_PART_ORDER).toContain('word/document.xml')
  })

  it('includes docProps/core.xml and docProps/app.xml', () => {
    expect(WORD_PART_ORDER).toContain('docProps/core.xml')
    expect(WORD_PART_ORDER).toContain('docProps/app.xml')
  })

  it('is a non-empty readonly array', () => {
    expect(WORD_PART_ORDER.length).toBeGreaterThan(0)
  })
})
