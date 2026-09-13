/**
 * Tests for src/docx/parser/headers.ts
 */

import { describe, it, expect } from 'vitest'
import { parseHeader } from '../headers'
import { DocxParseError } from '../unzip'
import { writeHeaderXml } from '../../serializer/headerWriter'

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

const SIMPLE_HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
       xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t>Page Header</w:t></w:r>
  </w:p>
</w:hdr>`

const MULTI_PARAGRAPH_HEADER = `<?xml version="1.0" encoding="UTF-8"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>First line</w:t></w:r></w:p>
  <w:p><w:r><w:t>Second line</w:t></w:r></w:p>
</w:hdr>`

const EMPTY_HEADER = `<?xml version="1.0" encoding="UTF-8"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
</w:hdr>`

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('parseHeader', () => {
  it('returns a Header node with kind="header"', () => {
    const header = parseHeader(SIMPLE_HEADER, 'rId3')
    expect(header.kind).toBe('header')
  })

  it('stores the supplied relationship id on the Header', () => {
    const header = parseHeader(SIMPLE_HEADER, 'rId3')
    expect(header.id).toBe('rId3')
  })

  it('parses the body into a real paragraph block, not an opaque UnknownNode', () => {
    const header = parseHeader(SIMPLE_HEADER, 'rId3')
    expect(header.blocks).toHaveLength(1)
    expect(header.blocks[0].kind).toBe('paragraph')
  })

  it('extracts the run text from the parsed paragraph', () => {
    const header = parseHeader(SIMPLE_HEADER, 'rId3')
    const block = header.blocks[0]
    expect(block.kind).toBe('paragraph')
    if (block.kind === 'paragraph') {
      const run = block.children[0]
      expect(run.kind).toBe('run')
      if (run.kind === 'run') {
        expect(run.children[0]).toEqual({ kind: 'text', value: 'Page Header' })
      }
    }
  })

  it('parses multiple paragraphs in source order', () => {
    const header = parseHeader(MULTI_PARAGRAPH_HEADER, 'rId3')
    expect(header.blocks).toHaveLength(2)
    expect(header.blocks[0].kind).toBe('paragraph')
    expect(header.blocks[1].kind).toBe('paragraph')
  })

  it('works with an empty header body', () => {
    const header = parseHeader(EMPTY_HEADER, 'rId5')
    expect(header.kind).toBe('header')
    expect(header.id).toBe('rId5')
    expect(header.blocks).toHaveLength(0)
  })

  it('defaults id to empty string when omitted', () => {
    const header = parseHeader(SIMPLE_HEADER)
    expect(header.id).toBe('')
  })

  it('throws DocxParseError for completely invalid XML', () => {
    expect(() => parseHeader('<<< not xml <<<', 'rId1')).toThrow(DocxParseError)
  })

  it('round-trips through the serializer without throwing (DXS-01)', () => {
    const header = parseHeader(SIMPLE_HEADER, 'rId3')
    const xml = writeHeaderXml(header)
    expect(xml).toContain('Page Header')
  })
})
