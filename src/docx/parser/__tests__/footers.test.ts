/**
 * Tests for src/docx/parser/footers.ts
 */

import { describe, it, expect } from 'vitest'
import { parseFooter } from '../footers'
import { DocxParseError } from '../unzip'
import { writeFooterXml } from '../../serializer/footerWriter'

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

const SIMPLE_FOOTER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t>Page Footer</w:t></w:r>
  </w:p>
</w:ftr>`

const MULTI_PARAGRAPH_FOOTER = `<?xml version="1.0" encoding="UTF-8"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>First line</w:t></w:r></w:p>
  <w:p><w:r><w:t>Second line</w:t></w:r></w:p>
</w:ftr>`

const EMPTY_FOOTER = `<?xml version="1.0" encoding="UTF-8"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
</w:ftr>`

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('parseFooter', () => {
  it('returns a Footer node with kind="footer"', () => {
    const footer = parseFooter(SIMPLE_FOOTER, 'rId4')
    expect(footer.kind).toBe('footer')
  })

  it('stores the supplied relationship id on the Footer', () => {
    const footer = parseFooter(SIMPLE_FOOTER, 'rId4')
    expect(footer.id).toBe('rId4')
  })

  it('parses the body into a real paragraph block, not an opaque UnknownNode', () => {
    const footer = parseFooter(SIMPLE_FOOTER, 'rId4')
    expect(footer.blocks).toHaveLength(1)
    expect(footer.blocks[0].kind).toBe('paragraph')
  })

  it('extracts the run text from the parsed paragraph', () => {
    const footer = parseFooter(SIMPLE_FOOTER, 'rId4')
    const block = footer.blocks[0]
    expect(block.kind).toBe('paragraph')
    if (block.kind === 'paragraph') {
      const run = block.children[0]
      expect(run.kind).toBe('run')
      if (run.kind === 'run') {
        expect(run.children[0]).toEqual({ kind: 'text', value: 'Page Footer' })
      }
    }
  })

  it('parses multiple paragraphs in source order', () => {
    const footer = parseFooter(MULTI_PARAGRAPH_FOOTER, 'rId4')
    expect(footer.blocks).toHaveLength(2)
    expect(footer.blocks[0].kind).toBe('paragraph')
    expect(footer.blocks[1].kind).toBe('paragraph')
  })

  it('works with an empty footer body', () => {
    const footer = parseFooter(EMPTY_FOOTER, 'rId6')
    expect(footer.kind).toBe('footer')
    expect(footer.id).toBe('rId6')
    expect(footer.blocks).toHaveLength(0)
  })

  it('defaults id to empty string when omitted', () => {
    const footer = parseFooter(SIMPLE_FOOTER)
    expect(footer.id).toBe('')
  })

  it('throws DocxParseError for completely invalid XML', () => {
    expect(() => parseFooter('<<< not xml <<<', 'rId2')).toThrow(DocxParseError)
  })

  it('round-trips through the serializer without throwing (DXS-01)', () => {
    const footer = parseFooter(SIMPLE_FOOTER, 'rId4')
    const xml = writeFooterXml(footer)
    expect(xml).toContain('Page Footer')
  })
})
