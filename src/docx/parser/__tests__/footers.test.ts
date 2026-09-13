/**
 * Tests for src/docx/parser/footers.ts
 */

import { describe, it, expect } from 'vitest'
import { parseFooter } from '../footers'
import { DocxParseError } from '../unzip'

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

const SIMPLE_FOOTER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t>Page Footer</w:t></w:r>
  </w:p>
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

  it('wraps inner XML in a single UnknownNode (Option B)', () => {
    const footer = parseFooter(SIMPLE_FOOTER, 'rId4')
    expect(footer.blocks).toHaveLength(1)
    expect(footer.blocks[0].kind).toBe('unknown')
  })

  it('preserves full XML in the UnknownNode so A.3 can re-parse it', () => {
    const footer = parseFooter(SIMPLE_FOOTER, 'rId4')
    const node = footer.blocks[0]
    expect(node.kind).toBe('unknown')
    if (node.kind === 'unknown') {
      expect(node.xml).toContain('Page Footer')
    }
  })

  it('works with an empty footer body', () => {
    const footer = parseFooter(EMPTY_FOOTER, 'rId6')
    expect(footer.kind).toBe('footer')
    expect(footer.id).toBe('rId6')
  })

  it('defaults id to empty string when omitted', () => {
    const footer = parseFooter(SIMPLE_FOOTER)
    expect(footer.id).toBe('')
  })

  it('throws DocxParseError for completely invalid XML', () => {
    expect(() => parseFooter('<<< not xml <<<', 'rId2')).toThrow(DocxParseError)
  })
})
