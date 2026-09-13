/**
 * Tests for src/docx/parser/endnotes.ts
 */

import { describe, it, expect } from 'vitest'
import { parseEndnotes } from '../endnotes'
import { DocxParseError } from '../unzip'

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

const FULL_ENDNOTES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:type="separator" w:id="-1">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:endnote>
  <w:endnote w:type="continuationSeparator" w:id="0">
    <w:p><w:r><w:continuationSeparator/></w:r></w:p>
  </w:endnote>
  <w:endnote w:id="1">
    <w:p><w:r><w:t>First endnote.</w:t></w:r></w:p>
  </w:endnote>
</w:endnotes>`

const SINGLE_ENDNOTE = `<?xml version="1.0" encoding="UTF-8"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:id="1">
    <w:p><w:r><w:t>Only endnote.</w:t></w:r></w:p>
  </w:endnote>
</w:endnotes>`

const EMPTY_ENDNOTES = `<?xml version="1.0" encoding="UTF-8"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
</w:endnotes>`

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('parseEndnotes', () => {
  it('parses all endnote entries including separators', () => {
    const map = parseEndnotes(FULL_ENDNOTES)
    expect(map.size).toBe(3)
  })

  it('separator endnote carries noteType="separator"', () => {
    const map = parseEndnotes(FULL_ENDNOTES)
    expect(map.get('-1')?.noteType).toBe('separator')
  })

  it('continuationSeparator endnote carries correct noteType', () => {
    const map = parseEndnotes(FULL_ENDNOTES)
    expect(map.get('0')?.noteType).toBe('continuationSeparator')
  })

  it('normal endnote has kind="endnote" and no noteType', () => {
    const map = parseEndnotes(FULL_ENDNOTES)
    const note = map.get('1')
    expect(note?.kind).toBe('endnote')
    expect(note?.noteType).toBeUndefined()
  })

  it('each endnote wraps body in a single UnknownNode (Option B)', () => {
    const map = parseEndnotes(SINGLE_ENDNOTE)
    const note = map.get('1')
    expect(note?.blocks).toHaveLength(1)
    expect(note?.blocks[0].kind).toBe('unknown')
  })

  it('returns empty map for empty <w:endnotes> element', () => {
    const map = parseEndnotes(EMPTY_ENDNOTES)
    expect(map.size).toBe(0)
  })

  it('throws DocxParseError on malformed XML', () => {
    expect(() => parseEndnotes('<<< invalid')).toThrow(DocxParseError)
  })
})
