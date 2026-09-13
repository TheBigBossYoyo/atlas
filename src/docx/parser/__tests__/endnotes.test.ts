/**
 * Tests for src/docx/parser/endnotes.ts
 */

import { describe, it, expect } from 'vitest'
import { parseEndnotes } from '../endnotes'
import { DocxParseError } from '../unzip'
import { writeEndnotesXml } from '../../serializer/footnotesWriter'

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

// wave 1 follow-up: an endnote containing a table must survive instead of
// being silently dropped.
const ENDNOTES_WITH_TABLE = `<?xml version="1.0" encoding="UTF-8"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:id="1">
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>Endnote table cell</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
  </w:endnote>
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

  it('parses each endnote body into a real paragraph block, not an opaque UnknownNode', () => {
    const map = parseEndnotes(SINGLE_ENDNOTE)
    const note = map.get('1')
    expect(note?.blocks).toHaveLength(1)
    expect(note?.blocks[0].kind).toBe('paragraph')
  })

  it('extracts the run text from the parsed endnote paragraph', () => {
    const map = parseEndnotes(SINGLE_ENDNOTE)
    const block = map.get('1')?.blocks[0]
    expect(block?.kind).toBe('paragraph')
    if (block?.kind === 'paragraph') {
      const run = block.children[0]
      expect(run.kind).toBe('run')
      if (run.kind === 'run') {
        expect(run.children[0]).toEqual({ kind: 'text', value: 'Only endnote.' })
      }
    }
  })

  it('returns empty map for empty <w:endnotes> element', () => {
    const map = parseEndnotes(EMPTY_ENDNOTES)
    expect(map.size).toBe(0)
  })

  it('throws DocxParseError on malformed XML', () => {
    expect(() => parseEndnotes('<<< invalid')).toThrow(DocxParseError)
  })

  it('round-trips through the serializer without throwing (DXS-01)', () => {
    const map = parseEndnotes(FULL_ENDNOTES)
    const xml = writeEndnotesXml([...map.values()])
    expect(xml).toContain('First endnote.')
  })

  describe('table blocks (wave 1 follow-up)', () => {
    it('parses a table block instead of silently dropping it', () => {
      const map = parseEndnotes(ENDNOTES_WITH_TABLE)
      expect(map.get('1')?.blocks[0]?.kind).toBe('table')
    })

    it('round-trips a table through the serializer without throwing', () => {
      const map = parseEndnotes(ENDNOTES_WITH_TABLE)
      const xml = writeEndnotesXml([...map.values()])

      expect(xml).toContain('<w:tbl>')
      expect(xml).toContain('Endnote table cell')

      const reparsed = parseEndnotes(xml)
      expect(reparsed.get('1')?.blocks[0]?.kind).toBe('table')
    })
  })
})
