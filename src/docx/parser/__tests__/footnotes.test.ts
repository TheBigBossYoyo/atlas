/**
 * Tests for src/docx/parser/footnotes.ts
 */

import { describe, it, expect } from 'vitest'
import { parseFootnotes } from '../footnotes'
import { DocxParseError } from '../unzip'
import { writeFootnotesXml } from '../../serializer/footnotesWriter'

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

const FULL_FOOTNOTES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0">
    <w:p><w:r><w:continuationSeparator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:id="1">
    <w:p><w:r><w:t>First footnote text.</w:t></w:r></w:p>
  </w:footnote>
  <w:footnote w:id="2">
    <w:p><w:r><w:t>Second footnote text.</w:t></w:r></w:p>
  </w:footnote>
</w:footnotes>`

const SINGLE_FOOTNOTE = `<?xml version="1.0" encoding="UTF-8"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1">
    <w:p><w:r><w:t>Only note.</w:t></w:r></w:p>
  </w:footnote>
</w:footnotes>`

const EMPTY_FOOTNOTES = `<?xml version="1.0" encoding="UTF-8"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
</w:footnotes>`

// wave 1 follow-up: one footnote among several contains a table — proves
// both that the table survives (rather than being silently dropped) and
// that per-id content extraction isn't confused by a sibling footnote's
// table when correlating ids to bodies.
const FOOTNOTES_WITH_TABLE = `<?xml version="1.0" encoding="UTF-8"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1">
    <w:p><w:r><w:t>Plain footnote.</w:t></w:r></w:p>
  </w:footnote>
  <w:footnote w:id="2">
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>Table footnote cell</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
  </w:footnote>
  <w:footnote w:id="3">
    <w:p><w:r><w:t>Another plain footnote.</w:t></w:r></w:p>
  </w:footnote>
</w:footnotes>`

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('parseFootnotes', () => {
  it('parses all footnote entries including separators', () => {
    const map = parseFootnotes(FULL_FOOTNOTES)
    expect(map.size).toBe(4)
  })

  it('keys are footnote id strings', () => {
    const map = parseFootnotes(FULL_FOOTNOTES)
    expect(map.has('-1')).toBe(true)
    expect(map.has('0')).toBe(true)
    expect(map.has('1')).toBe(true)
    expect(map.has('2')).toBe(true)
  })

  it('separator footnote carries noteType="separator"', () => {
    const map = parseFootnotes(FULL_FOOTNOTES)
    expect(map.get('-1')?.noteType).toBe('separator')
  })

  it('continuationSeparator footnote carries correct noteType', () => {
    const map = parseFootnotes(FULL_FOOTNOTES)
    expect(map.get('0')?.noteType).toBe('continuationSeparator')
  })

  it('normal footnote has kind="footnote" and no noteType', () => {
    const map = parseFootnotes(FULL_FOOTNOTES)
    const note = map.get('1')
    expect(note?.kind).toBe('footnote')
    expect(note?.noteType).toBeUndefined()
  })

  it('parses each footnote body into a real paragraph block, not an opaque UnknownNode', () => {
    const map = parseFootnotes(SINGLE_FOOTNOTE)
    const note = map.get('1')
    expect(note?.blocks).toHaveLength(1)
    expect(note?.blocks[0].kind).toBe('paragraph')
  })

  it('extracts the run text from the parsed footnote paragraph', () => {
    const map = parseFootnotes(SINGLE_FOOTNOTE)
    const block = map.get('1')?.blocks[0]
    expect(block?.kind).toBe('paragraph')
    if (block?.kind === 'paragraph') {
      const run = block.children[0]
      expect(run.kind).toBe('run')
      if (run.kind === 'run') {
        expect(run.children[0]).toEqual({ kind: 'text', value: 'Only note.' })
      }
    }
  })

  it('returns empty map for empty <w:footnotes> element', () => {
    const map = parseFootnotes(EMPTY_FOOTNOTES)
    expect(map.size).toBe(0)
  })

  it('handles a single (non-array) footnote element', () => {
    const map = parseFootnotes(SINGLE_FOOTNOTE)
    expect(map.size).toBe(1)
    expect(map.has('1')).toBe(true)
  })

  it('throws DocxParseError on malformed XML', () => {
    expect(() => parseFootnotes('<<< invalid')).toThrow(DocxParseError)
  })

  it('round-trips through the serializer without throwing (DXS-01)', () => {
    const map = parseFootnotes(FULL_FOOTNOTES)
    const xml = writeFootnotesXml([...map.values()])
    expect(xml).toContain('First footnote text.')
    expect(xml).toContain('Second footnote text.')
  })

  describe('table blocks (wave 1 follow-up)', () => {
    it('parses a table block instead of silently dropping it, without disturbing sibling footnotes', () => {
      const map = parseFootnotes(FOOTNOTES_WITH_TABLE)

      expect(map.get('1')?.blocks[0]?.kind).toBe('paragraph')
      expect(map.get('2')?.blocks[0]?.kind).toBe('table')
      expect(map.get('3')?.blocks[0]?.kind).toBe('paragraph')
    })

    it('round-trips a table through the serializer without throwing', () => {
      const map = parseFootnotes(FOOTNOTES_WITH_TABLE)
      const xml = writeFootnotesXml([...map.values()])

      expect(xml).toContain('<w:tbl>')
      expect(xml).toContain('Table footnote cell')

      const reparsed = parseFootnotes(xml)
      expect(reparsed.get('2')?.blocks[0]?.kind).toBe('table')
    })
  })
})
