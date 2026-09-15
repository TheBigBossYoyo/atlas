/**
 * Tests for src/docx/fonts/fontTable.ts (DEFER-4 / DXP-13)
 */

import { describe, expect, it } from 'vitest'

import { DocxParseError } from '../../parser/unzip'
import { parseFontTable } from '../fontTable'

describe('parseFontTable', () => {
  it('returns an empty array for a font table with no embedded fonts', () => {
    const xml = `<?xml version="1.0"?>
      <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/></w:font>
        <w:font w:name="Times New Roman"/>
      </w:fonts>`

    expect(parseFontTable(xml)).toEqual([{ name: 'Calibri' }, { name: 'Times New Roman' }])
  })

  it('parses all four embed variants with their relationship id and font key', () => {
    const xml = `<?xml version="1.0"?>
      <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:font w:name="Custom Font">
          <w:embedRegular r:id="rId1" w:fontKey="{11111111-1111-1111-1111-111111111111}"/>
          <w:embedBold r:id="rId2" w:fontKey="{22222222-2222-2222-2222-222222222222}"/>
          <w:embedItalic r:id="rId3" w:fontKey="{33333333-3333-3333-3333-333333333333}"/>
          <w:embedBoldItalic r:id="rId4" w:fontKey="{44444444-4444-4444-4444-444444444444}"/>
        </w:font>
      </w:fonts>`

    const [entry] = parseFontTable(xml)
    expect(entry?.name).toBe('Custom Font')
    expect(entry?.embedRegular).toEqual({ relId: 'rId1', fontKey: '{11111111-1111-1111-1111-111111111111}' })
    expect(entry?.embedBold).toEqual({ relId: 'rId2', fontKey: '{22222222-2222-2222-2222-222222222222}' })
    expect(entry?.embedItalic).toEqual({ relId: 'rId3', fontKey: '{33333333-3333-3333-3333-333333333333}' })
    expect(entry?.embedBoldItalic).toEqual({
      relId: 'rId4',
      fontKey: '{44444444-4444-4444-4444-444444444444}',
    })
  })

  it('parses w:subsetted as a boolean', () => {
    const xml = `<?xml version="1.0"?>
      <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:font w:name="Subset Font">
          <w:embedRegular r:id="rId1" w:fontKey="{11111111-1111-1111-1111-111111111111}" w:subsetted="1"/>
        </w:font>
      </w:fonts>`

    const [entry] = parseFontTable(xml)
    expect(entry?.embedRegular?.subsetted).toBe(true)
  })

  it('handles a single <w:font> without wrapping it in an array', () => {
    const xml = `<?xml version="1.0"?>
      <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:font w:name="Solo"/>
      </w:fonts>`

    expect(parseFontTable(xml)).toEqual([{ name: 'Solo' }])
  })

  it('skips a <w:font> with no w:name attribute', () => {
    const xml = `<?xml version="1.0"?>
      <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:font/>
        <w:font w:name="Named"/>
      </w:fonts>`

    expect(parseFontTable(xml)).toEqual([{ name: 'Named' }])
  })

  it('returns an empty array when the document has no <w:fonts> root', () => {
    expect(parseFontTable('<?xml version="1.0"?><w:other/>')).toEqual([])
  })

  it('throws DocxParseError on malformed XML', () => {
    expect(() => parseFontTable('<w:fonts><w:font')).toThrow(DocxParseError)
  })
})
