/**
 * Tests for src/docx/parser/xmlSizeGuard.ts (D21 / DXP-20)
 */

import { describe, it, expect } from 'vitest'

import { assertXmlPartSizeWithinLimit, MAX_XML_PART_LENGTH } from '../xmlSizeGuard'
import { DocxParseError } from '../unzip'

describe('assertXmlPartSizeWithinLimit', () => {
  it('does not throw for a small XML string', () => {
    expect(() => assertXmlPartSizeWithinLimit('<w:document/>', 'word/document.xml')).not.toThrow()
  })

  it('does not throw for a string exactly at the limit', () => {
    const xml = 'a'.repeat(MAX_XML_PART_LENGTH)
    expect(() => assertXmlPartSizeWithinLimit(xml, 'word/document.xml')).not.toThrow()
  })

  it('throws DocxParseError for a string one character over the limit', () => {
    const xml = 'a'.repeat(MAX_XML_PART_LENGTH + 1)
    expect(() => assertXmlPartSizeWithinLimit(xml, 'word/document.xml')).toThrow(DocxParseError)
  })

  it('includes the part label and size in the error message', () => {
    const xml = 'a'.repeat(MAX_XML_PART_LENGTH + 1)
    try {
      assertXmlPartSizeWithinLimit(xml, 'word/document.xml')
      expect.unreachable('expected assertXmlPartSizeWithinLimit to throw')
    } catch (error) {
      expect(error).toBeInstanceOf(DocxParseError)
      expect((error as DocxParseError).message).toContain('word/document.xml')
      expect((error as DocxParseError).path).toBe('word/document.xml')
    }
  })
})
