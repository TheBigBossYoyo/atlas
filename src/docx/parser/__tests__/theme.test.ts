/**
 * Tests for src/docx/parser/theme.ts
 *
 * Uses an inline Office theme1.xml fixture that matches the real schema
 * produced by Word (Office Open XML DrawingML theme namespace "a:").
 */

import { describe, it, expect } from 'vitest'
import { parseTheme } from '../theme'
import { DocxParseError } from '../unzip'

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

/** A realistic Office theme1.xml covering all colour slot types. */
const FULL_THEME = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A9D18E"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`

const THEME_MISSING_ROOT = `<?xml version="1.0"?>
<notATheme/>`

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('parseTheme', () => {
  it('parses major latin font correctly', () => {
    const theme = parseTheme(FULL_THEME)
    expect(theme.fontScheme.major.latin).toBe('Calibri Light')
  })

  it('parses minor latin font correctly', () => {
    const theme = parseTheme(FULL_THEME)
    expect(theme.fontScheme.minor.latin).toBe('Calibri')
  })

  it('resolves srgbClr accent1 to correct hex', () => {
    const theme = parseTheme(FULL_THEME)
    expect(theme.colorScheme.get('accent1')).toBe('4472C4')
  })

  it('resolves sysClr dk1 via lastClr fallback', () => {
    const theme = parseTheme(FULL_THEME)
    expect(theme.colorScheme.get('dk1')).toBe('000000')
  })

  it('resolves sysClr lt1 via lastClr fallback', () => {
    const theme = parseTheme(FULL_THEME)
    expect(theme.colorScheme.get('lt1')).toBe('FFFFFF')
  })

  it('resolves all 12 colour slots', () => {
    const theme = parseTheme(FULL_THEME)
    const slots = ['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink']
    for (const s of slots) {
      expect(theme.colorScheme.has(s)).toBe(true)
    }
    expect(theme.colorScheme.size).toBe(12)
  })

  it('colorScheme values are HexColor branded strings', () => {
    const theme = parseTheme(FULL_THEME)
    expect(typeof theme.colorScheme.get('accent1')).toBe('string')
  })

  it('throws DocxParseError for completely invalid XML', () => {
    expect(() => parseTheme('<<< not xml')).toThrow(DocxParseError)
  })

  it('throws DocxParseError when <a:theme> root element is missing', () => {
    expect(() => parseTheme(THEME_MISSING_ROOT)).toThrow(DocxParseError)
  })

  it('returns empty colorScheme when <a:clrScheme> is absent', () => {
    const minimal = `<?xml version="1.0"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
  <a:themeElements>
    <a:fontScheme name="T">
      <a:majorFont><a:latin typeface="Times"/></a:majorFont>
      <a:minorFont><a:latin typeface="Arial"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`
    const theme = parseTheme(minimal)
    expect(theme.colorScheme.size).toBe(0)
  })

  it('returns empty latin when <a:fontScheme> is absent', () => {
    const minimal = `<?xml version="1.0"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
  <a:themeElements/>
</a:theme>`
    const theme = parseTheme(minimal)
    expect(theme.fontScheme.major.latin).toBe('')
    expect(theme.fontScheme.minor.latin).toBe('')
  })
})
