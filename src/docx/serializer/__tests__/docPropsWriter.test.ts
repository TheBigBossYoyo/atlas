/**
 * Tests for src/docx/serializer/docPropsWriter.ts (D19 / DXS-13)
 */

import { describe, expect, it } from 'vitest'

import { updateCorePropsXml } from '../docPropsWriter'

const REAL_WORD_CORE_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
  + 'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
  + 'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
  + '<dc:title>Quarterly Report</dc:title>'
  + '<dc:creator>Original Author</dc:creator>'
  + '<cp:lastModifiedBy>Original Author</cp:lastModifiedBy>'
  + '<cp:revision>3</cp:revision>'
  + '<dcterms:created xsi:type="dcterms:W3CDTF">2020-01-01T00:00:00Z</dcterms:created>'
  + '<dcterms:modified xsi:type="dcterms:W3CDTF">2020-06-15T12:30:00Z</dcterms:modified>'
  + '</cp:coreProperties>'

const FIXED_NOW = new Date('2026-09-14T08:00:00.000Z')

describe('updateCorePropsXml', () => {
  it('bumps dcterms:modified to the given date, preserving its xsi:type attribute', () => {
    const xml = updateCorePropsXml(REAL_WORD_CORE_XML, undefined, FIXED_NOW)
    expect(xml).toContain('<dcterms:modified xsi:type="dcterms:W3CDTF">2026-09-14T08:00:00Z</dcterms:modified>')
  })

  it('leaves dcterms:created and every other unrelated property untouched', () => {
    const xml = updateCorePropsXml(REAL_WORD_CORE_XML, undefined, FIXED_NOW)
    expect(xml).toContain('<dcterms:created xsi:type="dcterms:W3CDTF">2020-01-01T00:00:00Z</dcterms:created>')
    expect(xml).toContain('<dc:title>Quarterly Report</dc:title>')
    expect(xml).toContain('<dc:creator>Original Author</dc:creator>')
    expect(xml).toContain('<cp:revision>3</cp:revision>')
  })

  it('leaves cp:lastModifiedBy untouched when no name is supplied', () => {
    const xml = updateCorePropsXml(REAL_WORD_CORE_XML, undefined, FIXED_NOW)
    expect(xml).toContain('<cp:lastModifiedBy>Original Author</cp:lastModifiedBy>')
  })

  it('replaces cp:lastModifiedBy when a name is supplied', () => {
    const xml = updateCorePropsXml(REAL_WORD_CORE_XML, 'Atlas', FIXED_NOW)
    expect(xml).toContain('<cp:lastModifiedBy>Atlas</cp:lastModifiedBy>')
    expect(xml).not.toContain('Original Author</cp:lastModifiedBy>')
  })

  it('escapes XML-significant characters in the supplied lastModifiedBy', () => {
    const xml = updateCorePropsXml(REAL_WORD_CORE_XML, 'A & B <Team>', FIXED_NOW)
    expect(xml).toContain('<cp:lastModifiedBy>A &amp; B &lt;Team&gt;</cp:lastModifiedBy>')
  })

  it('inserts cp:lastModifiedBy when the source part never had one', () => {
    const withoutLastModifiedBy = REAL_WORD_CORE_XML.replace('<cp:lastModifiedBy>Original Author</cp:lastModifiedBy>', '')
    const xml = updateCorePropsXml(withoutLastModifiedBy, 'Atlas', FIXED_NOW)
    expect(xml).toContain('<cp:lastModifiedBy>Atlas</cp:lastModifiedBy>')
    // Still well-formed: the rest of the document survives.
    expect(xml).toContain('<dc:title>Quarterly Report</dc:title>')
  })

  it('builds a minimal valid core.xml when the source package had none at all', () => {
    const xml = updateCorePropsXml(undefined, 'Atlas', FIXED_NOW)
    expect(xml.startsWith('<?xml')).toBe(true)
    expect(xml).toContain('<cp:coreProperties')
    expect(xml).toContain('<dcterms:modified xsi:type="dcterms:W3CDTF">2026-09-14T08:00:00Z</dcterms:modified>')
    expect(xml).toContain('<cp:lastModifiedBy>Atlas</cp:lastModifiedBy>')
  })

  it('defaults to the current time when no date is passed', () => {
    const before = Date.now()
    const xml = updateCorePropsXml(REAL_WORD_CORE_XML, undefined)
    const after = Date.now()

    const match = xml.match(/<dcterms:modified[^>]*>([^<]+)<\/dcterms:modified>/)
    expect(match).not.toBeNull()
    const parsed = Date.parse(match![1])
    // W3CDTF drops milliseconds, so allow a small tolerance window.
    expect(parsed).toBeGreaterThanOrEqual(Math.floor(before / 1000) * 1000)
    expect(parsed).toBeLessThanOrEqual(after)
  })
})
