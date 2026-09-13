/**
 * Tests for src/docx/parser/styles.ts
 */

import { describe, expect, it } from 'vitest'

import { halfPoint, twip } from '../../model'
import { parseStyles } from '../styles'
import { DocxParseError } from '../unzip'

const BASE_STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:b/>
        <w:color w:val="445566"/>
        <w:sz w:val="24"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:before="120" w:after="240"/>
        <w:jc w:val="center"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>

  <w:style w:type="paragraph" w:styleId="BaseParagraph" w:default="1">
    <w:name w:val="Base Paragraph"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="80"/>
      <w:ind w:left="720"/>
    </w:pPr>
    <w:rPr>
      <w:i/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="Heading1" w:customStyle="1">
    <w:name w:val="Heading 1"/>
    <w:aliases w:val="Heading 1,Heading1"/>
    <w:basedOn w:val="BaseParagraph"/>
    <w:next w:val="BodyText"/>
    <w:link w:val="Heading1Char"/>
    <w:pPr>
      <w:spacing w:before="480" w:after="120"/>
      <w:keepNext/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:color w:val="AA0000"/>
    </w:rPr>
    <w:rsid w:val="deadbeef"/>
  </w:style>

  <w:style w:type="character" w:styleId="Heading1Char">
    <w:name w:val="Heading 1 Char"/>
    <w:link w:val="Heading1"/>
    <w:rPr>
      <w:u w:val="single" w:color="00AA00"/>
      <w:sz w:val="28"/>
    </w:rPr>
  </w:style>

  <w:style w:type="table" w:styleId="AtlasTable">
    <w:name w:val="Atlas Table"/>
    <w:tblPr>
      <w:tblW w:type="dxa" w:w="7200"/>
      <w:tblInd w:type="dxa" w:w="360"/>
      <w:tblLayout w:val="fixed"/>
      <w:tblLook w:val="04A0" w:firstRow="1" w:noVBand="1"/>
      <w:jc w:val="center"/>
      <w:shd w:fill="EFEFEF"/>
    </w:tblPr>
  </w:style>

  <w:style w:type="numbering" w:styleId="ListStyle">
    <w:name w:val="List Style"/>
    <w:pPr>
      <w:numPr>
        <w:ilvl w:val="2"/>
        <w:numId w:val="42"/>
      </w:numPr>
    </w:pPr>
  </w:style>
</w:styles>`

const THREE_LEVEL_CHAIN_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Base"/>
  <w:style w:type="paragraph" w:styleId="Mid">
    <w:basedOn w:val="Base"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Leaf">
    <w:basedOn w:val="Mid"/>
  </w:style>
</w:styles>`

const MISSING_STYLE_ID_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph">
    <w:name w:val="Broken"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Valid"/>
</w:styles>`

describe('parseStyles', () => {
  it('extracts docDefaults run properties', () => {
    const part = parseStyles(BASE_STYLES_XML)

    expect(part.docDefaults.rPr).toEqual({
      bold: true,
      color: '445566',
      sz: halfPoint(24),
    })
  })

  it('extracts docDefaults paragraph properties', () => {
    const part = parseStyles(BASE_STYLES_XML)

    expect(part.docDefaults.pPr).toEqual({
      spacing: {
        before: twip(120),
        after: twip(240),
      },
      jc: 'center',
    })
  })

  it('parses paragraph styles with basedOn, next, and linked style ids', () => {
    const part = parseStyles(BASE_STYLES_XML)
    const heading = part.styles.get('Heading1')

    expect(heading).toMatchObject({
      id: 'Heading1',
      type: 'paragraph',
      basedOn: 'BaseParagraph',
      next: 'BodyText',
      linked: 'Heading1Char',
    })
  })

  it('parses paragraph style paragraph and run properties', () => {
    const part = parseStyles(BASE_STYLES_XML)
    const heading = part.styles.get('Heading1')

    expect(heading?.paragraph).toEqual({
      spacing: {
        before: twip(480),
        after: twip(120),
      },
      keepNext: true,
    })
    expect(heading?.run).toEqual({
      bold: true,
      color: 'AA0000',
    })
  })

  it('parses character styles', () => {
    const part = parseStyles(BASE_STYLES_XML)
    const linked = part.styles.get('Heading1Char')

    expect(linked).toMatchObject({
      id: 'Heading1Char',
      type: 'character',
      linked: 'Heading1',
    })
    expect(linked?.run).toEqual({
      underline: {
        style: 'single',
        color: '00AA00',
      },
      sz: halfPoint(28),
    })
  })

  it('parses table styles', () => {
    const part = parseStyles(BASE_STYLES_XML)
    const table = part.styles.get('AtlasTable')

    expect(table).toMatchObject({
      id: 'AtlasTable',
      type: 'table',
    })
    expect(table?.table).toEqual({
      width: { type: 'dxa', value: twip(7200) },
      indent: { type: 'dxa', value: twip(360) },
      layout: 'fixed',
      look: {
        value: '04A0',
        firstRow: true,
        noVBand: true,
      },
      justification: 'center',
      shading: {
        fill: 'EFEFEF',
      },
    })
  })

  it('parses numbering styles', () => {
    const part = parseStyles(BASE_STYLES_XML)
    const numbering = part.styles.get('ListStyle')

    expect(numbering).toMatchObject({
      id: 'ListStyle',
      type: 'numbering',
      numbering: {
        numId: '42',
        ilvl: 2,
      },
    })
    expect(numbering?.paragraph?.numPr).toEqual({ ilvl: 2, numId: '42' })
  })

  it('splits aliases into an immutable array', () => {
    const part = parseStyles(BASE_STYLES_XML)
    expect(part.styles.get('Heading1')?.aliases).toEqual(['Heading 1', 'Heading1'])
  })

  it('parses linked style pairs in both directions', () => {
    const part = parseStyles(BASE_STYLES_XML)

    expect(part.styles.get('Heading1')?.linked).toBe('Heading1Char')
    expect(part.styles.get('Heading1Char')?.linked).toBe('Heading1')
  })

  it('parses default-style and custom-style flags', () => {
    const part = parseStyles(BASE_STYLES_XML)

    expect(part.styles.get('BaseParagraph')?.isDefault).toBe(true)
    expect(part.styles.get('Heading1')?.custom).toBe(true)
  })

  it('retains three-level basedOn chains', () => {
    const part = parseStyles(THREE_LEVEL_CHAIN_XML)

    expect(part.styles.get('Mid')?.basedOn).toBe('Base')
    expect(part.styles.get('Leaf')?.basedOn).toBe('Mid')
  })

  it('skips styles with a missing styleId gracefully', () => {
    const part = parseStyles(MISSING_STYLE_ID_XML)

    expect(part.styles.size).toBe(1)
    expect(part.styles.has('')).toBe(false)
    expect(part.styles.get('Valid')).toMatchObject({ id: 'Valid', type: 'paragraph' })
  })

  it('ignores unsupported style children without failing', () => {
    const part = parseStyles(BASE_STYLES_XML)

    expect(part.styles.get('Heading1')?.name).toBe('Heading 1')
    expect(part.styles.get('Heading1')?.run?.bold).toBe(true)
  })

  describe('malformed XML handling (D28 / DXP-16)', () => {
    it('wraps a fast-xml-parser failure in DocxParseError instead of letting it propagate raw', () => {
      expect(() => parseStyles('<<< not xml <<<')).toThrow(DocxParseError)
    })
  })
})
