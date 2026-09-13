/**
 * Tests for src/docx/parser/numbering.ts
 */

import { describe, expect, it } from 'vitest'

import { halfPoint, twip } from '../../model'
import { parseNumbering } from '../numbering'
import { DocxParseError } from '../unzip'

const NUMBERING_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:b/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%2)"/>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:numFmt w:val="upperRoman"/>
      <w:lvlText w:val="%1.%2"/>
    </w:lvl>
    <w:lvl w:ilvl="3">
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
      <w:rPr>
        <w:sz w:val="20"/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="4"><w:numFmt w:val="decimal"/><w:lvlText w:val="%5."/></w:lvl>
    <w:lvl w:ilvl="5"><w:numFmt w:val="decimal"/><w:lvlText w:val="%6."/></w:lvl>
    <w:lvl w:ilvl="6"><w:numFmt w:val="decimal"/><w:lvlText w:val="%7."/></w:lvl>
    <w:lvl w:ilvl="7"><w:numFmt w:val="decimal"/><w:lvlText w:val="%8."/></w:lvl>
    <w:lvl w:ilvl="8"><w:numFmt w:val="decimal"/><w:lvlText w:val="%9."/></w:lvl>
  </w:abstractNum>

  <w:abstractNum w:abstractNumId="1">
    <w:styleLink w:val="ListBullet"/>
    <w:numStyleLink w:val="ListBulletStyle"/>
    <w:lvl w:ilvl="0" w:tentative="1">
      <w:start w:val="3"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1)"/>
      <w:suff w:val="space"/>
      <w:isLgl/>
    </w:lvl>
  </w:abstractNum>

  <w:num w:numId="10">
    <w:abstractNumId w:val="0"/>
    <w:lvlOverride w:ilvl="0">
      <w:startOverride w:val="5"/>
    </w:lvlOverride>
  </w:num>

  <w:num w:numId="11">
    <w:abstractNumId w:val="0"/>
  </w:num>

  <w:num w:numId="12">
    <w:abstractNumId w:val="1"/>
    <w:lvlOverride w:ilvl="1">
      <w:startOverride w:val="2"/>
      <w:lvl w:ilvl="1">
        <w:numFmt w:val="decimal"/>
        <w:lvlText w:val="override-%2"/>
      </w:lvl>
    </w:lvlOverride>
  </w:num>
</w:numbering>`

describe('parseNumbering', () => {
  it('parses an abstract numbering definition with nine levels', () => {
    const part = parseNumbering(NUMBERING_XML)

    expect(part.abstractNums.get('0')?.levels.size).toBe(9)
  })

  it('parses decimal, lowerLetter, upperRoman, and bullet formats', () => {
    const levels = parseNumbering(NUMBERING_XML).abstractNums.get('0')?.levels

    expect(levels?.get(0)?.format).toBe('decimal')
    expect(levels?.get(1)?.format).toBe('lowerLetter')
    expect(levels?.get(2)?.format).toBe('upperRoman')
    expect(levels?.get(3)?.format).toBe('bullet')
  })

  it('parses single-level lvlText placeholders', () => {
    const level = parseNumbering(NUMBERING_XML).abstractNums.get('0')?.levels.get(0)

    expect(level?.text).toEqual({
      value: '%1.',
      placeholders: [1],
    })
  })

  it('parses multi-level lvlText placeholders', () => {
    const level = parseNumbering(NUMBERING_XML).abstractNums.get('0')?.levels.get(2)

    expect(level?.text).toEqual({
      value: '%1.%2',
      placeholders: [1, 2],
    })
  })

  it('parses level paragraph and run properties', () => {
    const level = parseNumbering(NUMBERING_XML).abstractNums.get('0')?.levels.get(0)

    expect(level?.justification).toBe('start')
    expect(level?.paragraph?.ind).toEqual({
      left: twip(720),
      hanging: twip(360),
    })
    expect(level?.run).toEqual({ bold: true })
  })

  it('parses num instances with start overrides', () => {
    const num = parseNumbering(NUMBERING_XML).nums.get('10')

    expect(num?.abstractNumId).toBe('0')
    expect(num?.levelOverrides?.get(0)?.startOverride).toBe(5)
  })

  it('parses nested level definitions inside overrides', () => {
    const override = parseNumbering(NUMBERING_XML).nums.get('12')?.levelOverrides?.get(1)

    expect(override?.startOverride).toBe(2)
    expect(override?.levelDefinition).toMatchObject({
      level: 1,
      format: 'decimal',
      text: {
        value: 'override-%2',
        placeholders: [2],
      },
    })
  })

  it('supports multiple num instances that point to the same abstractNum', () => {
    const part = parseNumbering(NUMBERING_XML)

    expect(part.nums.get('10')?.abstractNumId).toBe('0')
    expect(part.nums.get('11')?.abstractNumId).toBe('0')
  })

  it('parses style links, suffixes, tentative, legal, and run props', () => {
    const abstractNum = parseNumbering(NUMBERING_XML).abstractNums.get('1')

    expect(abstractNum).toMatchObject({
      abstractNumId: '1',
      styleLink: 'ListBullet',
      numberStyleLink: 'ListBulletStyle',
    })
    expect(abstractNum?.levels.get(0)).toMatchObject({
      start: 3,
      suffix: 'space',
      tentative: true,
      legal: true,
    })
  })

  it('parses bullet run sizing', () => {
    const level = parseNumbering(NUMBERING_XML).abstractNums.get('0')?.levels.get(3)

    expect(level?.run).toEqual({
      sz: halfPoint(20),
    })
  })

  describe('malformed XML handling (D28 / DXP-16)', () => {
    it('wraps a fast-xml-parser failure in DocxParseError instead of letting it propagate raw', () => {
      expect(() => parseNumbering('<<< not xml <<<')).toThrow(DocxParseError)
    })
  })
})
