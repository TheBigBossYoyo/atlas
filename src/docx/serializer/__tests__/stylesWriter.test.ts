import { describe, expect, it } from 'vitest'

import { halfPoint, hexColor, twip, type Style } from '../../model'
import type { StylesPart } from '../../parser/styles'
import { writeStylesXml } from '../stylesWriter'

function makeStylesPart(styles: ReadonlyArray<Style>): StylesPart {
  return {
    docDefaults: {
      rPr: {
        bold: true,
        color: hexColor('445566'),
        sz: halfPoint(24),
      },
      pPr: {
        spacing: {
          before: twip(120),
          after: twip(240),
        },
        jc: 'center',
      },
    },
    styles: new Map(styles.map((style) => [style.id, style])),
  }
}

describe('writeStylesXml', () => {
  it('emits the XML declaration and docDefaults block', () => {
    const xml = writeStylesXml(makeStylesPart([]))

    expect(xml.startsWith('<?xml')).toBe(true)
    expect(xml).toContain('<w:styles ')
    expect(xml).toContain('<w:docDefaults>')
    expect(xml).toContain('<w:rPrDefault><w:rPr><w:b/><w:color w:val="445566"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>')
    expect(xml).toContain('<w:pPrDefault><w:pPr><w:spacing w:before="120" w:after="240"/><w:jc w:val="center"/></w:pPr></w:pPrDefault>')
  })

  it('writes paragraph styles with basedOn, next, and linked style ids', () => {
    const xml = writeStylesXml(
      makeStylesPart([
        {
          id: 'Heading1',
          type: 'paragraph',
          basedOn: 'BaseParagraph',
          next: 'BodyText',
          linked: 'Heading1Char',
          paragraph: {
            spacing: {
              before: twip(480),
              after: twip(120),
            },
            keepNext: true,
          },
          run: {
            bold: true,
            color: hexColor('AA0000'),
          },
        },
      ]),
    )

    expect(xml).toContain('<w:style w:type="paragraph" w:styleId="Heading1">')
    expect(xml).toContain('<w:basedOn w:val="BaseParagraph"/>')
    expect(xml).toContain('<w:next w:val="BodyText"/>')
    expect(xml).toContain('<w:link w:val="Heading1Char"/>')
    expect(xml).toContain('<w:pPr><w:spacing w:before="480" w:after="120"/><w:keepNext/></w:pPr>')
    expect(xml).toContain('<w:rPr><w:b/><w:color w:val="AA0000"/></w:rPr>')
  })

  it('writes character styles with run properties', () => {
    const xml = writeStylesXml(
      makeStylesPart([
        {
          id: 'Heading1Char',
          type: 'character',
          linked: 'Heading1',
          run: {
            underline: {
              style: 'single',
              color: hexColor('00AA00'),
            },
            sz: halfPoint(28),
          },
        },
      ]),
    )

    expect(xml).toContain('<w:style w:type="character" w:styleId="Heading1Char">')
    expect(xml).toContain('<w:link w:val="Heading1"/>')
    expect(xml).toContain('<w:rPr><w:u w:val="single" w:color="00AA00"/><w:sz w:val="28"/></w:rPr>')
  })

  it('writes table styles with table properties', () => {
    const xml = writeStylesXml(
      makeStylesPart([
        {
          id: 'AtlasTable',
          type: 'table',
          table: {
            width: { type: 'dxa', value: twip(7200) },
            indent: { type: 'dxa', value: twip(360) },
            layout: 'fixed',
            look: {
              value: '04A0',
              firstRow: true,
              noVBand: true,
            },
            justification: 'center',
            shading: { fill: hexColor('EFEFEF') },
          },
        },
      ]),
    )

    expect(xml).toContain('<w:style w:type="table" w:styleId="AtlasTable">')
    expect(xml).toContain('<w:tblPr><w:tblW w:type="dxa" w:w="7200"/><w:tblInd w:type="dxa" w:w="360"/><w:tblLayout w:val="fixed"/><w:tblLook w:val="04A0" w:firstRow="1" w:noVBand="1"/><w:jc w:val="center"/><w:shd w:fill="EFEFEF"/></w:tblPr>')
  })

  it('writes numbering styles from the numbering shortcut model', () => {
    const xml = writeStylesXml(
      makeStylesPart([
        {
          id: 'ListStyle',
          type: 'numbering',
          numbering: {
            numId: '42',
            ilvl: 2,
          },
        },
      ]),
    )

    expect(xml).toContain('<w:style w:type="numbering" w:styleId="ListStyle">')
    expect(xml).toContain('<w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="42"/></w:numPr></w:pPr>')
  })

  it('preserves basedOn chains across multiple styles', () => {
    const xml = writeStylesXml(
      makeStylesPart([
        { id: 'Base', type: 'paragraph' },
        { id: 'Mid', type: 'paragraph', basedOn: 'Base' },
        { id: 'Leaf', type: 'paragraph', basedOn: 'Mid' },
      ]),
    )

    expect(xml).toContain('<w:style w:type="paragraph" w:styleId="Mid"><w:basedOn w:val="Base"/></w:style>')
    expect(xml).toContain('<w:style w:type="paragraph" w:styleId="Leaf"><w:basedOn w:val="Mid"/></w:style>')
  })
})
