import { describe, expect, it } from 'vitest'

import { eighthPoint, halfPoint, hexColor, twip, type Style } from '../../model'
import { parseStyles, type StylesPart } from '../../parser/styles'
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
    // w:keepNext precedes w:spacing per CT_PPrBase's schema order.
    expect(xml).toContain('<w:pPr><w:keepNext/><w:spacing w:before="480" w:after="120"/></w:pPr>')
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
    // w:sz precedes w:u per CT_RPr's schema order.
    expect(xml).toContain('<w:rPr><w:sz w:val="28"/><w:u w:val="single" w:color="00AA00"/></w:rPr>')
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
    // Order follows CT_TblPrBase: tblW, jc, tblInd, shd, tblLayout, tblLook.
    expect(xml).toContain(
      '<w:tblPr><w:tblW w:type="dxa" w:w="7200"/><w:jc w:val="center"/><w:tblInd w:type="dxa" w:w="360"/><w:shd w:fill="EFEFEF"/><w:tblLayout w:val="fixed"/><w:tblLook w:val="04A0" w:firstRow="1" w:noVBand="1"/></w:tblPr>',
    )
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

  describe('table conditional formatting (D7 / DXP-06, DXL-08, DXS-05)', () => {
    it('emits one w:tblStylePr per conditional format, plus row/col band sizes', () => {
      const xml = writeStylesXml(
        makeStylesPart([
          {
            id: 'AtlasBandedTable',
            type: 'table',
            basedOn: 'TableNormal',
            table: { rowBandSize: 1, colBandSize: 1 },
            conditionalFormats: new Map([
              [
                'firstRow',
                {
                  run: { bold: true, color: hexColor('FFFFFF') },
                  cell: { shd: { pattern: 'clear', color: 'auto', fill: hexColor('4472C4') } },
                },
              ],
              [
                'band1Horz',
                { cell: { shd: { pattern: 'clear', color: 'auto', fill: hexColor('D9E2F3') } } },
              ],
            ]),
          },
        ]),
      )

      expect(xml).toContain('<w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/>')
      expect(xml).toContain(
        '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>'
          + '<w:tcPr><w:shd w:fill="4472C4" w:color="auto" w:val="clear"/></w:tcPr></w:tblStylePr>',
      )
      expect(xml).toContain(
        '<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:fill="D9E2F3" w:color="auto" w:val="clear"/></w:tcPr></w:tblStylePr>',
      )
    })

    it('omits w:tblStylePr entirely when a style has no conditional formats', () => {
      const xml = writeStylesXml(makeStylesPart([{ id: 'Plain', type: 'table' }]))
      expect(xml).not.toContain('w:tblStylePr')
    })
  })

  describe('w:latentStyles passthrough (D19 / DXS-14)', () => {
    it('re-emits the captured latentStyles node between docDefaults and the first style', () => {
      const part = makeStylesPart([{ id: 'Base', type: 'paragraph' }])
      const xml = writeStylesXml({
        ...part,
        latentStyles: {
          '@_w:defLockedState': '0',
          'w:lsdException': { '@_w:name': 'Normal', '@_w:uiPriority': '0' },
        },
      })

      const docDefaultsIndex = xml.indexOf('</w:docDefaults>')
      const latentStylesIndex = xml.indexOf('<w:latentStyles')
      const firstStyleIndex = xml.indexOf('<w:style ')

      expect(latentStylesIndex).toBeGreaterThan(docDefaultsIndex)
      expect(firstStyleIndex).toBeGreaterThan(latentStylesIndex)
      expect(xml).toContain('<w:latentStyles w:defLockedState="0"><w:lsdException w:name="Normal" w:uiPriority="0"/></w:latentStyles>')
    })

    it('omits w:latentStyles when the part has none', () => {
      const xml = writeStylesXml(makeStylesPart([]))
      expect(xml).not.toContain('w:latentStyles')
    })
  })

  // DOCX-1 — same theme-font gap as `documentWriter.ts`'s `buildFontSetElement`,
  // for the styles.xml-side builder (`w:docDefaults`/named-style `w:rFonts`).
  describe('w:rFonts theme references (DOCX-1)', () => {
    it('emits a theme-only rFonts (no literal font name) instead of dropping it', () => {
      const xml = writeStylesXml(
        makeStylesPart([
          {
            id: 'Heading1',
            type: 'paragraph',
            run: {
              rFonts: {
                asciiTheme: 'majorHAnsi',
                hAnsiTheme: 'majorHAnsi',
                csTheme: 'majorBidi',
                eastAsiaTheme: 'majorEastAsia',
              },
            },
          },
        ]),
      )

      expect(xml).toContain(
        '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi" w:eastAsiaTheme="majorEastAsia"/>',
      )
    })

    it('emits a literal-only rFonts unchanged', () => {
      const xml = writeStylesXml(
        makeStylesPart([
          { id: 'Heading1', type: 'paragraph', run: { rFonts: { ascii: 'Calibri', hAnsi: 'Calibri' } } },
        ]),
      )

      expect(xml).toContain('<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>')
    })

    it('emits a mixed literal + theme rFonts', () => {
      const xml = writeStylesXml(
        makeStylesPart([
          {
            id: 'Heading1',
            type: 'paragraph',
            run: { rFonts: { ascii: 'Calibri', hAnsiTheme: 'minorHAnsi' } },
          },
        ]),
      )

      expect(xml).toContain('<w:rFonts w:ascii="Calibri" w:hAnsiTheme="minorHAnsi"/>')
    })

    it('round-trips a docDefaults theme-only rFonts through parseStyles + writeStylesXml', () => {
      const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        + '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi"/></w:rPr></w:rPrDefault></w:docDefaults>'
        + '</w:styles>'

      const written = writeStylesXml(parseStyles(xml))

      expect(written).toContain('<w:rFonts w:asciiTheme="minorHAnsi"/>')
    })
  })

  // DOCX-12 — the six newly-modeled character effects/scaling, mirrored from
  // `documentWriter.ts`'s `buildRunPropertiesNode` into this styles.xml-side
  // builder so a named style (not just direct run formatting) round-trips
  // them too.
  describe('w:rPr character effects (DOCX-12)', () => {
    it('emits outline, emboss, imprint, bdr, em, and w (character scale)', () => {
      const xml = writeStylesXml(
        makeStylesPart([
          {
            id: 'Effects',
            type: 'character',
            run: {
              outline: true,
              emboss: true,
              imprint: true,
              em: 'dot',
              charScale: 150,
              bdr: {
                style: 'single',
                size: eighthPoint(4),
                space: twip(1),
                color: hexColor('FF0000'),
              },
            },
          },
        ]),
      )

      expect(xml).toContain('<w:outline/>')
      expect(xml).toContain('<w:emboss/>')
      expect(xml).toContain('<w:imprint/>')
      expect(xml).toContain('<w:bdr w:val="single" w:color="FF0000" w:sz="4" w:space="1"/>')
      expect(xml).toContain('<w:em w:val="dot"/>')
      expect(xml).toContain('<w:w w:val="150"/>')
    })

    it('round-trips a named style carrying these effects through parseStyles + writeStylesXml', () => {
      const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        + '<w:style w:type="character" w:styleId="Effects">'
        + '<w:rPr><w:outline/><w:emboss/><w:imprint/><w:em w:val="comma"/><w:w w:val="80"/></w:rPr>'
        + '</w:style>'
        + '</w:styles>'

      const written = writeStylesXml(parseStyles(xml))

      expect(written).toContain('<w:outline/>')
      expect(written).toContain('<w:emboss/>')
      expect(written).toContain('<w:imprint/>')
      expect(written).toContain('<w:em w:val="comma"/>')
      expect(written).toContain('<w:w w:val="80"/>')
    })
  })
})
