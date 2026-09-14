import { XMLParser } from 'fast-xml-parser'
import { describe, expect, it } from 'vitest'

import {
  type Block,
  type Hyperlink,
  type Paragraph,
  type Run,
  type Table,
} from '../../model'
import { DocxParseError, parseDocument } from '..'
import { MAX_XML_PART_LENGTH } from '../xmlSizeGuard'

const DOCX_NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
].join(' ')

function documentXml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DOCX_NAMESPACES}>
  <w:body>
    ${body}
  </w:body>
</w:document>`
}

function parseBody(body: string) {
  return parseDocument(documentXml(body))
}

function expectParagraph(block: Block): Paragraph {
  expect(block.kind).toBe('paragraph')
  if (block.kind !== 'paragraph') {
    throw new Error('Expected paragraph block')
  }
  return block
}

function expectTable(block: Block): Table {
  expect(block.kind).toBe('table')
  if (block.kind !== 'table') {
    throw new Error('Expected table block')
  }
  return block
}

function expectRun(child: Paragraph['children'][number] | Hyperlink['children'][number]): Run {
  expect(child.kind).toBe('run')
  if (child.kind !== 'run') {
    throw new Error('Expected run child')
  }
  return child
}

describe('parseDocument', () => {
  it('parses a single paragraph with text', () => {
    const document = parseBody('<w:p><w:r><w:t>Hello Atlas</w:t></w:r></w:p>')

    expect(document.sections).toHaveLength(1)
    expect(document.sections[0].blocks).toHaveLength(1)

    const paragraph = expectParagraph(document.sections[0].blocks[0])
    const run = expectRun(paragraph.children[0])

    expect(run.children).toEqual([{ kind: 'text', value: 'Hello Atlas' }])
  })

  it('parses bold, italic, and underline run props', () => {
    const document = parseBody(`
      <w:p>
        <w:r>
          <w:rPr>
            <w:b/>
            <w:i/>
            <w:u w:val="single" w:color="FF0000"/>
          </w:rPr>
          <w:t>Styled</w:t>
        </w:r>
      </w:p>
    `)

    const paragraph = expectParagraph(document.sections[0].blocks[0])
    const run = expectRun(paragraph.children[0])

    expect(run.props).toMatchObject({
      bold: true,
      italic: true,
      underline: {
        style: 'single',
        color: 'FF0000',
      },
    })
  })

  it('parses tabs and break variants', () => {
    const document = parseBody(`
      <w:p>
        <w:r>
          <w:tab/>
          <w:br/>
          <w:br w:type="page"/>
          <w:br w:type="column"/>
          <w:br w:type="textWrapping" w:clear="all"/>
        </w:r>
      </w:p>
    `)

    const paragraph = expectParagraph(document.sections[0].blocks[0])
    const run = expectRun(paragraph.children[0])

    expect(run.children).toEqual([
      { kind: 'tab' },
      { kind: 'break', breakType: 'line' },
      { kind: 'break', breakType: 'page' },
      { kind: 'break', breakType: 'column' },
      { kind: 'break', breakType: 'textWrapping', clear: 'all' },
    ])
  })

  it('parses table, row, cell props, and nested 2x2 tables', () => {
    const document = parseBody(`
      <w:tbl>
        <w:tblPr>
          <w:tblStyle w:val="GridTable5Dark"/>
          <w:tblW w:type="dxa" w:w="5000"/>
          <w:tblInd w:type="dxa" w:w="720"/>
          <w:tblLayout w:type="fixed"/>
          <w:jc w:val="center"/>
        </w:tblPr>
        <w:tr>
          <w:trPr>
            <w:trHeight w:val="360" w:hRule="exact"/>
            <w:cantSplit/>
            <w:tblHeader/>
          </w:trPr>
          <w:tc>
            <w:tcPr>
              <w:tcW w:type="dxa" w:w="2400"/>
              <w:gridSpan w:val="1"/>
              <w:vAlign w:val="center"/>
            </w:tcPr>
            <w:p><w:r><w:t>A1</w:t></w:r></w:p>
          </w:tc>
          <w:tc>
            <w:p><w:r><w:t>A2</w:t></w:r></w:p>
          </w:tc>
        </w:tr>
        <w:tr>
          <w:tc>
            <w:p><w:r><w:t>B1</w:t></w:r></w:p>
          </w:tc>
          <w:tc>
            <w:tcPr>
              <w:vMerge w:val="restart"/>
              <w:noWrap/>
              <w:hideMark/>
            </w:tcPr>
            <w:tbl>
              <w:tr>
                <w:tc><w:p><w:r><w:t>I1</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>I2</w:t></w:r></w:p></w:tc>
              </w:tr>
              <w:tr>
                <w:tc><w:p><w:r><w:t>I3</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>I4</w:t></w:r></w:p></w:tc>
              </w:tr>
            </w:tbl>
          </w:tc>
        </w:tr>
      </w:tbl>
    `)

    const table = expectTable(document.sections[0].blocks[0])

    expect(table.props).toMatchObject({
      tblStyle: 'GridTable5Dark',
      tblW: { type: 'dxa', value: 5000 },
      tblInd: { type: 'dxa', value: 720 },
      tblLayout: 'fixed',
      jc: 'center',
    })
    expect(table.rows).toHaveLength(2)

    const firstRow = table.rows[0]
    expect(firstRow.kind).toBe('table-row')
    if (firstRow.kind !== 'table-row') {
      throw new Error('Expected table row')
    }
    expect(firstRow.props).toMatchObject({
      trHeight: { val: 360, hRule: 'exact' },
      cantSplit: true,
      tblHeader: true,
    })
    expect(firstRow.cells).toHaveLength(2)

    const firstCell = firstRow.cells[0]
    expect(firstCell.kind).toBe('table-cell')
    if (firstCell.kind !== 'table-cell') {
      throw new Error('Expected table cell')
    }
    expect(firstCell.props).toMatchObject({
      tcW: { type: 'dxa', value: 2400 },
      gridSpan: 1,
      vAlign: 'center',
    })

    const secondRow = table.rows[1]
    expect(secondRow.kind).toBe('table-row')
    if (secondRow.kind !== 'table-row') {
      throw new Error('Expected table row')
    }

    const nestedHostCell = secondRow.cells[1]
    expect(nestedHostCell.kind).toBe('table-cell')
    if (nestedHostCell.kind !== 'table-cell') {
      throw new Error('Expected nested host cell')
    }

    expect(nestedHostCell.props).toMatchObject({
      vMerge: 'restart',
      noWrap: true,
      hideMark: true,
    })

    const nestedTable = expectTable(nestedHostCell.blocks[0])
    expect(nestedTable.rows).toHaveLength(2)
    expect(nestedTable.rows[0].kind).toBe('table-row')
  })

  it('parses w:tblGrid into Table.tblGrid (P1.5 / DXP-19)', () => {
    const document = parseBody(`
      <w:tbl>
        <w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>
        <w:tblGrid>
          <w:gridCol w:w="2400"/>
          <w:gridCol w:w="3600"/>
        </w:tblGrid>
        <w:tr>
          <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
    `)

    const table = expectTable(document.sections[0].blocks[0])
    expect(table.tblGrid).toEqual([2400, 3600])
  })

  it('omits Table.tblGrid when no w:tblGrid element is present', () => {
    const document = parseBody(`
      <w:tbl>
        <w:tr>
          <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
    `)

    const table = expectTable(document.sections[0].blocks[0])
    expect(table.tblGrid).toBeUndefined()
  })

  it('parses hyperlink attributes and children', () => {
    const document = parseBody(`
      <w:p>
        <w:hyperlink r:id="rId7" w:anchor="target" w:tooltip="Go" w:tgtFrame="_blank" w:history="1">
          <w:r><w:t>Example</w:t></w:r>
        </w:hyperlink>
      </w:p>
    `)

    const paragraph = expectParagraph(document.sections[0].blocks[0])
    expect(paragraph.children[0]).toMatchObject({
      kind: 'hyperlink',
      relationshipId: 'rId7',
      anchor: 'target',
      tooltip: 'Go',
      targetFrame: '_blank',
      history: true,
    })
  })

  it('parses bookmark boundaries', () => {
    const document = parseBody(`
      <w:p>
        <w:bookmarkStart w:id="9" w:name="anchor" w:colFirst="0" w:colLast="2"/>
        <w:r><w:t>Marked</w:t></w:r>
        <w:bookmarkEnd w:id="9"/>
      </w:p>
    `)

    const paragraph = expectParagraph(document.sections[0].blocks[0])

    expect(paragraph.children[0]).toMatchObject({
      kind: 'bookmark',
      id: '9',
      boundary: 'start',
      name: 'anchor',
      colFirst: 0,
      colLast: 2,
    })
    expect(paragraph.children[2]).toMatchObject({
      kind: 'bookmark',
      id: '9',
      boundary: 'end',
    })
  })

  it('parses numbering props from pPr', () => {
    const document = parseBody(`
      <w:p>
        <w:pPr>
          <w:numPr>
            <w:ilvl w:val="2"/>
            <w:numId w:val="17"/>
          </w:numPr>
        </w:pPr>
        <w:r><w:t>Item</w:t></w:r>
      </w:p>
    `)

    const paragraph = expectParagraph(document.sections[0].blocks[0])
    expect(paragraph.props).toMatchObject({
      numPr: {
        ilvl: 2,
        numId: '17',
      },
    })
  })

  it('parses section props and partitions paragraph/body section breaks', () => {
    const document = parseBody(`
      <w:p>
        <w:pPr>
          <w:sectPr>
            <w:type w:val="continuous"/>
            <w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>
            <w:pgMar w:top="1440" w:right="720" w:bottom="1440" w:left="720" w:header="360" w:footer="360" w:gutter="0"/>
            <w:cols w:num="2" w:space="720" w:sep="1" w:equalWidth="1">
              <w:col w:w="4680" w:space="720"/>
              <w:col w:w="4680" w:space="720"/>
            </w:cols>
            <w:pgNumType w:start="3" w:fmt="decimal"/>
            <w:titlePg/>
            <w:headerReference w:type="default" r:id="rIdHeader"/>
            <w:footerReference w:type="even" r:id="rIdFooter"/>
          </w:sectPr>
        </w:pPr>
        <w:r><w:t>Section A</w:t></w:r>
      </w:p>
      <w:p><w:r><w:t>Section B</w:t></w:r></w:p>
      <w:sectPr>
        <w:type w:val="nextPage"/>
      </w:sectPr>
    `)

    expect(document.sections).toHaveLength(2)
    expect(document.sections[0].props).toMatchObject({
      type: 'continuous',
      pgSz: { w: 12240, h: 15840, orient: 'portrait' },
      pgMar: {
        top: 1440,
        right: 720,
        bottom: 1440,
        left: 720,
        header: 360,
        footer: 360,
        gutter: 0,
      },
      cols: {
        num: 2,
        space: 720,
        sep: true,
        equalWidth: true,
        col: [
          { w: 4680, space: 720 },
          { w: 4680, space: 720 },
        ],
      },
      pgNumType: { start: 3, fmt: 'decimal' },
      titlePg: true,
      headerReference: [{ id: 'rIdHeader', type: 'default' }],
      footerReference: [{ id: 'rIdFooter', type: 'even' }],
    })
    expect(document.sections[1].props).toMatchObject({ type: 'nextPage' })
    expect(document.sections[1].blocks).toHaveLength(1)
  })

  it('parses shading on paragraph and run props', () => {
    const document = parseBody(`
      <w:p>
        <w:pPr>
          <w:shd w:val="clear" w:fill="EEE8AA" w:color="auto"/>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:shd w:val="solid" w:fill="111111" w:color="FFFFFF"/>
          </w:rPr>
          <w:t>Shade</w:t>
        </w:r>
      </w:p>
    `)

    const paragraph = expectParagraph(document.sections[0].blocks[0])
    const run = expectRun(paragraph.children[0])

    expect(paragraph.props?.shd).toMatchObject({
      pattern: 'clear',
      fill: 'EEE8AA',
      color: 'auto',
    })
    expect(run.props?.shd).toMatchObject({
      pattern: 'solid',
      fill: '111111',
      color: 'FFFFFF',
    })
  })

  it('parses paragraph and table cell borders', () => {
    const document = parseBody(`
      <w:p>
        <w:pPr>
          <w:pBdr>
            <w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>
            <w:bottom w:val="double" w:sz="12" w:space="4" w:color="FF00FF"/>
          </w:pBdr>
        </w:pPr>
      </w:p>
      <w:tbl>
        <w:tr>
          <w:tc>
            <w:tcPr>
              <w:tcBorders>
                <w:left w:val="single" w:sz="8" w:color="00FF00"/>
                <w:right w:val="single" w:sz="8" w:color="00FF00"/>
              </w:tcBorders>
            </w:tcPr>
            <w:p/>
          </w:tc>
        </w:tr>
      </w:tbl>
    `)

    const paragraph = expectParagraph(document.sections[0].blocks[0])
    expect(paragraph.props?.pBdr).toMatchObject({
      top: { style: 'single', size: 8, color: '000000', space: 0 },
      bottom: { style: 'double', size: 12, color: 'FF00FF', space: 4 },
    })

    const table = expectTable(document.sections[0].blocks[1])
    const row = table.rows[0]
    expect(row.kind).toBe('table-row')
    if (row.kind !== 'table-row') {
      throw new Error('Expected table row')
    }

    const cell = row.cells[0]
    expect(cell.kind).toBe('table-cell')
    if (cell.kind !== 'table-cell') {
      throw new Error('Expected table cell')
    }

    expect(cell.props?.tcBorders).toMatchObject({
      left: { style: 'single', size: 8, color: '00FF00' },
      right: { style: 'single', size: 8, color: '00FF00' },
    })
  })

  it('keeps style references alongside direct run and paragraph formatting', () => {
    const document = parseBody(`
      <w:p>
        <w:pPr>
          <w:pStyle w:val="Heading1"/>
          <w:spacing w:before="240" w:after="120" w:line="480" w:lineRule="auto"/>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:rStyle w:val="Strong"/>
            <w:b/>
            <w:color w:val="336699"/>
          </w:rPr>
          <w:t>Heading</w:t>
        </w:r>
      </w:p>
    `)

    const paragraph = expectParagraph(document.sections[0].blocks[0])
    const run = expectRun(paragraph.children[0])

    expect(paragraph.props).toMatchObject({
      pStyle: 'Heading1',
      spacing: {
        before: 240,
        after: 120,
        line: 480,
        lineRule: 'auto',
      },
    })
    expect(run.props).toMatchObject({
      rStyle: 'Strong',
      bold: true,
      color: '336699',
    })
  })

  it('preserves unsupported block XML as UnknownNode', () => {
    const document = parseBody(
      '<w:customBlock w:foo="bar"><w:customChild w:val="1"/></w:customBlock>',
    )

    expect(document.sections[0].blocks[0]).toEqual({
      kind: 'unknown',
      xml: '<w:customBlock w:foo="bar"><w:customChild w:val="1"/></w:customBlock>',
    })
  })

  // D19 / DXS-16
  describe('unknown-node raw substring preservation (D19 / DXS-16)', () => {
    it('captures the exact source text of an unknown element that declares its own namespace locally, byte-for-byte', () => {
      const document = parseBody(
        '<ext:widget xmlns:ext="urn:example:ext" ext:mode="live"><ext:payload>data &amp; more</ext:payload></ext:widget>',
      )

      expect(document.sections[0].blocks[0]).toEqual({
        kind: 'unknown',
        xml: '<ext:widget xmlns:ext="urn:example:ext" ext:mode="live"><ext:payload>data &amp; more</ext:payload></ext:widget>',
      })
    })

    it('preserves single-quoted attribute values and other exact source formatting a tree-rebuild would normalize away', () => {
      // fast-xml-parser's XMLBuilder always re-emits double-quoted attributes
      // regardless of how the source quoted them — a tree-rebuild (the old
      // `xmlBuilder.build([element])` approach) silently normalizes this,
      // which is not byte-for-byte preservation even though it's harmless
      // XML. Slicing the raw source substring keeps the original quoting.
      const document = parseBody("<w:customBlock w:foo='bar'/>")

      expect(document.sections[0].blocks[0]).toEqual({
        kind: 'unknown',
        xml: "<w:customBlock w:foo='bar'/>",
      })
    })

    it(
      're-declares, on the unknown element itself, a namespace prefix it uses that only a '
        + 'non-root ancestor (not the unknown element itself) declared — otherwise lost entirely '
        + 'since neither the raw substring nor a tree-rebuild of the unknown element alone '
        + 'includes an ancestor\'s own attributes',
      () => {
        const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body xmlns:ext="urn:example:ext">
    <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
    <ext:widget ext:mode="live"><ext:payload>data</ext:payload></ext:widget>
  </w:body>
</w:document>`

        const document = parseDocument(xml)
        const unknownBlock = document.sections[0].blocks[document.sections[0].blocks.length - 1]

        expect(unknownBlock).toEqual({
          kind: 'unknown',
          xml: '<ext:widget ext:mode="live" xmlns:ext="urn:example:ext"><ext:payload>data</ext:payload></ext:widget>',
        })

        // The re-declaration must actually make the captured fragment valid,
        // standalone XML — not just look plausible.
        expect(() => new XMLParser({ ignoreAttributes: false }).parse(
          `<root xmlns:ext="should-be-overridden">${(unknownBlock as { xml: string }).xml}</root>`,
        )).not.toThrow()
      },
    )

    it('does not re-declare a namespace the unknown element already declares itself, even if an ancestor also declares it', () => {
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body xmlns:ext="urn:example:ext">
    <ext:widget xmlns:ext="urn:example:ext-local"><ext:payload>data</ext:payload></ext:widget>
  </w:body>
</w:document>`

      const document = parseDocument(xml)

      expect(document.sections[0].blocks[0]).toEqual({
        kind: 'unknown',
        xml: '<ext:widget xmlns:ext="urn:example:ext-local"><ext:payload>data</ext:payload></ext:widget>',
      })
    })

    it('does not re-declare a namespace the document root itself already declares (DXS-15 already re-emits those)', () => {
      // `w:` is declared on the root `w:document`, not on some intermediate
      // ancestor — `rootNamespaces`/`mcIgnorable` already guarantee it
      // survives at the root, so injecting it again here would be redundant
      // clutter and would break byte-identical preservation for the
      // overwhelmingly common case (this is exactly the existing
      // "preserves unsupported block XML as UnknownNode" test's shape).
      const document = parseBody('<w:customBlock><w:customChild w:val="1"/></w:customBlock>')

      expect(document.sections[0].blocks[0]).toEqual({
        kind: 'unknown',
        xml: '<w:customBlock><w:customChild w:val="1"/></w:customBlock>',
      })
    })
  })

  it('preserves xml:space="preserve" on text nodes', () => {
    const document = parseBody(
      '<w:p><w:r><w:t xml:space="preserve">  spaced  </w:t></w:r></w:p>',
    )

    const paragraph = expectParagraph(document.sections[0].blocks[0])
    const run = expectRun(paragraph.children[0])

    expect(run.children[0]).toEqual({
      kind: 'text',
      value: '  spaced  ',
      preserveSpace: true,
    })
  })

  it('parses an empty paragraph', () => {
    const document = parseBody('<w:p/>')

    const paragraph = expectParagraph(document.sections[0].blocks[0])
    expect(paragraph.children).toEqual([])
  })

  it('partitions multiple sections across paragraph and document-level sectPr', () => {
    const document = parseBody(`
      <w:p><w:r><w:t>A</w:t></w:r></w:p>
      <w:p>
        <w:pPr><w:sectPr><w:type w:val="continuous"/></w:sectPr></w:pPr>
        <w:r><w:t>B</w:t></w:r>
      </w:p>
      <w:p>
        <w:pPr><w:sectPr><w:type w:val="evenPage"/></w:sectPr></w:pPr>
        <w:r><w:t>C</w:t></w:r>
      </w:p>
      <w:p><w:r><w:t>D</w:t></w:r></w:p>
      <w:sectPr><w:type w:val="oddPage"/></w:sectPr>
    `)

    expect(document.sections).toHaveLength(3)
    expect(document.sections[0].props.type).toBe('continuous')
    expect(document.sections[0].blocks).toHaveLength(2)
    expect(document.sections[1].props.type).toBe('evenPage')
    expect(document.sections[1].blocks).toHaveLength(1)
    expect(document.sections[2].props.type).toBe('oddPage')
    expect(document.sections[2].blocks).toHaveLength(1)
  })

  it('parses simple drawing metadata', () => {
    const document = parseBody(`
      <w:p>
        <w:r>
          <w:drawing>
            <wp:inline>
              <wp:extent cx="914400" cy="457200"/>
              <wp:docPr id="1" name="Picture 1" title="Chart" descr="Quarterly chart"/>
              <a:graphic>
                <a:graphicData>
                  <pic:pic>
                    <pic:blipFill>
                      <a:blip r:embed="rIdImage1"/>
                    </pic:blipFill>
                  </pic:pic>
                </a:graphicData>
              </a:graphic>
            </wp:inline>
          </w:drawing>
        </w:r>
      </w:p>
    `)

    const paragraph = expectParagraph(document.sections[0].blocks[0])
    const run = expectRun(paragraph.children[0])

    expect(run.children[0]).toMatchObject({
      kind: 'drawing',
      layout: 'inline',
      relationshipId: 'rIdImage1',
      title: 'Chart',
      description: 'Quarterly chart',
      name: 'Picture 1',
      extent: {
        cx: 914400,
        cy: 457200,
      },
    })
  })

  it('captures wp:anchor position/wrap children so a save stays schema-valid (P1.7 / DXS-03)', () => {
    const document = parseBody(`
      <w:p>
        <w:r>
          <w:drawing>
            <wp:anchor>
              <wp:simplePos x="0" y="0"/>
              <wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH>
              <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
              <wp:extent cx="914400" cy="457200"/>
              <wp:effectExtent l="0" t="0" r="0" b="0"/>
              <wp:wrapSquare wrapText="bothSides"/>
              <wp:docPr id="1" name="Picture 1"/>
              <wp:cNvGraphicFramePr/>
              <a:graphic>
                <a:graphicData>
                  <pic:pic>
                    <pic:blipFill>
                      <a:blip r:embed="rIdImage2"/>
                    </pic:blipFill>
                  </pic:pic>
                </a:graphicData>
              </a:graphic>
            </wp:anchor>
          </w:drawing>
        </w:r>
      </w:p>
    `)

    const paragraph = expectParagraph(document.sections[0].blocks[0])
    const run = expectRun(paragraph.children[0])
    const drawing = run.children[0]

    expect(drawing).toMatchObject({ kind: 'drawing', layout: 'anchor', relationshipId: 'rIdImage2' })
    if (drawing.kind !== 'drawing') {
      throw new Error('Expected drawing child')
    }

    const slots = (drawing.anchorChildren ?? []).map((entry) =>
      entry.kind === 'anchor-slot' ? entry.slot : 'unknown',
    )
    // Original order preserved: simplePos, positionH, positionV, extent,
    // effectExtent, wrapSquare, docPr, cNvGraphicFramePr, graphic.
    expect(slots).toEqual([
      'unknown',
      'unknown',
      'unknown',
      'extent',
      'unknown',
      'unknown',
      'docPr',
      'unknown',
      'graphic',
    ])

    const unknownXml = (drawing.anchorChildren ?? [])
      .filter((entry): entry is Extract<typeof entry, { kind: 'unknown' }> => entry.kind === 'unknown')
      .map((entry) => entry.xml)
      .join('')
    expect(unknownXml).toContain('wp:simplePos')
    expect(unknownXml).toContain('wp:positionH')
    expect(unknownXml).toContain('wp:positionV')
    expect(unknownXml).toContain('wp:wrapSquare')
    expect(unknownXml).toContain('wp:cNvGraphicFramePr')
  })

  it('preserves a non-picture graphicFrame (chart/SmartArt) as an unknown node instead of a lossy Drawing (P1.7 / DXS-04)', () => {
    const document = parseBody(`
      <w:p>
        <w:r>
          <w:drawing>
            <wp:inline>
              <wp:extent cx="914400" cy="457200"/>
              <wp:docPr id="2" name="Chart 1"/>
              <a:graphic>
                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <c:chart r:id="rIdChart1"/>
                </a:graphicData>
              </a:graphic>
            </wp:inline>
          </w:drawing>
        </w:r>
      </w:p>
    `)

    const paragraph = expectParagraph(document.sections[0].blocks[0])
    const run = expectRun(paragraph.children[0])
    const child = run.children[0]

    expect(child.kind).toBe('unknown')
    if (child.kind === 'unknown') {
      expect(child.xml).toContain('rIdChart1')
      expect(child.xml).toContain('c:chart')
    }
  })

  it('parses comment range boundaries and comment references', () => {
    const document = parseBody(`
      <w:p>
        <w:commentRangeStart w:id="5"/>
        <w:r><w:t>Commented</w:t></w:r>
        <w:commentRangeEnd w:id="5"/>
        <w:r><w:commentReference w:id="5"/></w:r>
      </w:p>
    `)

    const paragraph = expectParagraph(document.sections[0].blocks[0])
    const commentRun = expectRun(paragraph.children[3])

    expect(paragraph.children[0]).toMatchObject({
      kind: 'comment-range',
      id: '5',
      boundary: 'start',
    })
    expect(paragraph.children[2]).toMatchObject({
      kind: 'comment-range',
      id: '5',
      boundary: 'end',
    })
    expect(commentRun.children[0]).toMatchObject({
      kind: 'comment-reference',
      id: '5',
    })
  })

  it('parses footnote and endnote references', () => {
    const document = parseBody(`
      <w:p>
        <w:r><w:footnoteReference w:id="2" w:customMarkFollows="1"/></w:r>
        <w:r><w:endnoteReference w:id="3" w:customMarkFollows="0"/></w:r>
      </w:p>
    `)

    const paragraph = expectParagraph(document.sections[0].blocks[0])
    const footnoteRun = expectRun(paragraph.children[0])
    const endnoteRun = expectRun(paragraph.children[1])

    expect(footnoteRun.children[0]).toMatchObject({
      kind: 'footnote-reference',
      id: '2',
      customMarkFollows: true,
    })
    expect(endnoteRun.children[0]).toMatchObject({
      kind: 'endnote-reference',
      id: '3',
      customMarkFollows: false,
    })
  })

  describe('size guard (D21 / DXP-20)', () => {
    it('refuses a document.xml part over the size limit before parsing it', () => {
      const oversized = documentXml(`<w:p><w:r><w:t>${'x'.repeat(MAX_XML_PART_LENGTH)}</w:t></w:r></w:p>`)

      expect(() => parseDocument(oversized)).toThrow(DocxParseError)
    })
  })

  describe('paraId/textId/rsid passthrough (D19 / DXS-10)', () => {
    it('captures w14:paraId/w14:textId and w:rsid* attributes on a paragraph', () => {
      const document = parseBody(`
        <w:p w14:paraId="12AB34CD" w14:textId="56EF78AB" w:rsidR="00112233"
             w:rsidRDefault="00112233" w:rsidP="00445566" w:rsidRPr="00778899">
          <w:r><w:t>Hello</w:t></w:r>
        </w:p>
      `)

      const paragraph = expectParagraph(document.sections[0].blocks[0])
      expect(paragraph.paraId).toBe('12AB34CD')
      expect(paragraph.textId).toBe('56EF78AB')
      expect(paragraph.rsidR).toBe('00112233')
      expect(paragraph.rsidRDefault).toBe('00112233')
      expect(paragraph.rsidP).toBe('00445566')
      expect(paragraph.rsidRPr).toBe('00778899')
    })

    it('captures w:rsid* attributes on a run', () => {
      const document = parseBody(`
        <w:p>
          <w:r w:rsidR="00AA0011" w:rsidRPr="00AA0022" w:rsidDel="00AA0033">
            <w:t>Hello</w:t>
          </w:r>
        </w:p>
      `)

      const paragraph = expectParagraph(document.sections[0].blocks[0])
      const run = expectRun(paragraph.children[0])
      expect(run.rsidR).toBe('00AA0011')
      expect(run.rsidRPr).toBe('00AA0022')
      expect(run.rsidDel).toBe('00AA0033')
    })

    it('leaves paraId/textId/rsid undefined when the source paragraph never had them', () => {
      const document = parseBody(`<w:p><w:r><w:t>Hello</w:t></w:r></w:p>`)

      const paragraph = expectParagraph(document.sections[0].blocks[0])
      expect(paragraph.paraId).toBeUndefined()
      expect(paragraph.textId).toBeUndefined()
      expect(paragraph.rsidR).toBeUndefined()
    })
  })

  describe('w:sdt / mc:AlternateContent / w:sym (D8 / DXP-08)', () => {
    it('unwraps a block-level w:sdt to its sdtContent paragraph', () => {
      const document = parseBody(`
        <w:sdt>
          <w:sdtPr><w:alias w:val="Title Control"/></w:sdtPr>
          <w:sdtContent>
            <w:p><w:r><w:t>Content control text</w:t></w:r></w:p>
          </w:sdtContent>
        </w:sdt>
      `)

      expect(document.sections[0].blocks).toHaveLength(1)
      const paragraph = expectParagraph(document.sections[0].blocks[0])
      const run = expectRun(paragraph.children[0])
      expect(run.children[0]).toEqual({ kind: 'text', value: 'Content control text' })
    })

    it('unwraps an inline (paragraph-level) w:sdt to its sdtContent run', () => {
      const document = parseBody(`
        <w:p>
          <w:r><w:t>Before </w:t></w:r>
          <w:sdt>
            <w:sdtContent>
              <w:r><w:t>inline control</w:t></w:r>
            </w:sdtContent>
          </w:sdt>
          <w:r><w:t> after</w:t></w:r>
        </w:p>
      `)

      const paragraph = expectParagraph(document.sections[0].blocks[0])
      expect(paragraph.children).toHaveLength(3)
      const runs = paragraph.children.map((child) => expectRun(child))
      expect(runs.map((run) => run.children[0])).toEqual([
        { kind: 'text', value: 'Before ' },
        { kind: 'text', value: 'inline control' },
        { kind: 'text', value: ' after' },
      ])
    })

    it('unwraps a w:sdt with an empty sdtContent to nothing, without throwing', () => {
      const document = parseBody(`
        <w:sdt>
          <w:sdtContent/>
        </w:sdt>
      `)

      expect(document.sections[0].blocks).toHaveLength(0)
    })

    it('prefers mc:Choice over mc:Fallback in an mc:AlternateContent wrapper', () => {
      const document = parseBody(`
        <w:p>
          <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
            <mc:Choice Requires="wps"><w:r><w:t>modern shape</w:t></w:r></mc:Choice>
            <mc:Fallback><w:r><w:t>legacy VML</w:t></w:r></mc:Fallback>
          </mc:AlternateContent>
        </w:p>
      `)

      const paragraph = expectParagraph(document.sections[0].blocks[0])
      const run = expectRun(paragraph.children[0])
      expect(run.children[0]).toEqual({ kind: 'text', value: 'modern shape' })
    })

    it('falls back to mc:Fallback when an mc:AlternateContent wrapper has no mc:Choice', () => {
      const document = parseBody(`
        <w:p>
          <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
            <mc:Fallback><w:r><w:t>legacy VML only</w:t></w:r></mc:Fallback>
          </mc:AlternateContent>
        </w:p>
      `)

      const paragraph = expectParagraph(document.sections[0].blocks[0])
      const run = expectRun(paragraph.children[0])
      expect(run.children[0]).toEqual({ kind: 'text', value: 'legacy VML only' })
    })

    it('shows nothing (rather than throwing) for an mc:AlternateContent wrapper with neither branch', () => {
      const document = parseBody(`
        <w:p>
          <w:r><w:t>Before</w:t></w:r>
          <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
          </mc:AlternateContent>
          <w:r><w:t>After</w:t></w:r>
        </w:p>
      `)

      const paragraph = expectParagraph(document.sections[0].blocks[0])
      expect(paragraph.children).toHaveLength(2)
    })

    it('decodes a w:sym into a synthetic text node', () => {
      const document = parseBody(`
        <w:p><w:r><w:sym w:font="Wingdings" w:char="F0E0"/></w:r></w:p>
      `)

      const paragraph = expectParagraph(document.sections[0].blocks[0])
      const run = expectRun(paragraph.children[0])
      expect(run.children[0]).toEqual({ kind: 'text', value: String.fromCharCode(0xf0e0) })
    })

    it('produces an empty text node for a w:sym missing w:char instead of throwing', () => {
      const document = parseBody(`
        <w:p><w:r><w:sym w:font="Wingdings"/></w:r></w:p>
      `)

      const paragraph = expectParagraph(document.sections[0].blocks[0])
      const run = expectRun(paragraph.children[0])
      expect(run.children[0]).toEqual({ kind: 'text', value: '' })
    })
  })

  describe('malformed XML handling (D28 / DXP-16)', () => {
    it('wraps a fast-xml-parser failure in DocxParseError instead of letting it propagate raw', () => {
      expect(() => parseDocument('<<< not xml <<<')).toThrow(DocxParseError)
    })

    it('includes the underlying parser message in the DocxParseError', () => {
      try {
        parseDocument('<<< not xml <<<')
        expect.unreachable('parseDocument should have thrown')
      } catch (error) {
        expect(error).toBeInstanceOf(DocxParseError)
        expect((error as Error).message).toContain('Failed to parse document XML')
      }
    })
  })
})
