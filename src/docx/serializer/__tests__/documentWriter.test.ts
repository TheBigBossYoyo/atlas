import { describe, expect, it } from 'vitest'

import {
  hexColor,
  halfPoint,
  twip,
  type Document,
  type Paragraph,
  type Run,
  type Section,
} from '../../model'
import { parseDocument } from '../../parser'
import { writeDocumentXml } from '../documentWriter'

const BASE_NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
].join(' ')

function documentXml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${BASE_NAMESPACES}><w:body>${body}</w:body></w:document>`
}

function createDocument(sections: ReadonlyArray<Section>): Document {
  return {
    kind: 'document',
    sections,
    styles: new Map(),
    numbering: new Map(),
    comments: new Map(),
    footnotes: new Map(),
    endnotes: new Map(),
    headers: new Map(),
    footers: new Map(),
  }
}

function createParagraph(text: string, props?: Paragraph['props']): Paragraph {
  const run: Run = {
    kind: 'run',
    children: [{ kind: 'text', value: text }],
  }

  return {
    kind: 'paragraph',
    ...(props !== undefined ? { props } : {}),
    children: [run],
  }
}

function normalizeAst(value: unknown): unknown {
  if (value instanceof Map) {
    return Array.from(value.entries(), ([key, entry]) => [key, normalizeAst(entry)])
  }

  if (Array.isArray(value)) {
    return value.map((entry) => normalizeAst(entry))
  }

  if (value !== null && typeof value === 'object') {
    const record = value as Record<string, unknown>
    const normalized: Record<string, unknown> = {}

    for (const [key, entry] of Object.entries(record)) {
      if (record.kind === 'text' && key === 'preserveSpace') {
        continue
      }

      normalized[key] = normalizeAst(entry)
    }

    return normalized
  }

  return value
}

function expectRoundTrip(xml: string): void {
  const parsed = parseDocument(xml)
  const written = writeDocumentXml(parsed)
  const reparsed = parseDocument(written)
  expect(normalizeAst(reparsed)).toEqual(normalizeAst(parsed))
}

describe('writeDocumentXml', () => {
  it('writes an empty document with body and trailing sectPr', () => {
    const document = createDocument([
      {
        kind: 'section',
        props: {},
        blocks: [],
      },
    ])

    const xml = writeDocumentXml(document)

    expect(xml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')).toBe(true)
    expect(xml).toContain('<w:document')
    expect(xml).toContain('<w:body>')
    expect(xml).toContain('<w:sectPr/>')
  })

  it('round-trips a single paragraph with a single text run', () => {
    const document = createDocument([
      {
        kind: 'section',
        props: {},
        blocks: [createParagraph('Hello Atlas')],
      },
    ])

    const xml = writeDocumentXml(document)
    const reparsed = parseDocument(xml)

    expect(normalizeAst(reparsed)).toEqual(normalizeAst(document))
  })

  it('emits bold run formatting', () => {
    const document = createDocument([
      {
        kind: 'section',
        props: {},
        blocks: [{
          kind: 'paragraph',
          children: [{
            kind: 'run',
            props: { bold: true },
            children: [{ kind: 'text', value: 'Bold' }],
          }],
        }],
      },
    ])

    const xml = writeDocumentXml(document)
    expect(xml).toContain('<w:b/>')
  })

  it('emits italic run formatting', () => {
    const document = createDocument([
      {
        kind: 'section',
        props: {},
        blocks: [{
          kind: 'paragraph',
          children: [{
            kind: 'run',
            props: { italic: true },
            children: [{ kind: 'text', value: 'Italic' }],
          }],
        }],
      },
    ])

    const xml = writeDocumentXml(document)
    expect(xml).toContain('<w:i/>')
  })

  it('emits underline run formatting', () => {
    const document = createDocument([
      {
        kind: 'section',
        props: {},
        blocks: [{
          kind: 'paragraph',
          children: [{
            kind: 'run',
            props: {
              underline: {
                style: 'single',
                color: hexColor('FF0000'),
              },
            },
            children: [{ kind: 'text', value: 'Underline' }],
          }],
        }],
      },
    ])

    const xml = writeDocumentXml(document)
    expect(xml).toContain('<w:u w:val="single" w:color="FF0000"/>')
  })

  it('emits color, size, and font family run formatting', () => {
    const document = createDocument([
      {
        kind: 'section',
        props: {},
        blocks: [{
          kind: 'paragraph',
          children: [{
            kind: 'run',
            props: {
              color: hexColor('336699'),
              sz: halfPoint(28),
              rFonts: {
                ascii: 'Calibri',
                hAnsi: 'Calibri',
              },
            },
            children: [{ kind: 'text', value: 'Styled' }],
          }],
        }],
      },
    ])

    const xml = writeDocumentXml(document)
    expect(xml).toContain('<w:color w:val="336699"/>')
    expect(xml).toContain('<w:sz w:val="28"/>')
    expect(xml).toContain('<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>')
  })

  it('emits paragraph alignment, spacing, indent, and style', () => {
    const document = createDocument([
      {
        kind: 'section',
        props: {},
        blocks: [createParagraph('Paragraph props', {
          pStyle: 'Heading1',
          jc: 'center',
          spacing: {
            before: twip(240),
            after: twip(120),
            line: twip(480),
            lineRule: 'auto',
          },
          ind: {
            left: twip(720),
            firstLine: twip(360),
          },
        })],
      },
    ])

    const xml = writeDocumentXml(document)
    expect(xml).toContain('<w:pStyle w:val="Heading1"/>')
    expect(xml).toContain('<w:jc w:val="center"/>')
    expect(xml).toContain('<w:spacing w:before="240" w:after="120" w:line="480" w:lineRule="auto"/>')
    expect(xml).toContain('<w:ind w:left="720" w:firstLine="360"/>')
  })

  it('emits tab and break nodes', () => {
    const document = createDocument([
      {
        kind: 'section',
        props: {},
        blocks: [{
          kind: 'paragraph',
          children: [{
            kind: 'run',
            children: [
              { kind: 'tab' },
              { kind: 'break', breakType: 'line' },
              { kind: 'break', breakType: 'page' },
            ],
          }],
        }],
      },
    ])

    const xml = writeDocumentXml(document)
    expect(xml).toContain('<w:tab/>')
    expect(xml).toContain('<w:br/>')
    expect(xml).toContain('<w:br w:type="page"/>')
  })

  it('round-trips a 2x2 table', () => {
    const xml = documentXml(`
      <w:tbl>
        <w:tr>
          <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
      <w:sectPr/>
    `)

    expectRoundTrip(xml)
  })

  it('round-trips a table with w:tblGrid (P1.5 / DXS-02)', () => {
    const xml = documentXml(`
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
      <w:sectPr/>
    `)

    expectRoundTrip(xml)

    const written = writeDocumentXml(parseDocument(xml))
    expect(written).toContain('<w:tblGrid>')
    expect(written).toContain('<w:gridCol w:w="2400"/>')
    expect(written).toContain('<w:gridCol w:w="3600"/>')
  })

  it('emits w:tblGrid immediately after w:tblPr, before any w:tr', () => {
    const document = createDocument([
      {
        kind: 'section',
        props: {},
        blocks: [
          {
            kind: 'table',
            props: { tblLayout: 'fixed' },
            tblGrid: [twip(1000), twip(2000)],
            rows: [
              {
                kind: 'table-row',
                cells: [
                  { kind: 'table-cell', blocks: [createParagraph('A1')] },
                  { kind: 'table-cell', blocks: [createParagraph('A2')] },
                ],
              },
            ],
          },
        ],
      },
    ])

    const xml = writeDocumentXml(document)
    const tblPrIndex = xml.indexOf('<w:tblPr')
    const tblGridIndex = xml.indexOf('<w:tblGrid')
    const trIndex = xml.indexOf('<w:tr')

    expect(tblPrIndex).toBeGreaterThanOrEqual(0)
    expect(tblGridIndex).toBeGreaterThan(tblPrIndex)
    expect(trIndex).toBeGreaterThan(tblGridIndex)
  })

  it('round-trips an anchored image with position/wrap children as schema-valid wp:anchor output (P1.7 / DXS-03)', () => {
    const xml = documentXml(`
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

    expectRoundTrip(xml)

    const written = writeDocumentXml(parseDocument(xml))
    // The schema-required children the pre-fix serializer silently dropped.
    expect(written).toContain('wp:simplePos')
    expect(written).toContain('wp:positionH')
    expect(written).toContain('wp:positionV')
    expect(written).toContain('wp:wrapSquare')
    expect(written).toContain('wp:cNvGraphicFramePr')
    expect(written).toContain('rIdImage2')
  })

  it('preserves a non-picture graphicFrame (chart) verbatim instead of emitting an empty wrapper (P1.7 / DXS-04)', () => {
    const xml = documentXml(`
      <w:p>
        <w:r>
          <w:drawing>
            <wp:inline>
              <wp:extent cx="914400" cy="457200"/>
              <wp:docPr id="2" name="Chart 1"/>
              <a:graphic>
                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart1"/>
                </a:graphicData>
              </a:graphic>
            </wp:inline>
          </w:drawing>
        </w:r>
      </w:p>
    `)

    expectRoundTrip(xml)

    const written = writeDocumentXml(parseDocument(xml))
    expect(written).toContain('c:chart')
    expect(written).toContain('rIdChart1')
    // Never emit a picture-shaped wp:inline with no a:blip for chart content.
    expect(written).not.toContain('pic:pic')
  })

  it('round-trips multiple sections', () => {
    const xml = documentXml(`
      <w:p>
        <w:pPr>
          <w:sectPr>
            <w:type w:val="continuous"/>
            <w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>
          </w:sectPr>
        </w:pPr>
        <w:r><w:t>Section A</w:t></w:r>
      </w:p>
      <w:p><w:r><w:t>Section B</w:t></w:r></w:p>
      <w:sectPr><w:type w:val="nextPage"/></w:sectPr>
    `)

    const parsed = parseDocument(xml)
    const written = writeDocumentXml(parsed)
    const reparsed = parseDocument(written)

    expect(normalizeAst(reparsed)).toEqual(normalizeAst(parsed))
    expect((written.match(/<w:sectPr/g) ?? []).length).toBe(2)
  })

  it('preserves UnknownNode XML verbatim', () => {
    const xml = documentXml('<w:customBlock w:foo="bar"><w:customChild w:val="1"/></w:customBlock>')
    const parsed = parseDocument(xml)
    const written = writeDocumentXml(parsed)

    expect(written).toContain('<w:customBlock w:foo="bar"><w:customChild w:val="1"/></w:customBlock>')
    expectRoundTrip(xml)
  })

  it('escapes special characters and preserves unicode text', () => {
    const text = '5 < 7 & 9 > 2 "quoted" \'single\' café Ω'
    const document = createDocument([
      {
        kind: 'section',
        props: {},
        blocks: [createParagraph(text)],
      },
    ])

    const xml = writeDocumentXml(document)
    const reparsed = parseDocument(xml)

    expect(xml).toContain('5 &lt; 7 &amp; 9 &gt; 2')
    expect((reparsed.sections[0]?.blocks[0] as Paragraph).children[0]).toMatchObject({
      kind: 'run',
      children: [{ kind: 'text', value: text, preserveSpace: true }],
    })
  })

  it('adds xml:space="preserve" to every text node', () => {
    const document = createDocument([
      {
        kind: 'section',
        props: {},
        blocks: [{
          kind: 'paragraph',
          children: [
            { kind: 'run', children: [{ kind: 'text', value: 'A' }] },
            { kind: 'run', children: [{ kind: 'text', value: ' B ' }] },
          ],
        }],
      },
    ])

    const xml = writeDocumentXml(document)
    const matches = xml.match(/<w:t xml:space="preserve">/g) ?? []
    expect(matches).toHaveLength(2)
  })

  it.each([
    ['simple paragraph', documentXml('<w:p><w:r><w:t>Hello</w:t></w:r></w:p><w:sectPr/>')],
    ['run formatting', documentXml('<w:p><w:r><w:rPr><w:b/><w:color w:val="336699"/></w:rPr><w:t>Styled</w:t></w:r></w:p><w:sectPr/>')],
    ['paragraph formatting', documentXml('<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:jc w:val="center"/></w:pPr><w:r><w:t>Centered</w:t></w:r></w:p><w:sectPr/>')],
    ['table', documentXml('<w:tbl><w:tr><w:tc><w:p><w:r><w:t>One</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/>')],
    ['formatting fixture with sections', documentXml('<w:p><w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr><w:r><w:rPr><w:i/><w:sz w:val="24"/></w:rPr><w:t>Fixture</w:t></w:r></w:p><w:p><w:pPr><w:sectPr><w:type w:val="evenPage"/></w:sectPr></w:pPr><w:r><w:t>Break</w:t></w:r></w:p>')],
    ['ins revision', documentXml('<w:p><w:ins w:id="1" w:author="Alice" w:date="2024-01-02T03:04:05Z"><w:r><w:t>added</w:t></w:r></w:ins></w:p><w:sectPr/>')],
    ['del revision', documentXml('<w:p><w:del w:id="2" w:author="Bob" w:date="2024-02-03T04:05:06Z"><w:r><w:delText>removed</w:delText></w:r></w:del></w:p><w:sectPr/>')],
    [
      'paraId/textId/rsid bookkeeping attributes (D19 / DXS-10)',
      documentXml(
        '<w:p w14:paraId="12AB34CD" w14:textId="56EF78AB" w:rsidR="00112233" w:rsidRDefault="00112233" w:rsidP="00445566" w:rsidRPr="00778899">'
          + '<w:r w:rsidR="00AA0011" w:rsidRPr="00AA0022" w:rsidDel="00AA0033"><w:t>Hello</w:t></w:r>'
          + '</w:p><w:sectPr/>',
      ),
    ],
  ])('round-trips parsed ASTs for %s', (_label, xml) => {
    expectRoundTrip(xml)
  })

  it('re-emits w14:paraId/w14:textId and w:rsid* attributes on save (D19 / DXS-10)', () => {
    const xml = documentXml(
      '<w:p w14:paraId="12AB34CD" w14:textId="56EF78AB" w:rsidR="00112233" w:rsidRDefault="00112233" w:rsidP="00445566" w:rsidRPr="00778899">'
        + '<w:r w:rsidR="00AA0011" w:rsidRPr="00AA0022" w:rsidDel="00AA0033"><w:t>Hello</w:t></w:r>'
        + '</w:p><w:sectPr/>',
    )
    const written = writeDocumentXml(parseDocument(xml))

    expect(written).toContain('w14:paraId="12AB34CD"')
    expect(written).toContain('w14:textId="56EF78AB"')
    expect(written).toContain('w:rsidR="00112233"')
    expect(written).toContain('w:rsidRDefault="00112233"')
    expect(written).toContain('w:rsidP="00445566"')
    expect(written).toContain('w:rsidRPr="00778899"')
    expect(written).toContain('w:rsidR="00AA0011"')
    expect(written).toContain('w:rsidRPr="00AA0022"')
    expect(written).toContain('w:rsidDel="00AA0033"')
  })

  it('does not invent w14:paraId/w:rsid attributes for a paragraph that never had them', () => {
    const document = createDocument([
      { kind: 'section', props: {}, blocks: [createParagraph('Fresh paragraph')] },
    ])

    const xml = writeDocumentXml(document)

    expect(xml).not.toContain('w14:paraId')
    expect(xml).not.toContain('w:rsid')
  })
})
