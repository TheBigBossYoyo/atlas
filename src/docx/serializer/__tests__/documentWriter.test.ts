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

      // `rootNamespaces`/`mcIgnorable` (D19 / DXS-15) are a source-only
      // capture that `writeDocumentXml` deliberately unions with Atlas's
      // own required baseline namespace set rather than reproducing
      // byte-for-byte — a fixture that declares only a handful of
      // namespaces legitimately re-parses with the full baseline added on
      // top after a round-trip. That's the point of the fix, not a
      // regression, so it's excluded from this structural-equality check.
      if (record.kind === 'document' && (key === 'rootNamespaces' || key === 'mcIgnorable')) {
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
            <wp:anchor behindDoc="1" allowOverlap="0">
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
    // DXP-09: position/wrap now come from typed model fields, not a raw
    // XML replay — the written output must still carry the actual parsed
    // values (not just the element names above).
    expect(written).toContain('behindDoc="1"')
    expect(written).toContain('allowOverlap="0"')
    expect(written).toContain('relativeFrom="column"')
    expect(written).toContain('relativeFrom="paragraph"')
    expect(written).toContain('914400')
    expect(written).toContain('wrapText="bothSides"')
  })

  it('round-trips picture crop/rotation/flip through the pic:spPr/pic:blipFill subtree (DXS-09)', () => {
    const xml = documentXml(`
      <w:p>
        <w:r>
          <w:drawing>
            <wp:inline>
              <wp:extent cx="914400" cy="457200"/>
              <wp:docPr id="1" name="Picture 1"/>
              <a:graphic>
                <a:graphicData>
                  <pic:pic>
                    <pic:blipFill>
                      <a:blip r:embed="rIdImage3"/>
                      <a:srcRect l="10000" t="5000" r="10000" b="5000"/>
                    </pic:blipFill>
                    <pic:spPr>
                      <a:xfrm rot="2700000" flipH="1" flipV="0"/>
                    </pic:spPr>
                  </pic:pic>
                </a:graphicData>
              </a:graphic>
            </wp:inline>
          </w:drawing>
        </w:r>
      </w:p>
    `)

    expectRoundTrip(xml)

    const written = writeDocumentXml(parseDocument(xml))
    expect(written).toContain('a:srcRect')
    expect(written).toContain('l="10000"')
    expect(written).toContain('a:xfrm')
    expect(written).toContain('rot="2700000"')
    expect(written).toContain('flipH="1"')
    expect(written).toContain('rIdImage3')
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

  it(
    'declares a custom namespace prefix used only by unknown-node passthrough content (D19 / DXS-15, DXS-16)',
    () => {
      // Previously the document root always emitted a fixed, hardcoded
      // namespace set (DXS-15) — a source document's own custom prefix
      // (used only inside content Atlas doesn't otherwise model) had no
      // declaration anywhere in the saved output, an XML well-formedness
      // violation. DXS-15's `rootNamespaces` capture closes the gap for
      // the common real-world case (Word always declares every namespace
      // it uses on the document root, not on some inner ancestor), which
      // is also DXS-16's primary practical concern.
      const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        + 'xmlns:ext123="http://example.com/custom-extension">'
        + '<w:body><w:customBlock ext123:feature="on"><ext123:payload>value</ext123:payload></w:customBlock></w:body>'
        + '</w:document>'

      const parsed = parseDocument(xml)
      expect(parsed.rootNamespaces?.get('ext123')).toBe('http://example.com/custom-extension')

      const written = writeDocumentXml(parsed)
      expect(written).toContain('xmlns:ext123="http://example.com/custom-extension"')
      expect(written).toContain(
        '<w:customBlock ext123:feature="on"><ext123:payload>value</ext123:payload></w:customBlock>',
      )
    },
  )

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
    [
      'w:sdt-wrapped paragraph, unedited — passed through verbatim (round-trip fidelity audit, DXS round 2 follow-up)',
      documentXml(
        '<w:sdt><w:sdtContent><w:p><w:r><w:t>Content control</w:t></w:r></w:p></w:sdtContent></w:sdt><w:sectPr/>',
      ),
    ],
    [
      'right-to-left paragraph + run (round-trip fidelity audit, DXS round 2)',
      documentXml(
        '<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rtl/></w:rPr><w:t>مرحبا</w:t></w:r></w:p><w:sectPr/>',
      ),
    ],
    [
      'page borders + section bidi (round-trip fidelity audit, DXS round 2)',
      documentXml(
        '<w:p><w:r><w:t>Bordered section</w:t></w:r></w:p>'
          + '<w:sectPr>'
          + '<w:pgBorders w:offsetFrom="page" w:display="firstPage" w:zOrder="back">'
          + '<w:top w:val="single" w:sz="24" w:space="24" w:color="4472C4"/>'
          + '<w:left w:val="single" w:sz="24" w:space="24" w:color="4472C4"/>'
          + '<w:bottom w:val="single" w:sz="24" w:space="24" w:color="4472C4"/>'
          + '<w:right w:val="single" w:sz="24" w:space="24" w:color="4472C4"/>'
          + '</w:pgBorders>'
          + '<w:bidi/>'
          + '</w:sectPr>',
      ),
    ],
    [
      'docGrid/textDirection/rtlGutter/formProt/noEndnote (round-trip fidelity audit, DXS round 2)',
      documentXml(
        '<w:p><w:r><w:t>Vertical text section</w:t></w:r></w:p>'
          + '<w:sectPr>'
          + '<w:formProt w:val="true"/>'
          + '<w:noEndnote/>'
          + '<w:textDirection w:val="tbRl"/>'
          + '<w:rtlGutter/>'
          + '<w:docGrid w:type="linesAndChars" w:linePitch="360" w:charSpace="0"/>'
          + '</w:sectPr>',
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

  // DEFER-5 / DXS-20 — field serialization (`buildFieldNodes` and friends in
  // documentWriter.ts) had no direct test coverage: the pre-existing DOCX
  // round-trip corpus (src/docx/__tests__/roundtrip.corpus.test.ts) happens
  // to exercise the RAW-PASSTHROUGH path via header-footer-page-numbers.docx's
  // real PAGE/NUMPAGES fields, but nothing exercised the STRUCTURAL-REBUILD
  // path (`raw` cleared by "Update field(s)"/"Update table of contents") at
  // all, nor gave an isolated, fast, field-focused regression test for
  // either path. These close that gap.
  describe('field serialization (DEFER-5 / DXS-20)', () => {
    it('round-trips an unmodified w:fldSimple field byte-for-byte via raw passthrough', () => {
      expectRoundTrip(
        documentXml(
          '<w:p><w:fldSimple w:instr="PAGE \\* MERGEFORMAT"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p><w:sectPr/>',
        ),
      )
    })

    it('round-trips an unmodified complex field (fldChar begin/separate/end + instrText) byte-for-byte via raw passthrough', () => {
      expectRoundTrip(
        documentXml(
          '<w:p>'
            + '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            + '<w:r><w:instrText xml:space="preserve">AUTHOR</w:instrText></w:r>'
            + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            + '<w:r><w:t>A. Author</w:t></w:r>'
            + '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
            + '</w:p><w:sectPr/>',
        ),
      )
    })

    it('structurally rebuilds a regenerated (raw-cleared) simple field as a valid w:fldSimple', () => {
      const document = createDocument([
        {
          kind: 'section',
          props: {},
          blocks: [{
            kind: 'paragraph',
            children: [{
              kind: 'field',
              fieldType: 'AUTHOR',
              instruction: 'AUTHOR',
              simple: true,
              result: [{ kind: 'run', children: [{ kind: 'text', value: 'New Author' }] }],
            }],
          }],
        },
      ])

      const xml = writeDocumentXml(document)
      expect(xml).toContain('<w:fldSimple w:instr="AUTHOR">')
      expect(xml).toContain('New Author')

      const reparsed = parseDocument(xml)
      const block = reparsed.sections[0]?.blocks[0]
      const field = block?.kind === 'paragraph' ? block.children[0] : undefined
      expect(field?.kind === 'field' && field.fieldType).toBe('AUTHOR')
      expect(field?.kind === 'field' && field.simple).toBe(true)
      expect(field?.kind === 'field' && field.result).toEqual([
        { kind: 'run', children: [{ kind: 'text', value: 'New Author', preserveSpace: true }] },
      ])
    })

    it('structurally rebuilds a regenerated (raw-cleared) complex field as a valid begin/instrText/separate/result/end run sequence', () => {
      const document = createDocument([
        {
          kind: 'section',
          props: {},
          blocks: [{
            kind: 'paragraph',
            children: [{
              kind: 'field',
              fieldType: 'REF',
              instruction: 'REF _Ref1 \\h',
              result: [{ kind: 'run', children: [{ kind: 'text', value: 'Section One' }] }],
            }],
          }],
        },
      ])

      const xml = writeDocumentXml(document)
      // Exactly 5 sibling <w:r> elements: begin, instrText, separate, result, end.
      expect(xml.match(/<w:r>/g)).toHaveLength(5)
      expect(xml).toContain('fldCharType="begin"')
      expect(xml).toContain('<w:instrText xml:space="preserve">REF _Ref1 \\h</w:instrText>')
      expect(xml).toContain('fldCharType="separate"')
      expect(xml).toContain('Section One')
      expect(xml).toContain('fldCharType="end"')

      const reparsed = parseDocument(xml)
      const block = reparsed.sections[0]?.blocks[0]
      const field = block?.kind === 'paragraph' ? block.children[0] : undefined
      expect(field?.kind === 'field' && field.fieldType).toBe('REF')
      expect(field?.kind === 'field' && field.instruction).toBe('REF _Ref1 \\h')
      expect(field?.kind === 'field' && field.simple).toBeUndefined()
      expect(field?.kind === 'field' && field.result).toEqual([
        { kind: 'run', children: [{ kind: 'text', value: 'Section One', preserveSpace: true }] },
      ])
    })

    it('preserves run formatting (props) on a regenerated complex field\'s result', () => {
      const document = createDocument([
        {
          kind: 'section',
          props: {},
          blocks: [{
            kind: 'paragraph',
            children: [{
              kind: 'field',
              fieldType: 'AUTHOR',
              instruction: 'AUTHOR',
              result: [{ kind: 'run', props: { bold: true }, children: [{ kind: 'text', value: 'Bold Author' }] }],
            }],
          }],
        },
      ])

      const xml = writeDocumentXml(document)
      expect(xml).toContain('<w:b/>')
      expect(xml).toContain('Bold Author')
    })
  })
})

// Round-trip fidelity audit, DXS round 2 follow-up — `WrapperPassthrough`
// (see `../../model/document.ts`'s doc comment): an unedited `w:sdt`/
// `mc:AlternateContent` now round-trips byte-for-byte instead of always
// being stripped to its inner content, but only for as long as an identity
// check on its captured content holds. These tests exercise both sides of
// that check directly against `writeDocumentXml`, rather than only through
// `expectRoundTrip`'s structural-AST comparison above (which can't tell
// "the wrapper round-tripped" from "the wrapper was stripped the same way
// on both sides").
describe('writeDocumentXml — wrapper-passthrough regions', () => {
  it('re-emits an unedited body-level w:sdt verbatim, including metadata Atlas never models', () => {
    const xml = documentXml(
      '<w:sdt>'
        + '<w:sdtPr><w:id w:val="7"/><w:alias w:val="Reviewer"/><w:tag w:val="atlasTag"/></w:sdtPr>'
        + '<w:sdtContent><w:p><w:r><w:t>Control text</w:t></w:r></w:p></w:sdtContent>'
        + '</w:sdt><w:sectPr/>',
    )
    const parsed = parseDocument(xml)

    const written = writeDocumentXml(parsed)

    expect(written).toContain(
      '<w:sdt><w:sdtPr><w:id w:val="7"/><w:alias w:val="Reviewer"/><w:tag w:val="atlasTag"/></w:sdtPr>'
        + '<w:sdtContent><w:p><w:r><w:t>Control text</w:t></w:r></w:p></w:sdtContent></w:sdt>',
    )
  })

  it('falls back to stripping the w:sdt wrapper once its paragraph has been replaced (simulating an edit)', () => {
    const xml = documentXml(
      '<w:sdt>'
        + '<w:sdtPr><w:id w:val="7"/><w:alias w:val="Reviewer"/></w:sdtPr>'
        + '<w:sdtContent><w:p><w:r><w:t>Control text</w:t></w:r></w:p></w:sdtContent>'
        + '</w:sdt><w:sectPr/>',
    )
    const parsed = parseDocument(xml)

    // A real edit rebuilds the touched paragraph as a new object (Atlas's
    // edit pipeline never mutates one in place) — a shallow clone breaks
    // the region's identity check the same way without needing a whole
    // editor-command round trip.
    const edited: Document = {
      ...parsed,
      sections: parsed.sections.map((section) => ({
        ...section,
        blocks: section.blocks.map((block) => ({ ...block })),
      })),
    }

    const written = writeDocumentXml(edited)

    expect(written).not.toContain('<w:sdt>')
    expect(written).not.toContain('w:alias')
    expect(written).toContain('Control text')
  })

  it('re-emits an unedited run-level mc:AlternateContent verbatim, including the mc:Fallback branch it would otherwise drop', () => {
    const xml = documentXml(
      '<w:p><w:r>'
        + '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        + '<mc:Choice Requires="wps"><w:drawing/></mc:Choice>'
        + '<mc:Fallback><w:pict><v:rect/></w:pict></mc:Fallback>'
        + '</mc:AlternateContent>'
        + '</w:r></w:p><w:sectPr/>',
    )
    const parsed = parseDocument(xml)

    const written = writeDocumentXml(parsed)

    expect(written).toContain('<mc:Fallback><w:pict><v:rect/></w:pict></mc:Fallback>')
  })

  it('falls back to mc:Choice-only once the run holding mc:AlternateContent has been replaced (simulating an edit)', () => {
    const xml = documentXml(
      '<w:p><w:r>'
        + '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        + '<mc:Choice Requires="wps"><w:drawing/></mc:Choice>'
        + '<mc:Fallback><w:pict><v:rect/></w:pict></mc:Fallback>'
        + '</mc:AlternateContent>'
        + '</w:r></w:p><w:sectPr/>',
    )
    const parsed = parseDocument(xml)

    // The region's captured content is the *drawing* (the run's own child),
    // so the clone has to go one level deeper than the w:sdt case above:
    // cloning only the `Run` would leave its `children` array — and the
    // drawing inside it — as the exact same references, and the identity
    // check would (correctly) still consider that unedited.
    const edited: Document = {
      ...parsed,
      sections: parsed.sections.map((section) => ({
        ...section,
        blocks: section.blocks.map((block) =>
          block.kind === 'paragraph'
            ? {
                ...block,
                children: block.children.map((child) =>
                  child.kind === 'run'
                    ? { ...child, children: child.children.map((runChild) => ({ ...runChild }) as typeof runChild) }
                    : child,
                ),
              }
            : block,
        ),
      })),
    }

    const written = writeDocumentXml(edited)

    expect(written).not.toContain('mc:AlternateContent')
    expect(written).not.toContain('mc:Fallback')
    expect(written).toContain('<w:drawing/>')
  })
})

// DOCX-1 / DOCX-12 — theme fonts and the general `w:rPr` unknown-child
// passthrough. Each of these fails before the corresponding fix (the
// theme-only `w:rFonts` vanished entirely; the named effects/unknown child
// were parsed then silently dropped on save) and passes after.
describe('writeDocumentXml — DOCX-1 theme fonts / DOCX-12 rPr fidelity', () => {
  it('round-trips a run rFonts using only theme references (no literal font name)', () => {
    const xml = documentXml(
      '<w:p><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi" w:eastAsiaTheme="minorEastAsia"/></w:rPr>'
        + '<w:t>Themed</w:t></w:r></w:p><w:sectPr/>',
    )

    const written = writeDocumentXml(parseDocument(xml))

    expect(written).toContain(
      '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi" w:eastAsiaTheme="minorEastAsia"/>',
    )
    expectRoundTrip(xml)
  })

  it('round-trips a run rFonts using only literal font names', () => {
    const xml = documentXml(
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Arial" w:eastAsia="MS Mincho"/></w:rPr>'
        + '<w:t>Literal</w:t></w:r></w:p><w:sectPr/>',
    )

    const written = writeDocumentXml(parseDocument(xml))

    expect(written).toContain('<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Arial" w:eastAsia="MS Mincho"/>')
    expectRoundTrip(xml)
  })

  it('round-trips a run rFonts mixing literal names with theme references', () => {
    const xml = documentXml(
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/></w:rPr>'
        + '<w:t>Mixed</w:t></w:r></w:p><w:sectPr/>',
    )

    const written = writeDocumentXml(parseDocument(xml))

    expect(written).toContain('<w:rFonts w:ascii="Calibri" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>')
    expectRoundTrip(xml)
  })

  it('round-trips a style-referencing run (no literal rFonts) whose style itself carries a theme font', () => {
    // Guards against a regression where `w:rFonts` disappears specifically
    // when it carries *only* theme attributes and no literal ones at all —
    // the exact shape Normal's own default run formatting uses in the
    // overwhelming majority of real Word documents.
    const xml = documentXml(
      '<w:p><w:r><w:rPr><w:rFonts w:asciiTheme="majorHAnsi"/></w:rPr><w:t>x</w:t></w:r></w:p><w:sectPr/>',
    )

    const written = writeDocumentXml(parseDocument(xml))

    expect(written).toContain('<w:rPr><w:rFonts w:asciiTheme="majorHAnsi"/></w:rPr>')
  })

  it('round-trips each newly-modeled character effect', () => {
    const xml = documentXml(
      '<w:p><w:r><w:rPr>'
        + '<w:outline/><w:emboss/><w:imprint/>'
        + '<w:bdr w:val="single" w:sz="4" w:space="1" w:color="FF0000"/>'
        + '<w:em w:val="dot"/>'
        + '<w:w w:val="150"/>'
        + '</w:rPr><w:t>Effects</w:t></w:r></w:p><w:sectPr/>',
    )

    const written = writeDocumentXml(parseDocument(xml))

    expect(written).toContain('<w:outline/>')
    expect(written).toContain('<w:emboss/>')
    expect(written).toContain('<w:imprint/>')
    expect(written).toContain('<w:bdr w:val="single" w:sz="4" w:space="1" w:color="FF0000"/>')
    expect(written).toContain('<w:em w:val="dot"/>')
    expect(written).toContain('<w:w w:val="150"/>')
    expectRoundTrip(xml)
  })

  it('round-trips w:effect, w:eastAsianLayout, and w:fitText via the general unknown-child passthrough', () => {
    const xml = documentXml(
      '<w:p><w:r><w:rPr>'
        + '<w:effect w:val="sparkle"/>'
        + '<w:eastAsianLayout w:id="1" w:combine="1"/>'
        + '<w:fitText w:val="2880" w:id="2"/>'
        + '</w:rPr><w:t>x</w:t></w:r></w:p><w:sectPr/>',
    )

    const written = writeDocumentXml(parseDocument(xml))

    expect(written).toContain('<w:effect w:val="sparkle"/>')
    expect(written).toContain('<w:eastAsianLayout w:id="1" w:combine="1"/>')
    expect(written).toContain('<w:fitText w:val="2880" w:id="2"/>')
    expectRoundTrip(xml)
  })

  it('reinserts a deliberately invented, wholly unknown w:rPr child at its correct schema position', () => {
    // `w:atlasTestUnknown` is not a real OOXML element — invented so this
    // test can only pass via the *general* passthrough (not a longer fixed
    // list of specifically-recognized names). Placed between `w:b` and
    // `w:i` in the source: per CT_RPr's sequence both are modeled and
    // adjacent, so a correct implementation must reproduce that exact
    // position, not merely "somewhere in rPr".
    const xml = documentXml(
      '<w:p><w:r><w:rPr><w:b/><w:atlasTestUnknown w:foo="bar"><w:child/></w:atlasTestUnknown><w:i/></w:rPr>'
        + '<w:t>x</w:t></w:r></w:p><w:sectPr/>',
    )

    const written = writeDocumentXml(parseDocument(xml))

    expect(written).toContain('<w:b/><w:atlasTestUnknown w:foo="bar"><w:child/></w:atlasTestUnknown><w:i/>')
    expectRoundTrip(xml)
  })

  it('appends a trailing unknown w:rPr child (no following modeled sibling) at the end', () => {
    const xml = documentXml(
      '<w:p><w:r><w:rPr><w:b/><w:atlasTrailingUnknown/></w:rPr><w:t>x</w:t></w:r></w:p><w:sectPr/>',
    )

    const written = writeDocumentXml(parseDocument(xml))

    expect(written).toContain('<w:rPr><w:b/><w:atlasTrailingUnknown/></w:rPr>')
    expectRoundTrip(xml)
  })

  it('preserves relative order of two invented unknown children anchored to the same modeled sibling', () => {
    const xml = documentXml(
      '<w:p><w:r><w:rPr><w:atlasFirstUnknown/><w:atlasSecondUnknown/><w:i/></w:rPr><w:t>x</w:t></w:r></w:p><w:sectPr/>',
    )

    const written = writeDocumentXml(parseDocument(xml))

    expect(written).toContain('<w:rPr><w:atlasFirstUnknown/><w:atlasSecondUnknown/><w:i/></w:rPr>')
    expectRoundTrip(xml)
  })
})
