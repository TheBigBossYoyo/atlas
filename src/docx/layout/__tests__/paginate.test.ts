import { describe, expect, it, vi } from 'vitest'

import type { FontMetrics } from '../../fonts'
import type {
  Document,
  Endnote,
  Footnote,
  FooterReference,
  HeaderReference,
  NumberingDef,
  ParaProps,
  Paragraph,
  Section,
  SectionProps,
  Style,
  Table,
} from '../../model'
import { halfPoint, hexColor, twip } from '../../model'

import { MARKER_RUN_INDEX } from '../listMarkers'
import { paginate } from '../paginate'
import type { Page, PageLineRef } from '../pageTypes'
import type { FontResolver, LineBox, LineItem } from '../types'

describe('paginate', () => {
  it('returns one page for a short one-paragraph document', async () => {
    const pages = await paginate({
      document: createDocument([createSection([createParagraph(1)])]),
      fontResolver: createFontResolver(),
    })

    expect(pages).toHaveLength(1)
    expect(pageLineCounts(pages)).toEqual([1])
  })

  it('splits a 100-line paragraph across four pages', async () => {
    const pages = await paginate({
      document: createDocument([
        createSection([createParagraph(100)], {
          pageHeightPt: 600,
        }),
      ]),
      fontResolver: createFontResolver(),
    })

    expect(pages).toHaveLength(4)
    expect(paragraphLineCounts(pages, 0)).toEqual([30, 30, 30, 10])
  })

  it('honors pageBreakBefore on the second paragraph', async () => {
    const pages = await paginate({
      document: createDocument([
        createSection([
          createParagraph(1),
          createParagraph(1, { pageBreakBefore: true }),
        ]),
      ]),
      fontResolver: createFontResolver(),
    })

    expect(pages).toHaveLength(2)
    expect(paragraphLineCounts(pages, 0)).toEqual([1, 0])
    expect(paragraphLineCounts(pages, 1)).toEqual([0, 1])
  })

  it('moves a keep-with-next pair when only one line remains at the bottom', async () => {
    const pages = await paginate({
      document: createDocument([
        createSection(
          [
            createParagraph(5),
            createParagraph(1, { keepNext: true }),
            createParagraph(5),
          ],
          { pageHeightPt: 120 },
        ),
      ]),
      fontResolver: createFontResolver(),
    })

    expect(pageLineCounts(pages)).toEqual([5, 6])
    expect(paragraphLineCounts(pages, 1)).toEqual([0, 1])
    expect(paragraphLineCounts(pages, 2)).toEqual([0, 5])
  })

  it('keeps a keep-with-next pair on the current page when both paragraphs fit', async () => {
    const pages = await paginate({
      document: createDocument([
        createSection(
          [
            createParagraph(4),
            createParagraph(1, { keepNext: true }),
            createParagraph(5),
          ],
          { pageHeightPt: 200 },
        ),
      ]),
      fontResolver: createFontResolver(),
    })

    expect(pages).toHaveLength(1)
    expect(pageLineCounts(pages)).toEqual([10])
  })

  it('moves a keepLines paragraph to the next page when it cannot fully fit', async () => {
    const pages = await paginate({
      document: createDocument([
        createSection(
          [
            createParagraph(5),
            createParagraph(10, { keepLines: true }),
          ],
          { pageHeightPt: 200 },
        ),
      ]),
      fontResolver: createFontResolver(),
    })

    expect(pageLineCounts(pages)).toEqual([5, 10])
    expect(paragraphLineCounts(pages, 1)).toEqual([0, 10])
  })

  it('allows a paragraph without keepLines to split across pages', async () => {
    const pages = await paginate({
      document: createDocument([
        createSection(
          [
            createParagraph(5),
            createParagraph(10),
          ],
          { pageHeightPt: 200 },
        ),
      ]),
      fontResolver: createFontResolver(),
    })

    expect(pageLineCounts(pages)).toEqual([10, 5])
    expect(paragraphLineCounts(pages, 1)).toEqual([5, 5])
  })

  it('applies widow control by splitting 11 lines into 9 and 2', async () => {
    const pages = await paginate({
      document: createDocument([
        createSection([createParagraph(11)], {
          pageHeightPt: 200,
        }),
      ]),
      fontResolver: createFontResolver(),
    })

    expect(paragraphLineCounts(pages, 0)).toEqual([9, 2])
  })

  it('allows a 10/1 split when widow control is disabled', async () => {
    const pages = await paginate({
      document: createDocument([
        createSection(
          [createParagraph(11, { widowControl: false })],
          {
            pageHeightPt: 200,
          },
        ),
      ]),
      fontResolver: createFontResolver(),
    })

    expect(paragraphLineCounts(pages, 0)).toEqual([10, 1])
  })

  it('avoids a single orphan line at the bottom by moving the paragraph to the next page', async () => {
    const pages = await paginate({
      document: createDocument([
        createSection(
          [
            createParagraph(9),
            createParagraph(11),
          ],
          { pageHeightPt: 200 },
        ),
      ]),
      fontResolver: createFontResolver(),
    })

    expect(pageLineCounts(pages)).toEqual([9, 9, 2])
    expect(paragraphLineCounts(pages, 1)).toEqual([0, 9, 2])
  })

  it('allows an orphan line when widow control is disabled', async () => {
    const pages = await paginate({
      document: createDocument([
        createSection(
          [
            createParagraph(9),
            createParagraph(11, { widowControl: false }),
          ],
          { pageHeightPt: 200 },
        ),
      ]),
      fontResolver: createFontResolver(),
    })

    expect(pageLineCounts(pages)).toEqual([10, 10])
    expect(paragraphLineCounts(pages, 1)).toEqual([1, 10])
  })

  it('flows lines through two columns before creating a new page', async () => {
    const pages = await paginate({
      document: createDocument([
        createSection(
          [createParagraph(7)],
          {
            pageHeightPt: 60,
            pageWidthPt: 300,
            columnCount: 2,
            columnSpacePt: 20,
          },
        ),
      ]),
      fontResolver: createFontResolver(),
    })

    expect(pages).toHaveLength(2)
    expect(pages[0].columns.map((column) => column.lines.length)).toEqual([3, 2])
    expect(pages[1].columns.map((column) => column.lines.length)).toEqual([2, 0])
  })

  it('converts page size and margins to points in the output page model', async () => {
    const pages = await paginate({
      document: createDocument([
        createSection([createParagraph(1)], {
          pageWidthPt: 500,
          pageHeightPt: 700,
          topMarginPt: 10,
          rightMarginPt: 20,
          bottomMarginPt: 30,
          leftMarginPt: 40,
          headerMarginPt: 50,
          footerMarginPt: 60,
          gutterPt: 15,
        }),
      ]),
      fontResolver: createFontResolver(),
    })

    expect(pages[0].sizePt).toEqual({ width: 500, height: 700 })
    expect(pages[0].marginsPt).toEqual({
      top: 10,
      right: 20,
      bottom: 30,
      left: 40,
      header: 50,
      footer: 60,
      gutter: 15,
    })
  })

  it('starts a new page for a nextPage section break even when the current page has space', async () => {
    const pages = await paginate({
      document: createDocument([
        createSection([createParagraph(1)]),
        createSection([createParagraph(1)], { type: 'nextPage' }),
      ]),
      fontResolver: createFontResolver(),
    })

    expect(pages).toHaveLength(2)
    expect(pages[0].sectionIndex).toBe(0)
    expect(pages[1].sectionIndex).toBe(1)
  })

  it('pads to the next odd page for an oddPage section break', async () => {
    const pages = await paginate({
      document: createDocument([
        createSection([createParagraph(1)]),
        createSection([createParagraph(1)], { type: 'oddPage' }),
      ]),
      fontResolver: createFontResolver(),
    })

    expect(pages).toHaveLength(3)
    expect(pageLineCounts([pages[1]])).toEqual([0])
    expect(pageLineCounts([pages[2]])).toEqual([1])
  })

  it('forces a new page right after a manual page break, even when everything would otherwise fit (D10/DXL-06)', async () => {
    const paragraph = createParagraphWithManualBreak(2, 'page', 3)

    const pages = await paginate({
      // Default section is 200pt tall; 5 lines at the doc-default 20pt
      // line height (100pt) would easily fit on one page without D10.
      document: createDocument([createSection([paragraph])]),
      fontResolver: createFontResolver(),
    })

    expect(pages).toHaveLength(2)
    expect(paragraphLineCounts(pages, 0)).toEqual([2, 3])
  })

  it('advances to the next column for a manual column break when one remains (D10/DXL-06)', async () => {
    const paragraph = createParagraphWithManualBreak(1, 'column', 1)

    const pages = await paginate({
      document: createDocument([createSection([paragraph], { columnCount: 2 })]),
      fontResolver: createFontResolver(),
    })

    expect(pages).toHaveLength(1)
    expect(pages[0].columns.map((column) => column.lines.length)).toEqual([1, 1])
  })

  it('opens a new page for a manual column break when already on the last column (D10/DXL-06)', async () => {
    const paragraph = createParagraphWithManualBreak(2, 'column', 3)

    const pages = await paginate({
      document: createDocument([createSection([paragraph], { columnCount: 1 })]),
      fontResolver: createFontResolver(),
    })

    expect(pages).toHaveLength(2)
    expect(paragraphLineCounts(pages, 0)).toEqual([2, 3])
  })

  it("flows a continuous section break's content onto the previous section's page instead of an unwanted extra page break (D9/DXL-07)", async () => {
    const pages = await paginate({
      document: createDocument([
        createSection([createParagraph(1)]),
        createSection([createParagraph(1)], { type: 'continuous' }),
      ]),
      fontResolver: createFontResolver(),
    })

    expect(pages).toHaveLength(1)
    expect(pageLineCounts(pages)).toEqual([2])
  })

  it('recomputes column geometry for a continuous section break that changes column count (D9/DXL-07)', async () => {
    const pages = await paginate({
      document: createDocument([
        createSection([createParagraph(1)]),
        createSection([createParagraph(1)], { type: 'continuous', columnCount: 2 }),
      ]),
      fontResolver: createFontResolver(),
    })

    expect(pages).toHaveLength(1)
    expect(pages[0].columns).toHaveLength(2)
    expect(pageLineCounts(pages)).toEqual([2])
  })

  it('advances to the next column for a nextColumn section break when one remains (D9/DXL-07)', async () => {
    const pages = await paginate({
      document: createDocument([
        createSection([createParagraph(1)], { columnCount: 2 }),
        createSection([createParagraph(1)], { type: 'nextColumn', columnCount: 2 }),
      ]),
      fontResolver: createFontResolver(),
    })

    expect(pages).toHaveLength(1)
    expect(pages[0].columns.map((column) => column.lines.length)).toEqual([1, 1])
  })

  it('opens a new page for a nextColumn section break when already on the last column (D9/DXL-07)', async () => {
    const pages = await paginate({
      document: createDocument([
        createSection([createParagraph(1)]),
        createSection([createParagraph(1)], { type: 'nextColumn' }),
      ]),
      fontResolver: createFontResolver(),
    })

    expect(pages).toHaveLength(2)
    expect(pageLineCounts(pages)).toEqual([1, 1])
  })

  it('reserves header and footer line height from the page content area', async () => {
    const headerReference: HeaderReference = {
      id: 'header-default',
      type: 'default',
    }
    const footerReference: FooterReference = {
      id: 'footer-default',
      type: 'default',
    }
    const headerFooterLines = new Map<string, ReadonlyArray<LineBox>>([
      ['header-default', [createHeaderFooterLine(20)]],
      ['footer-default', [createHeaderFooterLine(20)]],
    ])

    const pages = await paginate({
      document: createDocument([
        createSection(
          [createParagraph(5)],
          {
            pageHeightPt: 120,
            headerReferences: [headerReference],
            footerReferences: [footerReference],
          },
        ),
      ]),
      fontResolver: createFontResolver(),
      headerFooterLines,
    })

    expect(pages).toHaveLength(2)
    expect(pages[0].headerLines).toHaveLength(1)
    expect(pages[0].footerLines).toHaveLength(1)
    expect(paragraphLineCounts(pages, 0)).toEqual([3, 2])
  })

  it('resolves a run inheriting a paragraph style chain, including its linked character style (D1/DXP-01/DXL-01)', async () => {
    const styles = makeStyles([
      {
        id: 'Normal',
        type: 'paragraph',
        isDefault: true,
      },
      {
        id: 'Heading1',
        type: 'paragraph',
        basedOn: 'Normal',
        linked: 'Heading1Char',
        run: { bold: true, sz: halfPoint(48) },
      },
      {
        id: 'Heading1Char',
        type: 'character',
        linked: 'Heading1',
        run: { color: hexColor('FF0000') },
      },
    ])

    const headingParagraph: Paragraph = {
      kind: 'paragraph',
      props: { pStyle: 'Heading1' },
      children: [{ kind: 'run', children: [{ kind: 'text', value: 'Title text' }] }],
    }

    const pages = await paginate({
      document: createDocument([createSection([headingParagraph])], styles),
      fontResolver: createFontResolver(),
    })

    const item = firstWordItem(pages)
    expect(item?.runProps.bold).toBe(true)
    expect(item?.runProps.sz).toBe(halfPoint(48))
    expect(item?.runProps.color).toBe('FF0000')
  })

  it("resolves a run's own rStyle instead of the enclosing paragraph's pStyle (D1/DXP-02 fix)", async () => {
    const styles = makeStyles([
      {
        id: 'Heading1',
        type: 'paragraph',
        run: { bold: true },
      },
      {
        id: 'Emphasis',
        type: 'character',
        run: { italic: true },
      },
    ])

    const paragraph: Paragraph = {
      kind: 'paragraph',
      // Deliberately give the paragraph a DIFFERENT style than the run's own
      // character style, so a bug that resolves against pStyle instead of
      // the run's rStyle would pick up `bold` from Heading1 and miss
      // `italic` from Emphasis entirely.
      props: { pStyle: 'Heading1' },
      children: [
        { kind: 'run', props: { rStyle: 'Emphasis' }, children: [{ kind: 'text', value: 'emphasized' }] },
      ],
    }

    const pages = await paginate({
      document: createDocument([createSection([paragraph])], styles),
      fontResolver: createFontResolver(),
    })

    const item = firstWordItem(pages)
    expect(item?.runProps.italic).toBe(true)
  })

  it('renders a right-aligned paragraph flush with the far edge of the column (D2/DXL-02)', async () => {
    const paragraph = createTextParagraph('aaaa', { jc: 'end' })

    const pages = await paginate({
      document: createDocument([createSection([paragraph], { pageWidthPt: 400 })]),
      fontResolver: createFontResolver(),
    })

    const lineRef = firstLineRef(pages)
    // 4 chars * 5.5pt (11pt default font, 500/1000 em advance) = 22pt wide;
    // a 400pt-wide column with no margins pushes it to leftPt = 400 - 22.
    expect(lineRef?.line.width).toBe(22)
    expect(lineRef?.leftPt).toBe(378)
  })

  it('renders a centered paragraph with equal space on both sides (D2/DXL-02)', async () => {
    const paragraph = createTextParagraph('aaaa', { jc: 'center' })

    const pages = await paginate({
      document: createDocument([createSection([paragraph], { pageWidthPt: 400 })]),
      fontResolver: createFontResolver(),
    })

    const lineRef = firstLineRef(pages)
    expect(lineRef?.leftPt).toBe((400 - 22) / 2)
  })

  it('offsets a paragraph by its left indent (D2/DXL-02)', async () => {
    const paragraph = createTextParagraph('aaaa', { ind: { left: twip(200) } })

    const pages = await paginate({
      document: createDocument([createSection([paragraph], { pageWidthPt: 400 })]),
      fontResolver: createFontResolver(),
    })

    const lineRef = firstLineRef(pages)
    expect(lineRef?.leftPt).toBe(10)
  })

  it('leaves a plain left-aligned paragraph flush with the column edge', async () => {
    const paragraph = createTextParagraph('aaaa')

    const pages = await paginate({
      document: createDocument([createSection([paragraph], { pageWidthPt: 400 })]),
      fontResolver: createFontResolver(),
    })

    const lineRef = firstLineRef(pages)
    expect(lineRef?.leftPt).toBe(0)
  })

  it('applies the larger of adjacent spacing.after/spacing.before between two paragraphs (D2/DXL-05)', async () => {
    const paragraphs = [
      createTextParagraph('first', { spacing: { after: twip(120) } }),
      createTextParagraph('second', { spacing: { before: twip(240) } }),
    ]

    const pages = await paginate({
      document: createDocument([createSection(paragraphs, { pageHeightPt: 500 })]),
      fontResolver: createFontResolver(),
    })

    const lineRefs = allLineRefs(pages)
    // Doc defaults pin every line's height at 20pt (exact spacing) — see
    // createDocument's defaults — so the gap is directly observable as the
    // difference between the second paragraph's top and 20pt.
    expect(lineRefs[0]?.topPt).toBe(0)
    expect(lineRefs[1]?.topPt).toBe(20 + 12)
  })

  it('suppresses spacing.before for the very first paragraph on a page (D2/DXL-05)', async () => {
    const paragraph = createTextParagraph('first', { spacing: { before: twip(240) } })

    const pages = await paginate({
      document: createDocument([createSection([paragraph], { pageHeightPt: 500 })]),
      fontResolver: createFontResolver(),
    })

    expect(firstLineRef(pages)?.topPt).toBe(0)
  })

  it('suppresses spacing between two contextualSpacing paragraphs sharing a style (D2/DXL-05)', async () => {
    const paragraphs = [
      createTextParagraph('first', { pStyle: 'ListParagraph', spacing: { after: twip(120) } }),
      createTextParagraph('second', {
        pStyle: 'ListParagraph',
        contextualSpacing: true,
        spacing: { before: twip(240) },
      }),
    ]

    const pages = await paginate({
      document: createDocument([createSection(paragraphs, { pageHeightPt: 500 })]),
      fontResolver: createFontResolver(),
    })

    const lineRefs = allLineRefs(pages)
    expect(lineRefs[1]?.topPt).toBe(20)
  })

  it('generates and increments decimal markers across paragraphs sharing a numId (D3/DXP-07/DXL-04)', async () => {
    const numbering = new Map<string, NumberingDef>([
      [
        '1',
        {
          numId: '1',
          levels: new Map([
            [0, { level: 0, format: 'decimal', text: { value: '%1.', placeholders: [1] }, suffix: 'tab' }],
          ]),
        },
      ],
    ])

    const paragraphs = [
      createTextParagraph('first item', {
        numPr: { numId: '1', ilvl: 0 },
        ind: { left: twip(720), hanging: twip(360) },
      }),
      createTextParagraph('second item', {
        numPr: { numId: '1', ilvl: 0 },
        ind: { left: twip(720), hanging: twip(360) },
      }),
    ]

    const pages = await paginate({
      document: createDocument([createSection(paragraphs, { pageWidthPt: 400 })], undefined, numbering),
      fontResolver: createFontResolver(),
    })

    const lineRefs = allLineRefs(pages)
    expect(markerWordOf(lineRefs[0])?.text).toBe('1.')
    expect(markerWordOf(lineRefs[1])?.text).toBe('2.')
  })

  it('positions the marker via hanging pull-back and lands its tab at the text-start indent (D3/DXL-04)', async () => {
    const numbering = new Map<string, NumberingDef>([
      [
        '1',
        {
          numId: '1',
          levels: new Map([
            [0, { level: 0, format: 'decimal', text: { value: '%1.', placeholders: [1] }, suffix: 'tab' }],
          ]),
        },
      ],
    ])

    const paragraph = createTextParagraph('item text', {
      numPr: { numId: '1', ilvl: 0 },
      ind: { left: twip(720), hanging: twip(360) },
    })

    const pages = await paginate({
      document: createDocument([createSection([paragraph], { pageWidthPt: 400 })], undefined, numbering),
      fontResolver: createFontResolver(),
    })

    const lineRef = firstLineRef(pages)
    const markerItem = markerWordOf(lineRef)
    const tabItem = lineRef?.line.items.find((item) => item.kind === 'tab')
    const firstRealWord = lineRef?.line.items.find(
      (item): item is Extract<LineItem, { kind: 'word' }> => item.kind === 'word' && item.runIndex === 0,
    )

    // Line 0 is pulled back by the 18pt hanging amount from the 36pt (720
    // twip) base left indent, so the marker starts at 18pt...
    expect(lineRef?.leftPt).toBe(18)
    // ...and the marker + its tab together land exactly back at the 36pt
    // text-start indent, where the real paragraph text begins.
    expect((lineRef?.leftPt ?? 0) + (markerItem?.width ?? 0) + (tabItem?.width ?? 0)).toBe(36)
    expect(firstRealWord?.text).toBe('item')
  })

  it("falls back to the numbering level's own indent when the paragraph sets none (D3)", async () => {
    const numbering = new Map<string, NumberingDef>([
      [
        '1',
        {
          numId: '1',
          levels: new Map([
            [
              0,
              {
                level: 0,
                format: 'decimal',
                text: { value: '%1.', placeholders: [1] },
                suffix: 'tab',
                paragraph: { ind: { left: twip(720), hanging: twip(360) } },
              },
            ],
          ]),
        },
      ],
    ])

    const paragraph = createTextParagraph('item', { numPr: { numId: '1', ilvl: 0 } })

    const pages = await paginate({
      document: createDocument([createSection([paragraph], { pageWidthPt: 400 })], undefined, numbering),
      fontResolver: createFontResolver(),
    })

    expect(firstLineRef(pages)?.leftPt).toBe(18)
  })

  it('renders a bulleted paragraph marker without disturbing the real run\'s own index (D3)', async () => {
    const numbering = new Map<string, NumberingDef>([
      [
        '1',
        {
          numId: '1',
          levels: new Map([
            [0, { level: 0, format: 'bullet', text: { value: '•', placeholders: [] }, suffix: 'tab' }],
          ]),
        },
      ],
    ])

    const paragraph = createTextParagraph('bulleted', {
      numPr: { numId: '1', ilvl: 0 },
      ind: { left: twip(720), hanging: twip(360) },
    })

    const pages = await paginate({
      document: createDocument([createSection([paragraph], { pageWidthPt: 400 })], undefined, numbering),
      fontResolver: createFontResolver(),
    })

    const lineRef = firstLineRef(pages)
    expect(markerWordOf(lineRef)?.text).toBe('•')
    const realWord = lineRef?.line.items.find(
      (item): item is Extract<LineItem, { kind: 'word' }> => item.kind === 'word' && item.runIndex === 0,
    )
    expect(realWord?.text).toBe('bulleted')
  })

  it('uses the injected tableLayout callback for table height', async () => {
    const tableLayout = vi.fn(async () => [40])
    const pages = await paginate({
      document: createDocument([
        createSection(
          [
            createParagraph(4),
            createTable(),
          ],
          { pageHeightPt: 100 },
        ),
      ]),
      fontResolver: createFontResolver(),
      tableLayout,
    })

    expect(tableLayout).toHaveBeenCalledTimes(1)
    expect(pageLineCounts(pages)).toEqual([4, 1])
    expect(paragraphLineCounts(pages, 1)).toEqual([0, 1])
  })
})

describe('paginate — headers/footers/vAlign (D11 milestone 1, D24/DXL-17)', () => {
  it('builds header/footer content directly from document.headers/document.footers', async () => {
    const pages = await paginate({
      document: createDocument(
        [
          createSection([createParagraph(1)], {
            headerReferences: [{ id: 'h1', type: 'default' }],
          }),
        ],
        undefined,
        undefined,
        { headers: new Map([['h1', { kind: 'header', id: 'h1', blocks: [createTextParagraph('Header Text')] }]]) },
      ),
      fontResolver: createFontResolver(),
    })

    expect(pages[0]?.headerLines).toHaveLength(1)
    expect(lineText(pages[0]?.headerLines[0])).toBe('Header Text')
  })

  it('ignores an even-typed header reference when evenAndOddHeaders is off', async () => {
    const document = createDocument(
      [
        createSection(
          [createParagraph(1), createParagraph(1, { pageBreakBefore: true })],
          {
            headerReferences: [
              { id: 'default', type: 'default' },
              { id: 'even', type: 'even' },
            ],
          },
        ),
      ],
      undefined,
      undefined,
      {
        headers: new Map([
          ['default', { kind: 'header', id: 'default', blocks: [createTextParagraph('Default')] }],
          ['even', { kind: 'header', id: 'even', blocks: [createTextParagraph('Even')] }],
        ]),
      },
    )

    const withoutSetting = await paginate({ document, fontResolver: createFontResolver() })
    expect(lineText(withoutSetting[1]?.headerLines[0])).toBe('Default')

    const withSetting = await paginate({ document, fontResolver: createFontResolver(), evenAndOddHeaders: true })
    expect(lineText(withSetting[1]?.headerLines[0])).toBe('Even')
  })

  it('centers page content vertically when the section is vAlign=center', async () => {
    const pages = await paginate({
      document: createDocument([
        createSection([createParagraph(1)], { pageHeightPt: 100, vAlign: 'center' }),
      ]),
      fontResolver: createFontResolver(),
    })

    // One 20pt line on a 100pt-tall page: 80pt slack, centered = 40pt offset.
    expect(firstLineRef(pages)?.topPt).toBe(40)
  })

  it('pushes page content to the bottom when the section is vAlign=bottom', async () => {
    const pages = await paginate({
      document: createDocument([
        createSection([createParagraph(1)], { pageHeightPt: 100, vAlign: 'bottom' }),
      ]),
      fontResolver: createFontResolver(),
    })

    expect(firstLineRef(pages)?.topPt).toBe(80)
  })

  it('does not offset a page vAlign=top (or the default) content', async () => {
    const pages = await paginate({
      document: createDocument([createSection([createParagraph(1)], { pageHeightPt: 100 })]),
      fontResolver: createFontResolver(),
    })

    expect(firstLineRef(pages)?.topPt).toBe(0)
  })
})

describe('paginate — footnotes (D11 milestones 2-3, 5)', () => {
  it('renders a footnote reference as a numbered, superscripted marker', async () => {
    const pages = await paginate({
      document: createDocument(
        [createSection([createFootnoteRefParagraph('See note', 'fn1')])],
        undefined,
        undefined,
        { footnotes: new Map([['fn1', createFootnote('fn1', 'Footnote body text')]]) },
      ),
      fontResolver: createFontResolver(),
    })

    const marker = findNoteRefItem(pages)
    expect(marker?.text).toBe('1')
    expect(marker?.runProps.vertAlign).toBe('superscript')
    expect(marker?.noteRef).toEqual({ kind: 'footnote', id: 'fn1' })
  })

  it('reserves bottom-of-page space and renders the footnote body with a separator', async () => {
    const pages = await paginate({
      document: createDocument(
        [createSection([createFootnoteRefParagraph('See note', 'fn1')], { pageHeightPt: 200 })],
        undefined,
        undefined,
        { footnotes: new Map([['fn1', createFootnote('fn1', 'Footnote body text')]]) },
      ),
      fontResolver: createFontResolver(),
    })

    expect(pages[0]?.hasFootnoteSeparator).toBe(true)
    expect(footnoteAreaText(pages, 0)).toContain('Footnote')
    // The footnote's own marker (mark + tab) precedes its body text.
    expect(findFootnoteAreaLines(pages, 0)[0]?.line.items[0]).toMatchObject({ text: '1' })
  })

  it('numbers two footnotes continuously in document order (default restart)', async () => {
    const pages = await paginate({
      document: createDocument(
        [
          createSection([
            createFootnoteRefParagraph('First', 'fn1'),
            createFootnoteRefParagraph('Second', 'fn2'),
          ], { pageHeightPt: 200 }),
        ],
        undefined,
        undefined,
        {
          footnotes: new Map([
            ['fn1', createFootnote('fn1', 'One')],
            ['fn2', createFootnote('fn2', 'Two')],
          ]),
        },
      ),
      fontResolver: createFontResolver(),
    })

    const marks = pages[0]?.footnoteLines
      .filter((line) => line.line.items[0]?.kind === 'word')
      .map((line) => (line.line.items[0] as Extract<LineItem, { kind: 'word' }>).text)
    expect(marks).toEqual(['1', '2'])
  })

  it('resets footnote numbering at the start of each section (eachSect restart)', async () => {
    const pages = await paginate({
      document: createDocument(
        [
          createSection([createFootnoteRefParagraph('First', 'fn1')], { pageHeightPt: 200 }),
          createSection([createFootnoteRefParagraph('Second', 'fn2')], { pageHeightPt: 200, type: 'nextPage' }),
        ],
        undefined,
        undefined,
        {
          footnotes: new Map([
            ['fn1', createFootnote('fn1', 'One')],
            ['fn2', createFootnote('fn2', 'Two')],
          ]),
        },
      ),
      fontResolver: createFontResolver(),
      footnoteNumbering: { restart: 'eachSect' },
    })

    expect((findFootnoteAreaLines(pages, 0)[0]?.line.items[0] as Extract<LineItem, { kind: 'word' }>)?.text).toBe('1')
    expect((findFootnoteAreaLines(pages, 1)[0]?.line.items[0] as Extract<LineItem, { kind: 'word' }>)?.text).toBe('1')
  })

  it('restarts footnote numbering on every page (eachPage restart)', async () => {
    const pages = await paginate({
      document: createDocument(
        [
          createSection(
            [
              createFootnoteRefParagraph('First', 'fn1'),
              createParagraph(1, { pageBreakBefore: true }),
              createFootnoteRefParagraph('Second', 'fn2'),
            ],
            { pageHeightPt: 200 },
          ),
        ],
        undefined,
        undefined,
        {
          footnotes: new Map([
            ['fn1', createFootnote('fn1', 'One')],
            ['fn2', createFootnote('fn2', 'Two')],
          ]),
        },
      ),
      fontResolver: createFontResolver(),
      footnoteNumbering: { restart: 'eachPage' },
    })

    expect(pages).toHaveLength(2)
    expect((findFootnoteAreaLines(pages, 0)[0]?.line.items[0] as Extract<LineItem, { kind: 'word' }>)?.text).toBe('1')
    expect((findFootnoteAreaLines(pages, 1)[0]?.line.items[0] as Extract<LineItem, { kind: 'word' }>)?.text).toBe('1')
    // The body markers were relabeled to match, not left at their
    // continuous placeholder values.
    const bodyMarks = pages.map((page) => findNoteRefItem([page])?.text)
    expect(bodyMarks).toEqual(['1', '1'])
  })
})

describe('paginate — endnotes (D11 milestone 4)', () => {
  it('renders endnote content once at the end of the document, not per page', async () => {
    const pages = await paginate({
      document: createDocument(
        [createSection([createEndnoteRefParagraph('See end', 'en1')])],
        undefined,
        undefined,
        { endnotes: new Map([['en1', createEndnote('en1', 'Endnote body text')]]) },
      ),
      fontResolver: createFontResolver(),
    })

    const marker = findNoteRefItem(pages)
    expect(marker?.text).toBe('1')
    expect(marker?.noteRef).toEqual({ kind: 'endnote', id: 'en1' })

    // Endnote content flows as ordinary placed lines (not a footnoteLines
    // area) at the end of the document.
    const allWords = pages
      .flatMap((page) => page.columns.flatMap((column) => column.lines))
      .flatMap((lineRef) => lineRef.line.items)
      .filter((item): item is Extract<LineItem, { kind: 'word' }> => item.kind === 'word')
      .map((item) => item.text)
      .join(' ')
    expect(allWords).toContain('Endnote')
    expect(allWords).toContain('body')
  })
})

function lineText(line: LineBox | undefined): string {
  if (line === undefined) {
    return ''
  }
  return line.items
    .map((item) => (item.kind === 'word' || item.kind === 'glyph-cluster' ? item.text : item.kind === 'space' ? ' ' : ''))
    .join('')
}

function createFontResolver(): FontResolver {
  const metrics: FontMetrics = {
    unitsPerEm: 1000,
    ascender: 800,
    descender: -200,
    lineGap: 0,
    xHeight: 500,
    capHeight: 700,
    advanceWidth: () => 500,
    hasGlyph: () => true,
  }

  return async () => metrics
}

function makeStyles(styles: ReadonlyArray<Style>): ReadonlyMap<string, Style> {
  const map = new Map<string, Style>()
  for (const style of styles) {
    map.set(style.id, style)
  }
  return map
}

function markerWordOf(lineRef: PageLineRef | undefined): Extract<LineItem, { kind: 'word' }> | undefined {
  return lineRef?.line.items.find(
    (item): item is Extract<LineItem, { kind: 'word' }> => item.kind === 'word' && item.runIndex === MARKER_RUN_INDEX,
  )
}

function firstWordItem(pages: ReadonlyArray<Page>): Extract<LineItem, { kind: 'word' }> | undefined {
  for (const page of pages) {
    for (const column of page.columns) {
      for (const lineRef of column.lines) {
        const word = lineRef.line.items.find((item): item is Extract<LineItem, { kind: 'word' }> => item.kind === 'word')
        if (word !== undefined) {
          return word
        }
      }
    }
  }
  return undefined
}

function createDocument(
  sections: ReadonlyArray<Section>,
  styles: ReadonlyMap<string, Style> = new Map(),
  numbering: ReadonlyMap<string, NumberingDef> = new Map(),
  notes: {
    footnotes?: ReadonlyMap<string, Footnote>
    endnotes?: ReadonlyMap<string, Endnote>
    headers?: Document['headers']
    footers?: Document['footers']
  } = {},
): Document {
  return {
    kind: 'document',
    sections,
    styles,
    numbering,
    defaults: {
      paragraph: {
        spacing: {
          line: twip(400),
          lineRule: 'exact',
        },
      },
    },
    comments: new Map(),
    footnotes: notes.footnotes ?? new Map(),
    endnotes: notes.endnotes ?? new Map(),
    headers: notes.headers ?? new Map(),
    footers: notes.footers ?? new Map(),
  }
}

function createFootnote(id: string, text: string): Footnote {
  return { kind: 'footnote', id, blocks: [createTextParagraph(text)] }
}

function createEndnote(id: string, text: string): Endnote {
  return { kind: 'endnote', id, blocks: [createTextParagraph(text)] }
}

function createFootnoteRefParagraph(text: string, footnoteId: string): Paragraph {
  return {
    kind: 'paragraph',
    props: {},
    children: [
      {
        kind: 'run',
        children: [
          { kind: 'text', value: text },
          { kind: 'footnote-reference', id: footnoteId },
        ],
      },
    ],
  }
}

function createEndnoteRefParagraph(text: string, endnoteId: string): Paragraph {
  return {
    kind: 'paragraph',
    props: {},
    children: [
      {
        kind: 'run',
        children: [
          { kind: 'text', value: text },
          { kind: 'endnote-reference', id: endnoteId },
        ],
      },
    ],
  }
}

function findFootnoteAreaLines(pages: ReadonlyArray<Page>, pageIndex: number): Page['footnoteLines'] {
  return pages[pageIndex]?.footnoteLines ?? []
}

function footnoteAreaText(pages: ReadonlyArray<Page>, pageIndex: number): string {
  return findFootnoteAreaLines(pages, pageIndex)
    .flatMap((footnoteLine) =>
      footnoteLine.line.items
        .filter((item): item is Extract<LineItem, { kind: 'word' | 'glyph-cluster' }> =>
          item.kind === 'word' || item.kind === 'glyph-cluster',
        )
        .map((item) => item.text),
    )
    .join('')
}

function findNoteRefItem(
  pages: ReadonlyArray<Page>,
): Extract<LineItem, { kind: 'word' }> | undefined {
  for (const page of pages) {
    for (const column of page.columns) {
      for (const lineRef of column.lines) {
        const item = lineRef.line.items.find(
          (candidate): candidate is Extract<LineItem, { kind: 'word' }> =>
            candidate.kind === 'word' && candidate.noteRef !== undefined,
        )
        if (item !== undefined) {
          return item
        }
      }
    }
  }
  return undefined
}

function createSection(
  blocks: Section['blocks'],
  options: {
    pageWidthPt?: number
    pageHeightPt?: number
    topMarginPt?: number
    rightMarginPt?: number
    bottomMarginPt?: number
    leftMarginPt?: number
    headerMarginPt?: number
    footerMarginPt?: number
    gutterPt?: number
    columnCount?: number
    columnSpacePt?: number
    type?: SectionProps['type']
    headerReferences?: ReadonlyArray<HeaderReference>
    footerReferences?: ReadonlyArray<FooterReference>
    vAlign?: SectionProps['vAlign']
  } = {},
): Section {
  return {
    kind: 'section',
    props: {
      pgSz: {
        w: twip((options.pageWidthPt ?? 400) * 20),
        h: twip((options.pageHeightPt ?? 200) * 20),
      },
      pgMar: {
        top: twip((options.topMarginPt ?? 0) * 20),
        right: twip((options.rightMarginPt ?? 0) * 20),
        bottom: twip((options.bottomMarginPt ?? 0) * 20),
        left: twip((options.leftMarginPt ?? 0) * 20),
        header: twip((options.headerMarginPt ?? 0) * 20),
        footer: twip((options.footerMarginPt ?? 0) * 20),
        gutter: twip((options.gutterPt ?? 0) * 20),
      },
      cols: {
        num: options.columnCount ?? 1,
        space: twip((options.columnSpacePt ?? 0) * 20),
        col: [],
      },
      type: options.type,
      headerReference: options.headerReferences,
      footerReference: options.footerReferences,
      vAlign: options.vAlign,
    },
    blocks,
  }
}

function createParagraph(lineCount: number, props: ParaProps = {}) {
  return {
    kind: 'paragraph' as const,
    props,
    children:
      lineCount > 1
        ? [
            {
              kind: 'run' as const,
              children: Array.from({ length: lineCount - 1 }, () => ({ kind: 'break' as const })),
            },
          ]
        : [],
  }
}

/**
 * A paragraph of `linesBefore + linesAfter` total lines (matching
 * `createParagraph`'s line-count convention — see D10's tests), with a
 * manual page/column break (`breakType`) ending the `linesBefore`-th line.
 */
function createParagraphWithManualBreak(
  linesBefore: number,
  breakType: 'page' | 'column',
  linesAfter: number,
): Paragraph {
  return {
    kind: 'paragraph',
    props: {},
    children: [
      {
        kind: 'run',
        children: [
          ...Array.from({ length: linesBefore - 1 }, () => ({ kind: 'break' as const })),
          { kind: 'break' as const, breakType },
          ...Array.from({ length: linesAfter - 1 }, () => ({ kind: 'break' as const })),
        ],
      },
    ],
  }
}

function createTextParagraph(text: string, props: ParaProps = {}): Paragraph {
  return {
    kind: 'paragraph',
    props,
    children: [{ kind: 'run', children: [{ kind: 'text', value: text }] }],
  }
}

function allLineRefs(pages: ReadonlyArray<Page>): ReadonlyArray<PageLineRef> {
  return pages.flatMap((page) => page.columns.flatMap((column) => column.lines))
}

function firstLineRef(pages: ReadonlyArray<Page>): PageLineRef | undefined {
  return allLineRefs(pages)[0]
}

function createTable(): Table {
  return {
    kind: 'table',
    rows: [],
  }
}

function createHeaderFooterLine(lineHeight: number): LineBox {
  return {
    items: [],
    width: 0,
    ascent: 0,
    descent: 0,
    lineHeight,
    isJustified: false,
    justificationStretch: 0,
  }
}

function pageLineCounts(pages: ReadonlyArray<Page>): number[] {
  return pages.map((page) => page.columns.reduce((count, column) => count + column.lines.length, 0))
}

function paragraphLineCounts(pages: ReadonlyArray<Page>, blockIndex: number): number[] {
  return pages.map((page) =>
    page.columns.reduce(
      (count, column) =>
        count + column.lines.filter((line) => line.paragraphPath[0] === blockIndex).length,
      0,
    ),
  )
}
