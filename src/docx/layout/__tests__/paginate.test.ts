import { describe, expect, it, vi } from 'vitest'

import type { FontMetrics } from '../../fonts'
import type {
  Document,
  FooterReference,
  HeaderReference,
  ParaProps,
  Section,
  SectionProps,
  Table,
} from '../../model'
import { twip } from '../../model'

import { paginate } from '../paginate'
import type { Page } from '../pageTypes'
import type { FontResolver, LineBox } from '../types'

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

function createDocument(sections: ReadonlyArray<Section>): Document {
  return {
    kind: 'document',
    sections,
    styles: new Map(),
    numbering: new Map(),
    defaults: {
      paragraph: {
        spacing: {
          line: twip(400),
          lineRule: 'exact',
        },
      },
    },
    comments: new Map(),
    footnotes: new Map(),
    endnotes: new Map(),
    headers: new Map(),
    footers: new Map(),
  }
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
