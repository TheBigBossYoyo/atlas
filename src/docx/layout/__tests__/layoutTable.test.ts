import { describe, expect, it, vi } from 'vitest'

vi.mock('../../fonts', () => ({
  measureRun: (text: string, metrics: { advanceWidth: number }) => metrics.advanceWidth * text.length,
  measureFragmentPt: () => null,
  measureLineMetricsPt: () => null,
}))

import type {
  Block,
  Paragraph,
  Style,
  Table,
  TableCell,
  TableProps,
  TableRow,
  TableRowProps,
  Twip,
  Width,
} from '../../model'
import { hexColor, pct, twip } from '../../model'

import { layoutTable } from '../layoutTable'
import type { FontResolver } from '../types'

const fontResolver: FontResolver = async () =>
  ({
    unitsPerEm: 1000,
    ascender: 800,
    descender: -200,
    lineGap: 0,
    advanceWidth: 5,
  }) as unknown as Awaited<ReturnType<FontResolver>>

describe('layoutTable', () => {
  it('uses tblGrid widths for fixed layouts', async () => {
    const result = await layoutTable({
      table: createTable({
        props: { tblLayout: 'fixed' },
        tblGrid: [twip(100), twip(200), twip(100)],
        rows: [createRow([createCell('a'), createCell('b'), createCell('c')])],
      }),
      availableWidthPt: 100,
      fontResolver,
    })

    expect(result.columnWidthsPt).toEqual([5, 10, 5])
  })

  it('sets fixed-layout width to the sum of grid columns', async () => {
    const result = await layoutTable({
      table: createTable({
        props: { tblLayout: 'fixed' },
        tblGrid: [twip(100), twip(200), twip(100)],
        rows: [createRow([createCell('a'), createCell('b'), createCell('c')])],
      }),
      availableWidthPt: 500,
      fontResolver,
    })

    expect(result.widthPt).toBe(20)
  })

  it('uses max content widths when autofit content fits available width', async () => {
    const result = await layoutTable({
      table: createTable({
        rows: [createRow([createCell('aaaa'), createCell('bb')])],
      }),
      availableWidthPt: 40,
      fontResolver,
    })

    expect(result.columnWidthsPt).toEqual([20, 10])
    expect(result.widthPt).toBe(30)
  })

  it('shrinks autofit columns proportionally toward min widths', async () => {
    const result = await layoutTable({
      table: createTable({
        rows: [createRow([createCell('aaaa aaaa'), createCell('bb bb')])],
      }),
      availableWidthPt: 55,
      fontResolver,
    })

    expect(result.columnWidthsPt[0]).toBeCloseTo(35.625)
    expect(result.columnWidthsPt[1]).toBeCloseTo(19.375)
    expect(result.widthPt).toBeCloseTo(55)
  })

  it('never shrinks autofit columns below min widths', async () => {
    const result = await layoutTable({
      table: createTable({
        rows: [createRow([createCell('aaaa aaaa'), createCell('bb bb')])],
      }),
      availableWidthPt: 25,
      fontResolver,
    })

    expect(result.columnWidthsPt).toEqual([20, 10])
    expect(result.widthPt).toBe(30)
  })

  it('uses the sum of spanned columns for gridSpan cells', async () => {
    const result = await layoutTable({
      table: createTable({
        props: { tblLayout: 'fixed' },
        tblGrid: [twip(100), twip(200), twip(100)],
        rows: [createRow([createCell('span', { gridSpan: 2 }), createCell('tail')])],
      }),
      availableWidthPt: 100,
      fontResolver,
    })

    expect(result.rows[0].cells[0].gridSpan).toBe(2)
    expect(result.rows[0].cells[0].widthPt).toBe(15)
  })

  it('marks vMerge restart cells as merge starts', async () => {
    const result = await layoutTable({
      table: createTable({
        props: { tblLayout: 'fixed' },
        tblGrid: [twip(100)],
        rows: [createRow([createCell('aaaa aaaa', { vMerge: 'restart' })])],
      }),
      availableWidthPt: 100,
      fontResolver,
    })

    expect(result.rows[0].cells[0].vMergeStart).toBe(true)
    expect(result.rows[0].cells[0].vMergeContinue).toBe(false)
  })

  it('ignores vMerge continue cells when computing row height', async () => {
    const result = await layoutTable({
      table: createTable({
        props: { tblLayout: 'fixed' },
        tblGrid: [twip(100), twip(100)],
        rows: [
          createRow([createCell('aaaa aaaa', { vMerge: 'restart' }), createCell('x')]),
          createRow([createCell('aaaa aaaa', { vMerge: 'continue' }), createCell([])]),
        ],
      }),
      availableWidthPt: 100,
      fontResolver,
    })

    expect(result.rows[1].heightPt).toBe(0)
    expect(result.rows[1].cells[0].vMergeContinue).toBe(true)
  })

  it('preserves width for vMerge continue cells', async () => {
    const result = await layoutTable({
      table: createTable({
        props: { tblLayout: 'fixed' },
        tblGrid: [twip(100), twip(200)],
        rows: [
          createRow([createCell('top', { vMerge: 'restart' }), createCell('x')]),
          createRow([createCell('bottom', { vMerge: 'continue' }), createCell('y')]),
        ],
      }),
      availableWidthPt: 100,
      fontResolver,
    })

    expect(result.rows[1].cells[0].widthPt).toBe(5)
  })

  it('counts first-row table headers for repeatHeaderRowCount', async () => {
    const result = await layoutTable({
      table: createTable({
        rows: [
          createRow([createCell('head')], { tblHeader: true }),
          createRow([createCell('body')]),
        ],
      }),
      availableWidthPt: 100,
      fontResolver,
    })

    expect(result.repeatHeaderRowCount).toBe(1)
    expect(result.rows[0].isHeader).toBe(true)
    expect(result.rows[1].isHeader).toBe(false)
  })

  it('stops header repetition at the first non-header row', async () => {
    const result = await layoutTable({
      table: createTable({
        rows: [
          createRow([createCell('head-1')], { tblHeader: true }),
          createRow([createCell('head-2')], { tblHeader: true }),
          createRow([createCell('body')]),
          createRow([createCell('late-head')], { tblHeader: true }),
        ],
      }),
      availableWidthPt: 100,
      fontResolver,
    })

    expect(result.repeatHeaderRowCount).toBe(2)
    expect(result.rows[2].isHeader).toBe(false)
    expect(result.rows[3].isHeader).toBe(false)
  })

  it('wraps overflowing cell content and grows row height', async () => {
    const result = await layoutTable({
      table: createTable({
        props: { tblLayout: 'fixed' },
        tblGrid: [twip(400)],
        rows: [createRow([createCell('aa aa')])],
      }),
      availableWidthPt: 100,
      fontResolver,
    })

    expect(result.rows[0].cells[0].contentLines).toHaveLength(2)
    expect(result.rows[0].heightPt).toBe(22)
  })

  it('uses only padding for empty-cell row height', async () => {
    const result = await layoutTable({
      table: createTable({
        props: { tblLayout: 'fixed' },
        tblGrid: [twip(100)],
        rows: [
          createRow([
            createCell([], {
              tcMar: {
                top: createDxaWidth(40),
                bottom: createDxaWidth(60),
              },
            }),
          ]),
        ],
      }),
      availableWidthPt: 100,
      fontResolver,
    })

    expect(result.rows[0].cells[0].contentLines).toHaveLength(0)
    expect(result.rows[0].heightPt).toBe(5)
  })

  it('inherits table margins and lets cell margins override them', async () => {
    const result = await layoutTable({
      table: createTable({
        props: {
          tblLayout: 'fixed',
          tblCellMar: {
            top: createDxaWidth(40),
            right: createDxaWidth(40),
            bottom: createDxaWidth(40),
            left: createDxaWidth(40),
          },
        },
        tblGrid: [twip(100)],
        rows: [
          createRow([
            createCell([], {
              tcMar: {
                top: createDxaWidth(80),
                left: createDxaWidth(60),
              },
            }),
          ]),
        ],
      }),
      availableWidthPt: 100,
      fontResolver,
    })

    expect(result.rows[0].cells[0].paddingPt).toEqual({
      top: 4,
      right: 2,
      bottom: 2,
      left: 3,
    })
  })

  it('uses tcW dxa as a preferred autofit width', async () => {
    const result = await layoutTable({
      table: createTable({
        rows: [createRow([createCell([], { tcW: createDxaWidth(200) })])],
      }),
      availableWidthPt: 100,
      fontResolver,
    })

    expect(result.columnWidthsPt).toEqual([10])
  })

  it('uses tcW pct as a preferred autofit width', async () => {
    const result = await layoutTable({
      table: createTable({
        rows: [createRow([createCell([], { tcW: createPctWidth(2500) })])],
      }),
      availableWidthPt: 100,
      fontResolver,
    })

    expect(result.columnWidthsPt).toEqual([50])
  })

  it('uses tblW pct to limit the autofit target width', async () => {
    const result = await layoutTable({
      table: createTable({
        props: { tblW: createPctWidth(2500) },
        rows: [createRow([createCell('aa aa aa'), createCell('bb bb bb')])],
      }),
      availableWidthPt: 100,
      fontResolver,
    })

    expect(result.columnWidthsPt).toEqual([25, 25])
    expect(result.widthPt).toBe(50)
  })

  it('falls back to top vertical alignment by default', async () => {
    const result = await layoutTable({
      table: createTable({
        rows: [createRow([createCell('x')])],
      }),
      availableWidthPt: 100,
      fontResolver,
    })

    expect(result.rows[0].cells[0].vAlign).toBe('top')
  })

  describe('table style conditional formatting (D7 / DXP-06, DXL-08, DXS-05)', () => {
    const bandedStyles: ReadonlyMap<string, Style> = new Map([
      [
        'AtlasBandedTable',
        {
          id: 'AtlasBandedTable',
          type: 'table',
          conditionalFormats: new Map([
            ['firstRow', { cell: { shd: { fill: hexColor('4472C4') } } }],
            ['band1Horz', { cell: { shd: { fill: hexColor('D9E2F3') } } }],
          ]),
        },
      ],
    ])

    function bandedTable(rowCount: number, cellProps?: TableCell['props']): Table {
      return createTable({
        props: {
          tblStyle: 'AtlasBandedTable',
          tblLook: { firstRow: true, noHBand: false, noVBand: true },
        },
        rows: Array.from({ length: rowCount }, () => createRow([createCell('x', cellProps)])),
      })
    }

    it('does nothing when no styles map is supplied (backward compatible)', async () => {
      const result = await layoutTable({ table: bandedTable(1), availableWidthPt: 100, fontResolver })
      expect(result.rows[0].cells[0].shadingFill).toBeUndefined()
    })

    it('shades the header row from the firstRow conditional format', async () => {
      const result = await layoutTable({
        table: bandedTable(1),
        availableWidthPt: 100,
        fontResolver,
        styles: bandedStyles,
      })
      expect(result.rows[0].cells[0].shadingFill).toBe('#4472C4')
    })

    it('bands alternating data rows from band1Horz, leaving band2Horz rows unshaded', async () => {
      const result = await layoutTable({
        table: bandedTable(4),
        availableWidthPt: 100,
        fontResolver,
        styles: bandedStyles,
      })

      expect(result.rows[0].cells[0].shadingFill).toBe('#4472C4') // firstRow
      expect(result.rows[1].cells[0].shadingFill).toBeUndefined() // band2Horz (no block)
      expect(result.rows[2].cells[0].shadingFill).toBe('#D9E2F3') // band1Horz
      expect(result.rows[3].cells[0].shadingFill).toBeUndefined() // band2Horz (no block)
    })

    it('lets the cell\'s own direct shading override the conditional format', async () => {
      const result = await layoutTable({
        table: bandedTable(1, { shd: { fill: hexColor('FF0000') } }),
        availableWidthPt: 100,
        fontResolver,
        styles: bandedStyles,
      })
      expect(result.rows[0].cells[0].shadingFill).toBe('#FF0000')
    })

    it('falls back to the resolved table style for table-level borders/shading when the table sets none directly', async () => {
      const stylesWithTableLevel: ReadonlyMap<string, Style> = new Map([
        [
          'Shaded',
          {
            id: 'Shaded',
            type: 'table',
            table: { shading: { fill: hexColor('EEEEEE') } },
          },
        ],
      ])
      const result = await layoutTable({
        table: createTable({
          props: { tblStyle: 'Shaded' },
          rows: [createRow([createCell('x')])],
        }),
        availableWidthPt: 100,
        fontResolver,
        styles: stylesWithTableLevel,
      })
      expect(result.shadingFill).toBe('#EEEEEE')
    })
  })

  // DEFER-5 / DXS-20 regression guard: `collectRunsFromParagraphChild`/
  // `collectRunsFromHyperlinkChild` had no case for the new `'field'`
  // ParagraphChild kind, so a field's cached result text inside a table
  // cell (e.g. a DATE/AUTHOR field, or a cross-reference) silently
  // contributed no content at all.
  describe('field content in a cell (DEFER-5 / DXS-20)', () => {
    it('lays out a field\'s cached result text inside a table cell', async () => {
      const cellParagraph: Paragraph = {
        kind: 'paragraph',
        children: [
          {
            kind: 'field',
            fieldType: 'AUTHOR',
            instruction: 'AUTHOR',
            result: [{ kind: 'run', children: [{ kind: 'text', value: 'A. Author' }] }],
          },
        ],
      }

      const result = await layoutTable({
        table: createTable({
          props: { tblLayout: 'fixed' },
          tblGrid: [twip(400)],
          rows: [createRow([createCell([cellParagraph])])],
        }),
        availableWidthPt: 100,
        fontResolver,
      })

      const items = result.rows[0].cells[0].contentLines.flatMap((line) => line.items)
      expect(items.some((item) => item.kind === 'word' && item.text === 'A.')).toBe(true)
    })

    it('lays out a field nested inside a hyperlink inside a table cell', async () => {
      const cellParagraph: Paragraph = {
        kind: 'paragraph',
        children: [
          {
            kind: 'hyperlink',
            anchor: 'Top',
            children: [
              {
                kind: 'field',
                fieldType: 'PAGE',
                instruction: 'PAGE',
                result: [{ kind: 'run', children: [{ kind: 'text', value: '3' }] }],
              },
            ],
          },
        ],
      }

      const result = await layoutTable({
        table: createTable({
          props: { tblLayout: 'fixed' },
          tblGrid: [twip(400)],
          rows: [createRow([createCell([cellParagraph])])],
        }),
        availableWidthPt: 100,
        fontResolver,
      })

      const items = result.rows[0].cells[0].contentLines.flatMap((line) => line.items)
      expect(items.some((item) => item.kind === 'word' && item.text === '3')).toBe(true)
    })
  })
})

function createTable(input: {
  rows: ReadonlyArray<TableRow>
  props?: TableProps
  tblGrid?: ReadonlyArray<Twip>
}): Table {
  return {
    kind: 'table',
    props: input.props,
    rows: input.rows,
    tblGrid: input.tblGrid,
  }
}

function createRow(cells: ReadonlyArray<TableCell>, props?: TableRowProps): TableRow {
  return {
    kind: 'table-row',
    props,
    cells,
  }
}

function createCell(text: string, props?: TableCell['props']): TableCell
function createCell(blocks: ReadonlyArray<Block>, props?: TableCell['props']): TableCell
function createCell(textOrBlocks: string | ReadonlyArray<Block>, props?: TableCell['props']): TableCell {
  return {
    kind: 'table-cell',
    props,
    blocks: typeof textOrBlocks === 'string' ? [createParagraph(textOrBlocks)] : textOrBlocks,
  }
}

function createParagraph(text: string): Paragraph {
  return {
    kind: 'paragraph',
    children:
      text.length > 0
        ? [
            {
              kind: 'run',
              children: [{ kind: 'text', value: text }],
            },
          ]
        : [],
  }
}

function createDxaWidth(value: number): Width {
  return {
    type: 'dxa',
    value: twip(value),
  }
}

function createPctWidth(value: number): Width {
  return {
    type: 'pct',
    value: pct(value),
  }
}
