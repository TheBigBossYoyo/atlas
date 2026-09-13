import { describe, expect, it, vi } from 'vitest'

vi.mock('../../fonts', () => ({
  measureRun: (text: string, metrics: { advanceWidth: number }) => metrics.advanceWidth * text.length,
  measureFragmentPt: () => null,
  measureLineMetricsPt: () => null,
}))

import type {
  Block,
  Paragraph,
  Table,
  TableCell,
  TableProps,
  TableRow,
  TableRowProps,
  Twip,
  Width,
} from '../../model'
import { pct, twip } from '../../model'

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
