import { describe, expect, it } from 'vitest'

import { applyCommand } from '../commands'
import type { Command } from '../commandTypes'
import type {
  Comment,
  Document,
  Endnote,
  Footer,
  Footnote,
  Header,
  NumberingDef,
  Paragraph,
  Section,
  Style,
  Table,
  TableCell,
  TableRow,
} from '../../model'
import { twip } from '../../model'

// DXE-14 — command-level coverage for table structural editing.

describe('DXE-14 — insert/delete table rows', () => {
  it('inserts an empty row at the given index, matching the table column count', () => {
    const original = createDocumentWithTable(createGrid(2, 3))

    const result = applyCommand(original, {
      kind: 'insert-table-row',
      tablePath: [0],
      at: 1,
    })

    const table = tableAt(result.document, [0])
    expect(table.rows).toHaveLength(3)
    const inserted = table.rows[1] as TableRow
    expect(inserted.cells).toHaveLength(3)
    expect(inserted.cells.every((cell) => cell.kind === 'table-cell')).toBe(true)

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('clamps an out-of-range insert index to append at the end', () => {
    const original = createDocumentWithTable(createGrid(1, 2))

    const result = applyCommand(original, { kind: 'insert-table-row', tablePath: [0], at: 99 })

    const table = tableAt(result.document, [0])
    expect(table.rows).toHaveLength(2)
  })

  it('deletes a row and restores it exactly on undo', () => {
    const original = createDocumentWithTable(createGrid(3, 2), (rowIndex, cellIndex) =>
      `r${rowIndex}c${cellIndex}`,
    )

    const cmd: Command = { kind: 'delete-table-row', tablePath: [0], rowIndex: 1 }
    const result = applyCommand(original, cmd)

    const table = tableAt(result.document, [0])
    expect(table.rows).toHaveLength(2)
    expect(cellText(table, 0, 0)).toBe('r0c0')
    expect(cellText(table, 1, 0)).toBe('r2c0')

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('refuses to delete a table\'s only remaining row', () => {
    const original = createDocumentWithTable(createGrid(1, 2))

    expect(() => applyCommand(original, { kind: 'delete-table-row', tablePath: [0], rowIndex: 0 })).toThrow()
  })
})

describe('DXE-14 — insert/delete table columns', () => {
  it('inserts a column before the given index in every row and grows tblGrid', () => {
    const original = createDocumentWithTable(createGrid(2, 2), (rowIndex, cellIndex) =>
      `r${rowIndex}c${cellIndex}`,
    )

    const result = applyCommand(original, { kind: 'insert-table-column', tablePath: [0], at: 1 })

    const table = tableAt(result.document, [0])
    expect(table.tblGrid).toHaveLength(3)
    expect(cellText(table, 0, 0)).toBe('r0c0')
    expect(cellText(table, 0, 1)).toBe('') // freshly inserted
    expect(cellText(table, 0, 2)).toBe('r0c1')
    expect(cellText(table, 1, 2)).toBe('r1c1')

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('widens a merged cell instead of splitting it when the index falls inside its span', () => {
    const merged = createDocumentWithTable(createGrid(1, 3))
    const mergedResult = applyCommand(merged, {
      kind: 'merge-table-cells',
      tablePath: [0],
      rowIndex: 0,
      fromCellIndex: 0,
      toCellIndex: 1,
    })

    const result = applyCommand(mergedResult.document, {
      kind: 'insert-table-column',
      tablePath: [0],
      at: 1,
    })

    const table = tableAt(result.document, [0])
    expect(table.rows[0].kind).toBe('table-row')
    const row = table.rows[0] as TableRow
    expect(row.cells).toHaveLength(2)
    const firstCell = row.cells[0] as TableCell
    expect(firstCell.props?.gridSpan).toBe(3)
  })

  it('deletes a column, narrowing a merged cell that spans it instead of removing it', () => {
    const merged = createDocumentWithTable(createGrid(1, 3))
    const mergedResult = applyCommand(merged, {
      kind: 'merge-table-cells',
      tablePath: [0],
      rowIndex: 0,
      fromCellIndex: 0,
      toCellIndex: 1,
    })

    const result = applyCommand(mergedResult.document, {
      kind: 'delete-table-column',
      tablePath: [0],
      columnIndex: 0,
    })

    const table = tableAt(result.document, [0])
    const row = table.rows[0] as TableRow
    expect(row.cells).toHaveLength(2)
    const firstCell = row.cells[0] as TableCell
    expect(firstCell.props?.gridSpan).toBe(1)

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(mergedResult.document)
  })

  it('refuses to delete a table\'s only remaining column', () => {
    const original = createDocumentWithTable(createGrid(2, 1))

    expect(() =>
      applyCommand(original, { kind: 'delete-table-column', tablePath: [0], columnIndex: 0 }),
    ).toThrow()
  })
})

describe('DXE-14 — delete table', () => {
  it('removes the whole table block and restores it verbatim on undo', () => {
    const original = createDocumentWithTable(createGrid(2, 2))

    const result = applyCommand(original, { kind: 'delete-table', tablePath: [0] })

    expect(result.document.sections[0].blocks).toHaveLength(0)

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })
})

describe('DXE-14 — merge and split cells', () => {
  it('merges a range of cells into one with the combined gridSpan and concatenated content', () => {
    const original = createDocumentWithTable(createGrid(1, 3), (_row, cellIndex) => `c${cellIndex}`)

    const result = applyCommand(original, {
      kind: 'merge-table-cells',
      tablePath: [0],
      rowIndex: 0,
      fromCellIndex: 0,
      toCellIndex: 2,
    })

    const table = tableAt(result.document, [0])
    const row = table.rows[0] as TableRow
    expect(row.cells).toHaveLength(1)
    const merged = row.cells[0] as TableCell
    expect(merged.props?.gridSpan).toBe(3)
    expect(merged.blocks).toHaveLength(3)
    expect(merged.blocks.map((block) => paragraphText(block as Paragraph))).toEqual(['c0', 'c1', 'c2'])

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('rejects merging a single cell', () => {
    const original = createDocumentWithTable(createGrid(1, 2))

    expect(() =>
      applyCommand(original, {
        kind: 'merge-table-cells',
        tablePath: [0],
        rowIndex: 0,
        fromCellIndex: 0,
        toCellIndex: 0,
      }),
    ).toThrow()
  })

  it('splits a merged cell back into evenly distributed single-span cells', () => {
    const merged = createDocumentWithTable(createGrid(1, 4))
    const mergedResult = applyCommand(merged, {
      kind: 'merge-table-cells',
      tablePath: [0],
      rowIndex: 0,
      fromCellIndex: 0,
      toCellIndex: 3,
    })

    const result = applyCommand(mergedResult.document, {
      kind: 'split-table-cell',
      tablePath: [0],
      rowIndex: 0,
      cellIndex: 0,
    })

    const table = tableAt(result.document, [0])
    const row = table.rows[0] as TableRow
    expect(row.cells).toHaveLength(4)
    expect(row.cells.every((cell) => (cell as TableCell).props?.gridSpan === undefined)).toBe(true)

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(mergedResult.document)
  })

  it('splits a merged cell into a specific number of cells, distributing span evenly', () => {
    const merged = createDocumentWithTable(createGrid(1, 5))
    const mergedResult = applyCommand(merged, {
      kind: 'merge-table-cells',
      tablePath: [0],
      rowIndex: 0,
      fromCellIndex: 0,
      toCellIndex: 4,
    })

    const result = applyCommand(mergedResult.document, {
      kind: 'split-table-cell',
      tablePath: [0],
      rowIndex: 0,
      cellIndex: 0,
      into: 2,
    })

    const table = tableAt(result.document, [0])
    const row = table.rows[0] as TableRow
    expect(row.cells).toHaveLength(2)
    const spans = row.cells.map((cell) => (cell as TableCell).props?.gridSpan ?? 1)
    expect(spans.reduce((a, b) => a + b, 0)).toBe(5)
  })
})

describe('DXE-14 — resize a column', () => {
  it('updates tblGrid and the unspanned cells in that column', () => {
    const original = createDocumentWithTable(createGrid(2, 2))

    const result = applyCommand(original, {
      kind: 'resize-table-column',
      tablePath: [0],
      columnIndex: 0,
      widthTwips: 2000,
    })

    const table = tableAt(result.document, [0])
    expect(table.tblGrid?.[0]).toBe(twip(2000))
    const row0 = table.rows[0] as TableRow
    expect((row0.cells[0] as TableCell).props?.tcW).toEqual({ type: 'dxa', value: twip(2000) })

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('rejects a width below the minimum', () => {
    const original = createDocumentWithTable(createGrid(1, 1))

    expect(() =>
      applyCommand(original, { kind: 'resize-table-column', tablePath: [0], columnIndex: 0, widthTwips: 10 }),
    ).toThrow()
  })
})

describe('DXE-14 — table properties (apply-table-props)', () => {
  it("replaces the table's props wholesale and restores the original verbatim on undo", () => {
    const original = createDocumentWithTable(createGrid(1, 1))

    const result = applyCommand(original, {
      kind: 'apply-table-props',
      tablePath: [0],
      props: { jc: 'center', tblW: { type: 'dxa', value: twip(5000) } },
    })

    const table = tableAt(result.document, [0])
    expect(table.props).toEqual({ jc: 'center', tblW: { type: 'dxa', value: twip(5000) } })

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('clears props entirely when given undefined, and undo restores the prior props', () => {
    const withProps: Table = { ...createGrid(1, 1), props: { jc: 'end' } }
    const original = createDocumentWithTable(withProps)

    const result = applyCommand(original, { kind: 'apply-table-props', tablePath: [0], props: undefined })

    expect(tableAt(result.document, [0]).props).toBeUndefined()

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('throws when the table path does not resolve', () => {
    const original = createDocumentWithTable(createGrid(1, 1))

    expect(() =>
      applyCommand(original, { kind: 'apply-table-props', tablePath: [9], props: { jc: 'center' } }),
    ).toThrow()
  })
})

// ─── helpers ──────────────────────────────────────────────────────────────

function createGrid(rows: number, cols: number): Table {
  return {
    kind: 'table',
    tblGrid: Object.freeze(Array.from({ length: cols }, () => twip(1440))),
    rows: Object.freeze(
      Array.from({ length: rows }, () =>
        Object.freeze({
          kind: 'table-row',
          cells: Object.freeze(
            Array.from({ length: cols }, () =>
              Object.freeze({ kind: 'table-cell', blocks: Object.freeze([emptyParagraph()]) }),
            ),
          ),
        } satisfies TableRow),
      ),
    ),
  } satisfies Table
}

function withCellText(table: Table, text: (rowIndex: number, cellIndex: number) => string): Table {
  return {
    ...table,
    rows: table.rows.map((row, rowIndex) => {
      if (row.kind !== 'table-row') return row
      return {
        ...row,
        cells: row.cells.map((cell, cellIndex) => {
          if (cell.kind !== 'table-cell') return cell
          return { ...cell, blocks: [createParagraph(text(rowIndex, cellIndex))] }
        }),
      }
    }),
  }
}

function createParagraph(text: string): Paragraph {
  return Object.freeze({
    kind: 'paragraph',
    children: Object.freeze([
      Object.freeze({
        kind: 'run',
        children: Object.freeze([Object.freeze({ kind: 'text', value: text })]),
      }),
    ]),
  }) satisfies Paragraph
}

function emptyParagraph(): Paragraph {
  return Object.freeze({ kind: 'paragraph', children: Object.freeze([]) }) satisfies Paragraph
}

function createDocumentWithTable(
  table: Table,
  text?: (rowIndex: number, cellIndex: number) => string,
): Document {
  const finalTable = text !== undefined ? withCellText(table, text) : table
  const section = Object.freeze({
    kind: 'section',
    props: {},
    blocks: Object.freeze([finalTable]),
  }) satisfies Section

  return Object.freeze({
    kind: 'document',
    sections: Object.freeze([section]),
    styles: new Map<string, Style>(),
    numbering: new Map<string, NumberingDef>(),
    comments: new Map<string, Comment>(),
    footnotes: new Map<string, Footnote>(),
    endnotes: new Map<string, Endnote>(),
    headers: new Map<string, Header>(),
    footers: new Map<string, Footer>(),
  }) satisfies Document
}

function tableAt(document: Document, tablePath: readonly [number]): Table {
  const [index] = tablePath
  return document.sections[0].blocks[index] as Table
}

function cellText(table: Table, rowIndex: number, cellIndex: number): string {
  const row = table.rows[rowIndex]
  if (row === undefined || row.kind !== 'table-row') return ''
  const cell = row.cells[cellIndex]
  if (cell === undefined || cell.kind !== 'table-cell') return ''
  const paragraph = cell.blocks[0]
  if (paragraph === undefined || paragraph.kind !== 'paragraph') return ''
  return paragraphText(paragraph)
}

function paragraphText(paragraph: Paragraph): string {
  return paragraph.children
    .map((child) => (child.kind === 'run' ? child.children.map((c) => (c.kind === 'text' ? c.value : '')).join('') : ''))
    .join('')
}
