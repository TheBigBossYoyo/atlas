/**
 * SHEET-SORT-1 — sorting a block of rows by one column.
 *
 * The properties that matter are mostly about what moves WITH a row: its
 * formatting, its height, and its save-path bookkeeping. A sort that reordered
 * only the visible text would leave a cell's colour behind on the row that used
 * to be there, and would make the passthrough writer attribute an untouched
 * cell to the wrong original row.
 *
 * The ordering rules are deliberately the same three the formula evaluator's
 * comparison operators use, so a sort and a `>` in a formula cannot disagree
 * about which of two cells is larger.
 */
import { describe, expect, it } from 'vitest'

import {
  createDocument,
  formatForCell,
  setRangeFormat,
  sortRows,
  type SpreadsheetDocument,
} from '../spreadsheetDocument'
import type { ParsedSheet } from '../../shared/spreadsheetGrid'

function sheetOf(rows: string[][], formulas: Record<string, string> = {}): ParsedSheet {
  return {
    name: 'Data',
    hidden: false,
    grid: {
      rows,
      colCount: Math.max(...rows.map(r => r.length), 1),
      merges: [],
      colWidthsPx: [],
      rowHeightsPx: [],
      formulas: rows.map((row, r) => row.map((_, c) => formulas[`${r}:${c}`])),
    },
  }
}

const col = (doc: SpreadsheetDocument, index: number): string[] =>
  doc.sheets[0].rows.map(row => row[index])

describe('sortRows (SHEET-SORT-1)', () => {
  it('sorts ascending and descending by a column', () => {
    const doc = createDocument([sheetOf([['Ink'], ['Paper'], ['Card']])])

    expect(col(sortRows(doc, 0, { col: 0, direction: 'asc' }).document, 0)).toEqual(['Card', 'Ink', 'Paper'])
    expect(col(sortRows(doc, 0, { col: 0, direction: 'desc' }).document, 0)).toEqual(['Paper', 'Ink', 'Card'])
  })

  it('sorts numbers numerically, not as text', () => {
    // The classic bug: "10" sorting before "9" because it compares as text.
    const doc = createDocument([sheetOf([['9'], ['10'], ['2']])])
    expect(col(sortRows(doc, 0, { col: 0, direction: 'asc' }).document, 0)).toEqual(['2', '9', '10'])
  })

  it('puts numbers before text, matching the formula comparison rules', () => {
    const doc = createDocument([sheetOf([['apple'], ['3'], ['banana'], ['1']])])
    expect(col(sortRows(doc, 0, { col: 0, direction: 'asc' }).document, 0)).toEqual([
      '1',
      '3',
      'apple',
      'banana',
    ])
  })

  it('compares text case-insensitively', () => {
    const doc = createDocument([sheetOf([['beta'], ['Alpha'], ['gamma']])])
    expect(col(sortRows(doc, 0, { col: 0, direction: 'asc' }).document, 0)).toEqual([
      'Alpha',
      'beta',
      'gamma',
    ])
  })

  it('keeps blanks last in BOTH directions', () => {
    // An empty cell is absence, not a small value. Sorting descending with
    // blanks "largest" would bury the data under them.
    const doc = createDocument([sheetOf([['b'], [''], ['a']])])
    expect(col(sortRows(doc, 0, { col: 0, direction: 'asc' }).document, 0)).toEqual(['a', 'b', ''])
    expect(col(sortRows(doc, 0, { col: 0, direction: 'desc' }).document, 0)).toEqual(['b', 'a', ''])
  })

  it('moves the whole row, not just the sorted column', () => {
    // "Paper" before "Ink" so the sort has real work to do — the first version
    // of this test used an input that was already in order, which it would have
    // passed without moving anything.
    const doc = createDocument([
      sheetOf([
        ['Paper', '10'],
        ['Ink', '25'],
      ]),
    ])
    const sorted = sortRows(doc, 0, { col: 0, direction: 'asc' }).document
    expect(sorted.sheets[0].rows).toEqual([
      ['Ink', '25'],
      ['Paper', '10'],
    ])
  })

  it('leaves header rows in place', () => {
    const doc = createDocument([sheetOf([['Item'], ['Paper'], ['Ink']])])
    const sorted = sortRows(doc, 0, { col: 0, direction: 'asc', headerRows: 1 }).document
    expect(col(sorted, 0)).toEqual(['Item', 'Ink', 'Paper'])
  })

  it('is stable for rows that compare equal', () => {
    // Two rows with the same key keep their original order, so a second sort on
    // another column does not scramble the first — and a descending sort does
    // not reverse ties, which an unstable comparator would.
    const doc = createDocument([
      sheetOf([
        ['same', 'first'],
        ['same', 'second'],
        ['same', 'third'],
      ]),
    ])
    expect(col(sortRows(doc, 0, { col: 0, direction: 'asc' }).document, 1)).toEqual([
      'first',
      'second',
      'third',
    ])
    expect(col(sortRows(doc, 0, { col: 0, direction: 'desc' }).document, 1)).toEqual([
      'first',
      'second',
      'third',
    ])
  })

  it('refuses when the block holds a formula, and says why', () => {
    // Deliberate: Excel adjusts a moved formula's relative references, and
    // getting those rules wrong would silently produce a workbook with wrong
    // numbers in it. Refusing is the safe answer, and the UI can explain it.
    const doc = createDocument([sheetOf([['1'], ['2']], { '1:0': 'A1*2' })])
    const result = sortRows(doc, 0, { col: 0, direction: 'asc' })
    expect(result.sorted).toBe(false)
    expect(result.refusal).toBe('has-formulas')
    expect(result.document).toBe(doc)
  })

  it('reports nothing to sort for a single row', () => {
    const doc = createDocument([sheetOf([['only']])])
    const result = sortRows(doc, 0, { col: 0, direction: 'asc' })
    expect(result.sorted).toBe(false)
    expect(result.refusal).toBe('no-rows')
  })

  it('returns the same document when already in order', () => {
    // So a no-op sort does not add an undo entry, which the editor detects by
    // reference.
    const doc = createDocument([sheetOf([['a'], ['b']])])
    const result = sortRows(doc, 0, { col: 0, direction: 'asc' })
    expect(result.sorted).toBe(false)
    expect(result.document).toBe(doc)
  })
})

describe('what travels with a sorted row (SHEET-SORT-1)', () => {
  it('carries the row height', () => {
    const parsed = sheetOf([['b'], ['a']])
    const withHeights: ParsedSheet = { ...parsed, grid: { ...parsed.grid, rowHeightsPx: [40, 20] } }
    const sorted = sortRows(createDocument([withHeights]), 0, { col: 0, direction: 'asc' }).document
    // Row "a" was second with height 20 and is now first.
    expect(sorted.sheets[0].rowHeightsPx).toEqual([20, 40])
  })

  it('carries the save-path row sources', () => {
    // `rowSources` is how the passthrough writer knows which ORIGINAL row an
    // untouched cell came from. Sorting without permuting it would make the
    // writer copy the wrong cell's bytes — a silent data swap on save.
    const parsed = sheetOf([['b'], ['a']])
    const withSource: ParsedSheet = { ...parsed, sourcePath: 'xl/worksheets/sheet1.xml' }
    const sorted = sortRows(createDocument([withSource]), 0, { col: 0, direction: 'asc' }).document
    expect(sorted.sheets[0].rowSources).toEqual([1, 0])
  })

  it('carries formatting with the row it was applied to', () => {
    const doc = createDocument([sheetOf([['b'], ['a']])])
    // Make the SECOND row (which sorts to the top) bold.
    const formatted = setRangeFormat(doc, 0, { row0: 1, col0: 0, row1: 1, col1: 0 }, { bold: true })
    const sorted = sortRows(formatted, 0, { col: 0, direction: 'asc' }).document

    expect(formatForCell(sorted.sheets[0], 0, 0).bold).toBe(true)
    expect(formatForCell(sorted.sheets[0], 1, 0).bold).toBe(false)
  })
})
