/**
 * SHEETFN-5 — formulas that cross sheets.
 *
 * `recalculateSheet` could only ever see one sheet, so `=Sheet2!A1` was
 * unresolvable and the cell displayed its own formula text. Resolving it needs
 * three things that only exist document-wide:
 *
 *   - a sheet-aware lookup;
 *   - ONE dependency graph spanning all sheets, so `Sheet1!A1 = Sheet2!B1` is
 *     evaluated after `Sheet2!B1` — ordering per sheet leaves one of them
 *     reading a stale value, and which one would depend on sheet order;
 *   - cycle detection across sheets, so a mutual reference reports `#REF!`
 *     instead of looping or silently keeping whatever was there.
 *
 * The chained cases below are the ones that fail if the graph is per-sheet but
 * pass if the sheets happen to be visited in a lucky order, so each is written
 * to depend on the UNLUCKY order.
 */
import { describe, expect, it } from 'vitest'

import {
  createDocument,
  recalculateDocument,
  setCellValue,
  type SpreadsheetDocument,
} from '../spreadsheetDocument'
import type { ParsedSheet } from '../../shared/spreadsheetGrid'

function sheet(
  name: string,
  rows: string[][],
  formulas: Record<string, string> = {},
): ParsedSheet {
  const formulaGrid = rows.map((row, r) => row.map((_, c) => formulas[`${r}:${c}`]))
  return {
    name,
    hidden: false,
    grid: {
      rows,
      colCount: Math.max(...rows.map(r => r.length), 1),
      merges: [],
      colWidthsPx: [],
      rowHeightsPx: [],
      formulas: formulaGrid,
    },
  }
}

const cell = (doc: SpreadsheetDocument, sheetIndex: number, row: number, col: number): string =>
  doc.sheets[sheetIndex].rows[row][col]

describe('cross-sheet references (SHEETFN-5)', () => {
  it('reads a cell from another sheet', () => {
    const doc = createDocument([
      sheet('Summary', [['']], { '0:0': 'Data!A1' }),
      sheet('Data', [['42']]),
    ])
    expect(cell(doc, 0, 0, 0)).toBe('42')
  })

  it('resolves on LOAD, not only after an edit', () => {
    // `createDocument` used to recalculate sheet by sheet, so a reference to a
    // sheet that had not been built yet could not resolve at all.
    const doc = createDocument([
      sheet('Summary', [['']], { '0:0': 'Data!A1*2' }),
      sheet('Data', [['21']]),
    ])
    expect(cell(doc, 0, 0, 0)).toBe('42')
  })

  it('reads a quoted sheet name with a space in it', () => {
    const doc = createDocument([
      sheet('Summary', [['']], { '0:0': "'Raw Data'!A1" }),
      sheet('Raw Data', [['hello']]),
    ])
    expect(cell(doc, 0, 0, 0)).toBe('hello')
  })

  it('matches a sheet name case-insensitively, as Excel does', () => {
    const doc = createDocument([
      sheet('Summary', [['']], { '0:0': 'dATA!A1' }),
      sheet('Data', [['7']]),
    ])
    expect(cell(doc, 0, 0, 0)).toBe('7')
  })

  it('sums a range on another sheet', () => {
    const doc = createDocument([
      sheet('Summary', [['']], { '0:0': 'SUM(Data!A1:A3)' }),
      sheet('Data', [['1'], ['2'], ['4']]),
    ])
    expect(cell(doc, 0, 0, 0)).toBe('7')
  })

  it('reports #REF! for a sheet that does not exist, not an empty cell', () => {
    // The distinction that matters: treating a missing sheet as blank would
    // make `=NoSuch!A1` quietly evaluate to 0.
    const doc = createDocument([sheet('Summary', [['']], { '0:0': 'NoSuch!A1' })])
    expect(cell(doc, 0, 0, 0)).toBe('#REF!')
  })

  it('reports #REF! for a range on a missing sheet', () => {
    const doc = createDocument([sheet('Summary', [['']], { '0:0': 'SUM(NoSuch!A1:A3)' })])
    expect(cell(doc, 0, 0, 0)).toBe('#REF!')
  })

  it('allows a sheet to qualify its OWN name', () => {
    const doc = createDocument([sheet('Data', [['5', '']], { '0:1': 'Data!A1*3' })])
    expect(cell(doc, 0, 0, 1)).toBe('15')
  })

  it('updates when the referenced sheet is edited', () => {
    let doc = createDocument([
      sheet('Summary', [['']], { '0:0': 'Data!A1' }),
      sheet('Data', [['1']]),
    ])
    expect(cell(doc, 0, 0, 0)).toBe('1')

    doc = setCellValue(doc, 1, 0, 0, '99')

    // The whole point: editing Sheet 2 has to update Sheet 1, which a
    // per-sheet recalculation never looks at.
    expect(cell(doc, 0, 0, 0)).toBe('99')
  })

  it('updates a chain that runs BACKWARDS through sheet order', () => {
    // Sheet 0 depends on sheet 1, which depends on sheet 2. Evaluating sheets
    // in file order without a global graph leaves sheet 0 one step stale.
    let doc = createDocument([
      sheet('First', [['']], { '0:0': 'Second!A1+1' }),
      sheet('Second', [['']], { '0:0': 'Third!A1+1' }),
      sheet('Third', [['10']]),
    ])
    expect(cell(doc, 1, 0, 0)).toBe('11')
    expect(cell(doc, 0, 0, 0)).toBe('12')

    doc = setCellValue(doc, 2, 0, 0, '20')
    expect(cell(doc, 1, 0, 0)).toBe('21')
    expect(cell(doc, 0, 0, 0)).toBe('22')
  })

  it('reports a circular reference that spans two sheets', () => {
    // Without cross-sheet cycle detection this either loops or settles on a
    // stale value, depending on visit order.
    const doc = createDocument([
      sheet('A', [['']], { '0:0': 'B!A1' }),
      sheet('B', [['']], { '0:0': 'A!A1' }),
    ])
    expect(cell(doc, 0, 0, 0)).toBe('#REF!')
    expect(cell(doc, 1, 0, 0)).toBe('#REF!')
  })

  it('mixes a local and a cross-sheet reference in one formula', () => {
    const doc = createDocument([
      sheet('Summary', [['10', '']], { '0:1': 'A1+Data!A1' }),
      sheet('Data', [['5']]),
    ])
    expect(cell(doc, 0, 0, 1)).toBe('15')
  })

  it('composes with the function library', () => {
    const doc = createDocument([
      sheet('Summary', [['']], { '0:0': 'IF(SUM(Data!A1:A2)>5,"big","small")' }),
      sheet('Data', [['4'], ['3']]),
    ])
    expect(cell(doc, 0, 0, 0)).toBe('big')
  })

  it('looks a value up on another sheet', () => {
    const doc = createDocument([
      sheet('Summary', [['']], { '0:0': 'VLOOKUP("Ink",Data!A1:B2,2)' }),
      sheet('Data', [['Paper', '10'], ['Ink', '25']]),
    ])
    expect(cell(doc, 0, 0, 0)).toBe('25')
  })
})

describe('recalculateDocument identity (SHEETFN-5)', () => {
  it('returns the same document when nothing changed', () => {
    // The editor's undo history detects a no-op by reference, so a fresh object
    // on every call would push a duplicate entry for every blocked action.
    const doc = createDocument([sheet('Data', [['1', '']], { '0:1': 'A1+1' })])
    expect(recalculateDocument(doc)).toBe(doc)
  })

  it('returns the same document when it has no formulas at all', () => {
    const doc = createDocument([sheet('Data', [['1', '2']])])
    expect(recalculateDocument(doc)).toBe(doc)
  })

  it('leaves untouched sheets as the same objects', () => {
    const doc = createDocument([
      sheet('Summary', [['']], { '0:0': 'Data!A1' }),
      sheet('Data', [['1']]),
      sheet('Unrelated', [['x']]),
    ])
    const after = setCellValue(doc, 1, 0, 0, '2')
    // The sheet nothing referenced is the SAME object, so React re-renders of
    // its grid are not triggered by an edit elsewhere.
    expect(after.sheets[2]).toBe(doc.sheets[2])
  })
})
