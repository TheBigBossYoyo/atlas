/**
 * Spreadsheet editing (wave 3) — pure document-model operations.
 */
import { describe, expect, it } from 'vitest'

import {
  addSheet,
  createDocument,
  createDocumentFromRows,
  deleteColumnAt,
  deleteRowAt,
  deleteSheet,
  insertColumnAt,
  insertRowAt,
  pasteRange,
  recalculateSheet,
  renameSheet,
  setCellValue,
  type SpreadsheetDocument,
} from '../spreadsheetDocument'
import type { ParsedSheet } from '../../shared/spreadsheetGrid'

function sheetFixture(rows: string[][], overrides: Partial<ParsedSheet> = {}): ParsedSheet {
  const colCount = Math.max(...rows.map((r) => r.length), 0)
  return {
    name: 'Sheet1',
    hidden: false,
    grid: {
      rows,
      colCount,
      merges: [],
      colWidthsPx: [],
      rowHeightsPx: [],
      formulas: rows.map((r) => r.map(() => undefined)),
    },
    ...overrides,
  }
}

function basicDoc(): SpreadsheetDocument {
  return createDocument([
    sheetFixture([
      ['Name', 'Score'],
      ['Alice', '10'],
      ['Bob', '20'],
    ]),
  ])
}

describe('createDocument / createDocumentFromRows', () => {
  it('builds an editable document from parsed sheets', () => {
    const doc = basicDoc()
    expect(doc.sheets).toHaveLength(1)
    expect(doc.sheets[0].rows).toEqual([
      ['Name', 'Score'],
      ['Alice', '10'],
      ['Bob', '20'],
    ])
  })

  it('carries the freeze field through only when present', () => {
    const withFreeze = createDocument([sheetFixture([['a']], { freeze: { cols: 1, rows: 0 } })])
    expect(withFreeze.sheets[0].freeze).toEqual({ cols: 1, rows: 0 })

    const without = createDocument([sheetFixture([['a']])])
    expect(without.sheets[0].freeze).toBeUndefined()
  })

  it('builds a single-sheet document from plain CSV/TSV rows', () => {
    const doc = createDocumentFromRows(
      [
        ['a', 'b'],
        ['1', '2'],
      ],
      2,
    )
    expect(doc.sheets).toHaveLength(1)
    expect(doc.sheets[0].colCount).toBe(2)
  })
})

describe('createDocument — formula hydration on load', () => {
  it('evaluates a formula our evaluator supports even when the file cached a different (or no) value', () => {
    const doc = createDocument([
      sheetFixture(
        [
          ['1', '2'],
          ['', ''],
        ],
        {
          grid: {
            rows: [
              ['1', '2'],
              ['', ''],
            ],
            colCount: 2,
            merges: [],
            colWidthsPx: [],
            rowHeightsPx: [],
            formulas: [
              [undefined, undefined],
              ['A1+B1', undefined],
            ],
          },
        },
      ),
    ])
    expect(doc.sheets[0].rows[1][0]).toBe('3')
  })

  it('preserves the file-cached display text for a formula our evaluator cannot handle', () => {
    const doc = createDocument([
      sheetFixture([['looked up value']], {
        grid: {
          rows: [['looked up value']],
          colCount: 1,
          merges: [],
          colWidthsPx: [],
          rowHeightsPx: [],
          formulas: [['VLOOKUP(A1,B:C,2,FALSE)']],
        },
      }),
    ])
    expect(doc.sheets[0].rows[0][0]).toBe('looked up value')
  })

  it('falls back to the literal formula text for an unsupported formula with no cached value at all', () => {
    const doc = createDocument([
      sheetFixture([['']], {
        grid: {
          rows: [['']],
          colCount: 1,
          merges: [],
          colWidthsPx: [],
          rowHeightsPx: [],
          formulas: [['VLOOKUP(A1,B:C,2,FALSE)']],
        },
      }),
    ])
    expect(doc.sheets[0].rows[0][0]).toBe('=VLOOKUP(A1,B:C,2,FALSE)')
  })
})

describe('recalculateSheet — perf fast path (T2/DAT-07)', () => {
  it('returns the exact same sheet reference when it has no formulas at all (no row copy)', () => {
    const sheet = basicDoc().sheets[0]
    expect(recalculateSheet(sheet)).toBe(sheet)
  })

  it('still copies (and never mutates) rows when a formula is present', () => {
    const sheet = setCellValue(basicDoc(), 0, 0, 0, '=1+1').sheets[0]
    const originalRows = sheet.rows
    const result = recalculateSheet(sheet)
    expect(result).not.toBe(sheet)
    expect(sheet.rows).toBe(originalRows)
  })

  it('never clones a row that has no formula cell, even when another row in the same sheet does (no full-sheet copy on an unrelated edit)', () => {
    const doc = createDocument([sheetFixture([['1', ''], ['', ''], ['', '']])])
    const withFormula = setCellValue(doc, 0, 0, 1, '=A1+1').sheets[0]
    const untouchedRow1 = withFormula.rows[1]
    const untouchedRow2 = withFormula.rows[2]

    const recalculated = recalculateSheet(withFormula)

    // Row 0 (the formula's own row) is necessarily a fresh array; rows 1
    // and 2 have no formula cell anywhere in them and must keep the EXACT
    // SAME reference — recalculating must not degrade into an O(rows) (let
    // alone O(rows x cols)) copy of every row just because SOME row,
    // anywhere in the sheet, has a formula in it.
    expect(recalculated.rows[1]).toBe(untouchedRow1)
    expect(recalculated.rows[2]).toBe(untouchedRow2)
  })
})

/** Builds a `ParsedSheet` with explicit `rows`/`formulas` grids, for dependency-ordering tests that need formula cells at specific coordinates. */
function gridSheet(rows: string[][], formulas: ReadonlyArray<ReadonlyArray<string | undefined>>, colCount: number): ParsedSheet {
  return {
    name: 'Sheet1',
    hidden: false,
    grid: { rows, colCount, merges: [], colWidthsPx: [], rowHeightsPx: [], formulas },
  }
}

describe('recalculateSheet — dependency ordering (SHEET-5)', () => {
  it('computes a forward dependency (a cell references a LATER formula cell) fresh on the first pass', () => {
    // 5 rows x 3 cols. A1 (row 0, col 0) = "=C5+1"; C5 (row 4, col 2) is
    // ITSELF a formula ("=5+5"), not a static value — the exact shape that
    // stayed stale under the old row-major single pass, since A1 is visited
    // long before C5 in row-major order.
    const rows = Array.from({ length: 5 }, () => ['', '', ''])
    const formulas: (string | undefined)[][] = Array.from({ length: 5 }, () => [undefined, undefined, undefined])
    formulas[0][0] = 'C5+1'
    formulas[4][2] = '5+5'

    const doc = createDocument([gridSheet(rows, formulas, 3)])

    expect(doc.sheets[0].rows[4][2]).toBe('10')
    expect(doc.sheets[0].rows[0][0]).toBe('11')
  })

  it('still computes a backward dependency (a cell references an EARLIER formula cell) correctly', () => {
    const rows = [['', '']]
    const formulas: (string | undefined)[][] = [['2+3', 'A1+1']]

    const doc = createDocument([gridSheet(rows, formulas, 2)])

    expect(doc.sheets[0].rows[0][0]).toBe('5')
    expect(doc.sheets[0].rows[0][1]).toBe('6')
  })

  it('resolves a three-cell forward chain (A1 -> B1 -> C1) in one pass', () => {
    // Row-major visits A1, B1, C1 in that order, but A1 depends on B1
    // (later) which depends on C1 (later still) — two levels of look-ahead.
    const rows = [['', '', '']]
    const formulas: (string | undefined)[][] = [['B1+1', 'C1+1', '7']]

    const doc = createDocument([gridSheet(rows, formulas, 3)])

    expect(doc.sheets[0].rows[0]).toEqual(['9', '8', '7'])
  })

  it('recomputes correctly through a forward dependency on every subsequent edit, not just the first pass', () => {
    const rows = Array.from({ length: 5 }, () => ['', '', ''])
    const formulas: (string | undefined)[][] = Array.from({ length: 5 }, () => [undefined, undefined, undefined])
    formulas[0][0] = 'C5+1'
    formulas[4][2] = '5+5'
    let doc = createDocument([gridSheet(rows, formulas, 3)])
    expect(doc.sheets[0].rows[0][0]).toBe('11')

    // Editing C5's formula directly must still propagate to A1 immediately.
    doc = setCellValue(doc, 0, 4, 2, '=100')
    expect(doc.sheets[0].rows[4][2]).toBe('100')
    expect(doc.sheets[0].rows[0][0]).toBe('101')
  })
})

describe('recalculateSheet — circular references (SHEET-5)', () => {
  it('resolves a two-cell cycle (A1 <-> B1) to a defined error instead of hanging', () => {
    const rows = [['', '']]
    const formulas: (string | undefined)[][] = [['B1+1', 'A1+1']]

    const doc = createDocument([gridSheet(rows, formulas, 2)])

    expect(doc.sheets[0].rows[0][0]).toBe('#REF!')
    expect(doc.sheets[0].rows[0][1]).toBe('#REF!')
  })

  it('resolves a three-cell cycle (A1 -> B1 -> C1 -> A1) to a defined error instead of hanging', () => {
    const rows = [['', '', '']]
    const formulas: (string | undefined)[][] = [['B1+1', 'C1+1', 'A1+1']]

    const doc = createDocument([gridSheet(rows, formulas, 3)])

    expect(doc.sheets[0].rows[0]).toEqual(['#REF!', '#REF!', '#REF!'])
  })

  it('does not let a cycle elsewhere in the sheet stop unrelated formula cells from computing normally', () => {
    const rows = [['', '', '', '']]
    const formulas: (string | undefined)[][] = [['B1+1', 'A1+1', undefined, '2+3']]

    const doc = createDocument([gridSheet(rows, formulas, 4)])

    expect(doc.sheets[0].rows[0][0]).toBe('#REF!')
    expect(doc.sheets[0].rows[0][1]).toBe('#REF!')
    expect(doc.sheets[0].rows[0][3]).toBe('5')
  })
})

describe('setCellValue', () => {
  it('sets a plain value with no formula', () => {
    const doc = setCellValue(basicDoc(), 0, 1, 0, 'Carol')
    expect(doc.sheets[0].rows[1][0]).toBe('Carol')
    expect(doc.sheets[0].formulas[1][0]).toBeUndefined()
  })

  it('stores a leading-= input as a formula and computes its display text', () => {
    const doc = setCellValue(basicDoc(), 0, 1, 1, '=1+2')
    expect(doc.sheets[0].formulas[1][1]).toBe('1+2')
    expect(doc.sheets[0].rows[1][1]).toBe('3')
  })

  it('falls back to showing the literal formula when it cannot be evaluated', () => {
    const doc = setCellValue(basicDoc(), 0, 1, 1, '=NOTAFUNCTION(1)')
    expect(doc.sheets[0].rows[1][1]).toBe('=NOTAFUNCTION(1)')
  })

  it('recomputes a dependent formula elsewhere in the sheet when its referenced cell changes', () => {
    // Sum lives in its own column (C) so it never references its own cell.
    let doc = insertColumnAt(basicDoc(), 0, 2)
    doc = setCellValue(doc, 0, 2, 2, '=B2+B3')
    expect(doc.sheets[0].rows[2][2]).toBe('30')

    doc = setCellValue(doc, 0, 1, 1, '100')
    expect(doc.sheets[0].rows[2][2]).toBe('120')
  })

  it('treats a bare "=" as a plain value, not an (empty) formula', () => {
    const doc = setCellValue(basicDoc(), 0, 0, 0, '=')
    expect(doc.sheets[0].formulas[0][0]).toBeUndefined()
    expect(doc.sheets[0].rows[0][0]).toBe('=')
  })

  it('is a no-op out of bounds', () => {
    const original = basicDoc()
    expect(setCellValue(original, 0, 99, 0, 'x')).toBe(original)
    expect(setCellValue(original, 5, 0, 0, 'x')).toBe(original)
  })

  it('does not clobber another cell\'s unsupported-formula cached text when an unrelated cell is edited', () => {
    let doc = createDocument([
      sheetFixture([['looked up value', '1']], {
        grid: {
          rows: [['looked up value', '1']],
          colCount: 2,
          merges: [],
          colWidthsPx: [],
          rowHeightsPx: [],
          formulas: [['VLOOKUP(A1,B:C,2,FALSE)', undefined]],
        },
      }),
    ])

    doc = setCellValue(doc, 0, 0, 1, '2')

    expect(doc.sheets[0].rows[0][0]).toBe('looked up value')
    expect(doc.sheets[0].rows[0][1]).toBe('2')
  })

  it('never mutates the input document (immutability)', () => {
    const original = basicDoc()
    const originalRows = original.sheets[0].rows.map((row) => [...row])
    setCellValue(original, 0, 0, 0, 'changed')
    expect(original.sheets[0].rows).toEqual(originalRows)
  })

  it('keeps every untouched row\'s exact array reference (perf: no O(rows) copy for one cell)', () => {
    const original = basicDoc()
    const doc = setCellValue(original, 0, 1, 0, 'Carol')
    // Row 1 (the edited one) is a fresh array; rows 0 and 2 must be the
    // EXACT SAME reference as before — cloning every row of a 100k-row
    // sheet to change one cell does not scale (see the function's own
    // header comment).
    expect(doc.sheets[0].rows[0]).toBe(original.sheets[0].rows[0])
    expect(doc.sheets[0].rows[2]).toBe(original.sheets[0].rows[2])
    expect(doc.sheets[0].rows[1]).not.toBe(original.sheets[0].rows[1])
  })
})

describe('insertRowAt / deleteRowAt', () => {
  it('inserts an empty row at the given index', () => {
    const doc = insertRowAt(basicDoc(), 0, 1)
    expect(doc.sheets[0].rows).toEqual([
      ['Name', 'Score'],
      ['', ''],
      ['Alice', '10'],
      ['Bob', '20'],
    ])
  })

  it('deletes a row at the given index', () => {
    const doc = deleteRowAt(basicDoc(), 0, 1)
    expect(doc.sheets[0].rows).toEqual([
      ['Name', 'Score'],
      ['Bob', '20'],
    ])
  })

  it('refuses to delete the last remaining row', () => {
    const single = createDocument([sheetFixture([['only']])])
    expect(deleteRowAt(single, 0, 0)).toBe(single)
  })

  it('shifts a merge range down when a row is inserted above it', () => {
    const withMerge = createDocument([
      sheetFixture([
        ['a', 'a'],
        ['b', 'c'],
      ], {
        grid: {
          rows: [
            ['a', 'a'],
            ['b', 'c'],
          ],
          colCount: 2,
          merges: [{ r0: 1, c0: 0, r1: 1, c1: 1 }],
          colWidthsPx: [],
          rowHeightsPx: [],
          formulas: [[undefined, undefined], [undefined, undefined]],
        },
      }),
    ])

    const doc = insertRowAt(withMerge, 0, 0)
    expect(doc.sheets[0].merges).toEqual([{ r0: 2, c0: 0, r1: 2, c1: 1 }])
  })

  it('grows (rather than shifts) a merge range when a row is inserted strictly inside it', () => {
    const withMerge = createDocument([
      sheetFixture(
        [
          ['a', 'a'],
          ['a', 'a'],
          ['b', 'c'],
        ],
        {
          grid: {
            rows: [
              ['a', 'a'],
              ['a', 'a'],
              ['b', 'c'],
            ],
            colCount: 2,
            merges: [{ r0: 0, c0: 0, r1: 1, c1: 1 }],
            colWidthsPx: [],
            rowHeightsPx: [],
            formulas: [
              [undefined, undefined],
              [undefined, undefined],
              [undefined, undefined],
            ],
          },
        },
      ),
    ])

    // Row 1 lands strictly inside the r0:0-r1:1 merge — the merge should
    // grow to r1:2, not shift down whole (which would incorrectly leave a
    // gap between the merge and its own top row).
    const doc = insertRowAt(withMerge, 0, 1)
    expect(doc.sheets[0].merges).toEqual([{ r0: 0, c0: 0, r1: 2, c1: 1 }])
  })

  it('shrinks (rather than leaves unchanged) a merge range when its own bottom row is deleted', () => {
    const withMerge = createDocument([
      sheetFixture(
        [
          ['a', 'a'],
          ['a', 'a'],
          ['b', 'c'],
        ],
        {
          grid: {
            rows: [
              ['a', 'a'],
              ['a', 'a'],
              ['b', 'c'],
            ],
            colCount: 2,
            merges: [{ r0: 0, c0: 0, r1: 1, c1: 1 }],
            colWidthsPx: [],
            rowHeightsPx: [],
            formulas: [
              [undefined, undefined],
              [undefined, undefined],
              [undefined, undefined],
            ],
          },
        },
      ),
    ])

    // Row 1 is the merge's own BOTTOM row (r0:0-r1:1). Deleting it must
    // shrink the merge to a single row (r1:0), not leave r1 at its old
    // numeric value of 1 — which, after the delete, would point at the
    // OLD row 2 (now slid up to index 1) and incorrectly absorb it into
    // the merge even though it was never part of it.
    const doc = deleteRowAt(withMerge, 0, 1)
    expect(doc.sheets[0].merges).toEqual([{ r0: 0, c0: 0, r1: 0, c1: 1 }])
    expect(doc.sheets[0].rows).toEqual([
      ['a', 'a'],
      ['b', 'c'],
    ])
  })
})

describe('insertColumnAt / deleteColumnAt', () => {
  it('inserts an empty column and grows colCount', () => {
    const doc = insertColumnAt(basicDoc(), 0, 1)
    expect(doc.sheets[0].colCount).toBe(3)
    expect(doc.sheets[0].rows[1]).toEqual(['Alice', '', '10'])
  })

  it('deletes a column and shrinks colCount', () => {
    const doc = deleteColumnAt(basicDoc(), 0, 0)
    expect(doc.sheets[0].colCount).toBe(1)
    expect(doc.sheets[0].rows[1]).toEqual(['10'])
  })

  it('refuses to delete the last remaining column', () => {
    const single = createDocument([sheetFixture([['only']])])
    expect(deleteColumnAt(single, 0, 0)).toBe(single)
  })

  it('grows (rather than shifts) a merge range when a column is inserted strictly inside it', () => {
    const withMerge = createDocument([
      sheetFixture(
        [
          ['a', 'a', 'a'],
          ['b', 'c', 'd'],
        ],
        {
          grid: {
            rows: [
              ['a', 'a', 'a'],
              ['b', 'c', 'd'],
            ],
            colCount: 3,
            merges: [{ r0: 0, c0: 0, r1: 0, c1: 1 }],
            colWidthsPx: [],
            rowHeightsPx: [],
            formulas: [
              [undefined, undefined, undefined],
              [undefined, undefined, undefined],
            ],
          },
        },
      ),
    ])

    // Column 1 lands strictly inside the c0:0-c1:1 merge — it should grow to
    // c1:2, not shift right whole.
    const doc = insertColumnAt(withMerge, 0, 1)
    expect(doc.sheets[0].merges).toEqual([{ r0: 0, c0: 0, r1: 0, c1: 2 }])
  })

  it('shrinks (rather than leaves unchanged) a merge range when its own right column is deleted', () => {
    const withMerge = createDocument([
      sheetFixture(
        [
          ['a', 'a', 'b'],
          ['c', 'd', 'e'],
        ],
        {
          grid: {
            rows: [
              ['a', 'a', 'b'],
              ['c', 'd', 'e'],
            ],
            colCount: 3,
            merges: [{ r0: 0, c0: 0, r1: 0, c1: 1 }],
            colWidthsPx: [],
            rowHeightsPx: [],
            formulas: [
              [undefined, undefined, undefined],
              [undefined, undefined, undefined],
            ],
          },
        },
      ),
    ])

    // Column 1 is the merge's own RIGHT column (c0:0-c1:1). Deleting it must
    // shrink the merge to a single column (c1:0), not leave c1 at its old
    // numeric value of 1 — which would otherwise absorb the old column 2
    // (now slid left to index 1) even though it was never part of the merge.
    const doc = deleteColumnAt(withMerge, 0, 1)
    expect(doc.sheets[0].merges).toEqual([{ r0: 0, c0: 0, r1: 0, c1: 0 }])
    expect(doc.sheets[0].rows).toEqual([
      ['a', 'b'],
      ['c', 'e'],
    ])
  })
})

describe('pasteRange', () => {
  it('writes a rectangular block starting at the given cell', () => {
    const doc = pasteRange(basicDoc(), 0, 1, 0, [
      ['X', 'Y'],
      ['Z', 'W'],
    ])
    expect(doc.sheets[0].rows[1]).toEqual(['X', 'Y'])
    expect(doc.sheets[0].rows[2]).toEqual(['Z', 'W'])
  })

  it('grows the sheet when the pasted block runs past its current bounds', () => {
    const doc = pasteRange(basicDoc(), 0, 2, 2, [['new']])
    expect(doc.sheets[0].rows).toHaveLength(3)
    expect(doc.sheets[0].colCount).toBe(3)
    expect(doc.sheets[0].rows[2][2]).toBe('new')
  })

  it('detects a leading-= pasted cell as a formula', () => {
    const doc = pasteRange(basicDoc(), 0, 1, 1, [['=1+1']])
    expect(doc.sheets[0].formulas[1][1]).toBe('1+1')
    expect(doc.sheets[0].rows[1][1]).toBe('2')
  })

  it('keeps rows outside the pasted block\'s exact array reference when the paste stays within bounds (perf)', () => {
    const original = basicDoc()
    // Pastes into row 1 only; row 0 is untouched and colCount doesn't grow.
    const doc = pasteRange(original, 0, 1, 0, [['X', 'Y']])
    expect(doc.sheets[0].rows[0]).toBe(original.sheets[0].rows[0])
    expect(doc.sheets[0].rows[1]).not.toBe(original.sheets[0].rows[1])
  })
})

describe('sheet management', () => {
  it('adds a new sheet with a default, unique name', () => {
    const doc = addSheet(basicDoc())
    expect(doc.sheets).toHaveLength(2)
    expect(doc.sheets[1].name).toBe('Sheet2')
  })

  it('renames a sheet', () => {
    const doc = renameSheet(basicDoc(), 0, 'Renamed')
    expect(doc.sheets[0].name).toBe('Renamed')
  })

  it('refuses to rename to a name already used by another sheet', () => {
    const twoSheets = addSheet(basicDoc(), 'Other')
    const result = renameSheet(twoSheets, 1, 'Sheet1')
    expect(result.sheets[1].name).toBe('Other')
  })

  it('deletes a sheet', () => {
    const twoSheets = addSheet(basicDoc(), 'Other')
    const result = deleteSheet(twoSheets, 1)
    expect(result.sheets).toHaveLength(1)
  })

  it('refuses to delete the only remaining sheet', () => {
    const doc = basicDoc()
    expect(deleteSheet(doc, 0)).toBe(doc)
  })
})
