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
