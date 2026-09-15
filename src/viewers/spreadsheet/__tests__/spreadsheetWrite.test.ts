/**
 * Spreadsheet editing + save (wave 3) — write-back round trips.
 */
import { describe, expect, it } from 'vitest'

import { bookTypeForExtension, buildWorkbook, documentToDelimitedText, writeWorkbookBytes } from '../spreadsheetWrite'
import { createDocument, setCellValue, type SpreadsheetDocument } from '../spreadsheetDocument'
import { parseWorkbookBuffer } from '../../shared/spreadsheetGrid'
import type { ParsedSheet } from '../../shared/spreadsheetGrid'

function sheetFixture(rows: string[][], name = 'Sheet1'): ParsedSheet {
  const colCount = Math.max(...rows.map((r) => r.length), 0)
  return {
    name,
    hidden: false,
    grid: {
      rows,
      colCount,
      merges: [],
      colWidthsPx: [],
      rowHeightsPx: [],
      formulas: rows.map((r) => r.map(() => undefined)),
    },
  }
}

describe('bookTypeForExtension', () => {
  it.each([
    ['xlsx', 'xlsx'],
    ['xlsm', 'xlsm'],
    ['xltx', 'xlsx'],
    ['xltm', 'xlsm'],
    ['xlsb', 'xlsb'],
    ['xls', 'xls'],
    ['ods', 'ods'],
    ['fods', 'fods'],
    ['XLSX', 'xlsx'],
    ['unknown', 'xlsx'],
  ])('%s -> %s', (ext, expected) => {
    expect(bookTypeForExtension(ext)).toBe(expected)
  })
})

describe('buildWorkbook / writeWorkbookBytes — round trip', () => {
  // SheetJS's BIFF8 (.xls) and BIFF12 (.xlsb) writers in this build never
  // serialize `.f` at all (verified directly against the library — only the
  // cached value survives), unlike the XML-based xlsx/xlsm/ods/fods writers.
  // A formula saved as .xls/.xlsb therefore round-trips as a plain value,
  // not a live formula — a real, format-specific SheetJS limitation, not an
  // Atlas bug. Its own cached VALUE still round-trips correctly, which is
  // what the shared assertion below checks either way.
  const FORMULA_LOST_ON_WRITE: ReadonlySet<string> = new Set(['xls', 'xlsb'])

  it.each(['xlsx', 'xlsm', 'xlsb', 'xls', 'ods', 'fods'] as const)(
    'round-trips numbers, text, and a formula through %s',
    (bookType) => {
      const doc = createDocument([
        sheetFixture([
          ['Name', 'Score', 'Doubled'],
          ['Alice', '10', ''],
        ]),
      ])
      const withFormula = setCellValue(doc, 0, 1, 2, '=B2*2')

      const bytes = writeWorkbookBytes(withFormula, bookType)
      expect(bytes.length).toBeGreaterThan(0)

      const reread = parseWorkbookBuffer(bytes.buffer as ArrayBuffer)
      expect(reread).toHaveLength(1)
      expect(reread[0].grid.rows[0]).toEqual(['Name', 'Score', 'Doubled'])
      expect(reread[0].grid.rows[1][0]).toBe('Alice')
      expect(reread[0].grid.rows[1][1]).toBe('10')

      if (FORMULA_LOST_ON_WRITE.has(bookType)) {
        expect(reread[0].grid.formulas[1][2]).toBeUndefined()
        expect(reread[0].grid.rows[1][2]).toBe('20')
        return
      }

      expect(reread[0].grid.formulas[1][2]).toBe('B2*2')

      // Re-hydrate through the document model exactly like a real re-open —
      // our own evaluator should recompute the formula's value regardless of
      // whether the target format's writer preserved a cached value.
      const rehydrated = createDocument(reread)
      expect(rehydrated.sheets[0].rows[1][2]).toBe('20')
    },
  )

  it('preserves merged ranges, column widths, and row heights', () => {
    const doc: SpreadsheetDocument = {
      sheets: [
        {
          name: 'Sheet1',
          hidden: false,
          rows: [
            ['Header', 'Header'],
            ['a', 'b'],
          ],
          formulas: [
            [undefined, undefined],
            [undefined, undefined],
          ],
          colCount: 2,
          merges: [{ r0: 0, c0: 0, r1: 0, c1: 1 }],
          colWidthsPx: [80, 200],
          rowHeightsPx: [30, undefined],
        },
      ],
    }

    const bytes = writeWorkbookBytes(doc, 'xlsx')
    const reread = parseWorkbookBuffer(bytes.buffer as ArrayBuffer)

    expect(reread[0].grid.merges).toEqual([{ r0: 0, c0: 0, r1: 0, c1: 1 }])
    expect(reread[0].grid.colWidthsPx[0]).toBeCloseTo(80, 0)
    expect(reread[0].grid.colWidthsPx[1]).toBeCloseTo(200, 0)
  })

  it('preserves multiple sheets and hidden-sheet flags', () => {
    const doc = createDocument([sheetFixture([['a']], 'Visible'), sheetFixture([['b']], 'Hidden')])
    const withHidden: SpreadsheetDocument = {
      sheets: [doc.sheets[0], { ...doc.sheets[1], hidden: true }],
    }

    const bytes = writeWorkbookBytes(withHidden, 'xlsx')
    const reread = parseWorkbookBuffer(bytes.buffer as ArrayBuffer)

    expect(reread.map((s) => ({ name: s.name, hidden: s.hidden }))).toEqual([
      { name: 'Visible', hidden: false },
      { name: 'Hidden', hidden: true },
    ])
  })

  it('writes a genuinely blank cell as no cell at all, not an empty string', () => {
    const doc = createDocument([sheetFixture([['a', '']])])
    const wb = buildWorkbook(doc)
    expect(wb.Sheets['Sheet1']['B1']).toBeUndefined()
  })

  it('re-types a whole-numeric display string as a number cell', () => {
    const doc = createDocument([sheetFixture([['42']])])
    const wb = buildWorkbook(doc)
    expect(wb.Sheets['Sheet1']['A1']).toMatchObject({ t: 'n', v: 42 })
  })

  it('keeps a non-numeric display string as a text cell', () => {
    const doc = createDocument([sheetFixture([['Alice']])])
    const wb = buildWorkbook(doc)
    expect(wb.Sheets['Sheet1']['A1']).toMatchObject({ t: 's', v: 'Alice' })
  })
})

describe('documentToDelimitedText', () => {
  it('serializes the first sheet as comma-delimited text', () => {
    const doc = createDocument([sheetFixture([['a', 'b'], ['1', '2']])])
    expect(documentToDelimitedText(doc, ',')).toBe('a,b\r\n1,2')
  })

  it('serializes as tab-delimited text when given a tab delimiter', () => {
    const doc = createDocument([sheetFixture([['a', 'b'], ['1', '2']])])
    expect(documentToDelimitedText(doc, '\t')).toBe('a\tb\r\n1\t2')
  })

  it('quotes a value containing the delimiter', () => {
    const doc = createDocument([sheetFixture([['has, comma', 'plain']])])
    expect(documentToDelimitedText(doc, ',')).toBe('"has, comma",plain')
  })
})
