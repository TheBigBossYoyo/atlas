/**
 * SHEETFN-3 — a formula result wears the cell's own number format.
 *
 * Atlas displayed a formula's result exactly as the evaluator rendered it,
 * ignoring the cell's number format: `=A1*1.2` in a currency column showed
 * `15.6` where Excel shows `$15.60`, and a date formula in a `dd/mm/yyyy`
 * column showed an ISO date instead of the workbook's own convention. A stored
 * value and a computed one with the same format rendered differently, in the
 * same column.
 *
 * The formatting is SheetJS's `SSF` — the same engine that already produces the
 * display text for every value READ from a workbook — so the two cannot
 * disagree. A second implementation of OOXML format codes would.
 */
import { describe, expect, it } from 'vitest'

import { createDocument, recalculateSheet, setCellValue, setRangeFormat } from '../spreadsheetDocument'
import { applyNumberFormat } from '../formatNumber'
import { serialFromIsoDate } from '../excelDate'
import type { ParsedSheet } from '../../shared/spreadsheetGrid'

function sheetFixture(rows: string[][], formulas: (string | undefined)[][] = []): ParsedSheet {
  return {
    name: 'Sheet1',
    hidden: false,
    grid: {
      rows,
      colCount: Math.max(...rows.map(r => r.length), 0),
      merges: [],
      colWidthsPx: [],
      rowHeightsPx: [],
      formulas: rows.map((row, r) => row.map((_, c) => formulas[r]?.[c])),
    },
  }
}

describe('applyNumberFormat (SHEETFN-3)', () => {
  it('formats with a real OOXML format code', () => {
    expect(applyNumberFormat(1234.5, '#,##0.00')).toBe('1,234.50')
    expect(applyNumberFormat(0.125, '0%')).toBe('13%')
    expect(applyNumberFormat(12.5, '0.000')).toBe('12.500')
  })

  it('formats a date serial with the workbook\'s own date convention', () => {
    const serial = serialFromIsoDate('2026-10-02')!
    expect(applyNumberFormat(serial, 'dd/mm/yyyy')).toBe('02/10/2026')
    expect(applyNumberFormat(serial, 'yyyy-mm-dd')).toBe('2026-10-02')
  })

  it('returns null for General and for no format, so the caller keeps its own rendering', () => {
    expect(applyNumberFormat(1.5, undefined)).toBeNull()
    expect(applyNumberFormat(1.5, '')).toBeNull()
    expect(applyNumberFormat(1.5, 'General')).toBeNull()
    expect(applyNumberFormat(1.5, 'general')).toBeNull()
  })

  it('returns null rather than throwing on a format code it cannot apply', () => {
    // A malformed code must leave the number readable, not break recalculation
    // or replace the cell with an error.
    expect(applyNumberFormat(1.5, '[[[not a format')).toBeNull()
  })
})

describe('recalculation applies the cell format (SHEETFN-3)', () => {
  it('renders a computed number in the cell\'s own format', () => {
    let doc = createDocument([sheetFixture([['10', '']], [[undefined, 'A1*1.2']])])
    doc = setRangeFormat(doc, 0, { row0: 0, col0: 1, row1: 0, col1: 1 }, { numberFormat: '#,##0.00' })
    const sheet = recalculateSheet(doc.sheets[0])
    expect(sheet.rows[0][1]).toBe('12.00')
  })

  it('renders a computed date in the cell\'s own date format', () => {
    let doc = createDocument([sheetFixture([['']], [['DATE(2026,10,2)']])])
    doc = setRangeFormat(doc, 0, { row0: 0, col0: 0, row1: 0, col1: 0 }, { numberFormat: 'dd/mm/yyyy' })
    const sheet = recalculateSheet(doc.sheets[0])
    expect(sheet.rows[0][0]).toBe('02/10/2026')
  })

  it('falls back to the evaluator\'s own rendering when the cell has no format', () => {
    // Which is what keeps a date readable as ISO rather than as its serial —
    // the behaviour SHEETFN-2 relies on.
    const doc = createDocument([sheetFixture([['']], [['DATE(2026,10,2)']])])
    const sheet = recalculateSheet(doc.sheets[0])
    expect(sheet.rows[0][0]).toBe('2026-10-02')
  })

  it('leaves a text result alone — a number format has nothing to say about it', () => {
    let doc = createDocument([sheetFixture([['']], [['"hello"']])])
    doc = setRangeFormat(doc, 0, { row0: 0, col0: 0, row1: 0, col1: 0 }, { numberFormat: '#,##0.00' })
    const sheet = recalculateSheet(doc.sheets[0])
    expect(sheet.rows[0][0]).toBe('hello')
  })

  it('leaves an error result alone', () => {
    let doc = createDocument([sheetFixture([['']], [['1/0']])])
    doc = setRangeFormat(doc, 0, { row0: 0, col0: 0, row1: 0, col1: 0 }, { numberFormat: '#,##0.00' })
    const sheet = recalculateSheet(doc.sheets[0])
    expect(sheet.rows[0][0]).toBe('#DIV/0!')
  })

  it('applies a percent format, which changes the number the user sees', () => {
    // The clearest case: 0.125 displayed as 13% is the format doing real work,
    // not cosmetic rounding.
    let doc = createDocument([sheetFixture([['0.5', '']], [[undefined, 'A1/4']])])
    doc = setRangeFormat(doc, 0, { row0: 0, col0: 1, row1: 0, col1: 1 }, { numberFormat: '0%' })
    const sheet = recalculateSheet(doc.sheets[0])
    expect(sheet.rows[0][1]).toBe('13%')
  })

  it('keeps following the format after the formula is edited', () => {
    let doc = createDocument([sheetFixture([['10', '']], [[undefined, 'A1*2']])])
    doc = setRangeFormat(doc, 0, { row0: 0, col0: 1, row1: 0, col1: 1 }, { numberFormat: '0.00' })
    doc = setCellValue(doc, 0, 0, 0, '25')
    expect(doc.sheets[0].rows[0][1]).toBe('50.00')
  })
})
