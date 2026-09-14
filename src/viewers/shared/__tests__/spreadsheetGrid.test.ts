/**
 * T1/T3/T4 (DAT-04/08/09/10/11) — cell formatting, merged-cell backfill,
 * column-width/row-height carry-through, and hidden-sheet detection.
 */
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { attachFrozenPanes, parseWorkbookBuffer, sheetToGrid } from '../spreadsheetGrid'

describe('sheetToGrid', () => {
  it('renders a formatted date via cell.w instead of a raw serial number (DAT-04)', () => {
    const ws: XLSX.WorkSheet = {
      '!ref': 'A1:A1',
      A1: { t: 'd', v: new Date(Date.UTC(2024, 2, 15)), w: '3/15/24' } as XLSX.CellObject,
    }

    const grid = sheetToGrid(ws)

    expect(grid.rows).toEqual([['3/15/24']])
  })

  it('renders a formula-error cell as its error text instead of blank (DAT-04)', () => {
    const ws: XLSX.WorkSheet = {
      '!ref': 'A1:B1',
      A1: { t: 's', v: 'ok', w: 'ok' } as XLSX.CellObject,
      B1: { t: 'e', v: 7, w: '#DIV/0!' } as XLSX.CellObject,
    }

    const grid = sheetToGrid(ws)

    expect(grid.rows).toEqual([['ok', '#DIV/0!']])
  })

  it('falls back to a formula-error code table when a cell has no cached .w', () => {
    const ws: XLSX.WorkSheet = {
      '!ref': 'A1:A1',
      A1: { t: 'e', v: 0x07 } as XLSX.CellObject,
    }

    const grid = sheetToGrid(ws)

    expect(grid.rows).toEqual([['#DIV/0!']])
  })

  it('formats an unformatted date value using the OS locale, not a hardcoded one (DEFER-6)', () => {
    const date = new Date(Date.UTC(2024, 0, 5))
    const ws: XLSX.WorkSheet = {
      '!ref': 'A1:A1',
      A1: { t: 'd', v: date } as XLSX.CellObject,
    }

    const grid = sheetToGrid(ws)

    expect(grid.rows).toEqual([[new Intl.DateTimeFormat().format(date)]])
  })

  it('backfills a merged range so its value is not blank outside the top-left cell (DAT-09)', () => {
    const ws: XLSX.WorkSheet = {
      '!ref': 'A1:C2',
      A1: { t: 's', v: 'Header', w: 'Header' } as XLSX.CellObject,
      '!merges': [{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }],
    }

    const grid = sheetToGrid(ws)

    expect(grid.rows[0]).toEqual(['Header', 'Header', 'Header'])
    expect(grid.merges).toEqual([{ r0: 0, c0: 0, r1: 0, c1: 2 }])
  })

  it('carries column widths and row heights from !cols/!rows (DAT-10)', () => {
    const ws: XLSX.WorkSheet = {
      '!ref': 'A1:B2',
      '!cols': [{ wpx: 80 }, { wpx: 200 }],
      '!rows': [{ hpx: 40 }, { hpx: 20 }],
    }

    const grid = sheetToGrid(ws)

    expect(grid.colWidthsPx).toEqual([80, 200])
    expect(grid.rowHeightsPx).toEqual([40, 20])
  })

  it('returns an empty grid for a sheet with no !ref', () => {
    const grid = sheetToGrid({})
    expect(grid).toEqual({
      rows: [],
      colCount: 0,
      merges: [],
      colWidthsPx: [],
      rowHeightsPx: [],
      formulas: [],
    })
  })

  it('carries each cellformula text (no leading "=") alongside its formatted display text (wave 3 editing)', () => {
    const ws: XLSX.WorkSheet = {
      '!ref': 'A1:B1',
      A1: { t: 'n', v: 3, w: '3' } as XLSX.CellObject,
      B1: { t: 'n', v: 5, w: '5', f: 'A1+2' } as XLSX.CellObject,
    }

    const grid = sheetToGrid(ws)

    expect(grid.formulas).toEqual([[undefined, 'A1+2']])
  })
})

describe('attachFrozenPanes', () => {
  it('merges frozen-pane info onto the sheet with a matching name', () => {
    const sheets = [
      { name: 'Sheet1', hidden: false, grid: sheetToGrid({}) },
      { name: 'Sheet2', hidden: false, grid: sheetToGrid({}) },
    ]

    const result = attachFrozenPanes(sheets, { Sheet2: { cols: 1, rows: 2 } })

    expect(result[0].freeze).toBeUndefined()
    expect(result[1].freeze).toEqual({ cols: 1, rows: 2 })
  })

  it('leaves sheets unchanged when the pane map is empty', () => {
    const sheets = [{ name: 'Sheet1', hidden: false, grid: sheetToGrid({}) }]
    expect(attachFrozenPanes(sheets, {})).toEqual(sheets)
  })
})

describe('parseWorkbookBuffer', () => {
  it('flags a hidden sheet (DAT-11) and keeps a visible one unflagged', () => {
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['visible']]), 'Visible')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['hidden']]), 'Hidden')
    wb.Workbook = { Sheets: [{ Hidden: 0 }, { Hidden: 1 }] }

    const buffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer
    const sheets = parseWorkbookBuffer(buffer)

    expect(sheets.map((s) => ({ name: s.name, hidden: s.hidden }))).toEqual([
      { name: 'Visible', hidden: false },
      { name: 'Hidden', hidden: true },
    ])
  })

  it('parses real formatted-number and date cells end to end through XLSX.write/read', () => {
    const ws = XLSX.utils.aoa_to_sheet([['Amount'], [1234.5]])
    ws['A2'].z = '$#,##0.00'
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')

    const buffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer
    const sheets = parseWorkbookBuffer(buffer)

    expect(sheets[0].grid.rows[1][0]).toBe('$1,234.50')
  })
})
