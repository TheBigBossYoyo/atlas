/**
 * USR-17 — the Excel-like blank margin shared by SpreadsheetViewer and
 * CsvViewer: extra empty rows/columns drawn past the data so a value can be
 * typed into an empty cell, plus the grid-row -> sheet-row mapping that makes
 * those margin rows addressable. The margin is off while the row search
 * filter is active (filtered rows are not contiguous, so "the row after the
 * data" has no meaning there).
 */
import { useCallback } from 'react'

import type { EditableSheet } from '../spreadsheet/spreadsheetDocument'

const BLANK_MARGIN_ROWS = 50
const BLANK_MARGIN_COLS = 10
const MIN_GRID_ROWS = 100
const MIN_GRID_COLS = 26

export type UseGridBlankMarginOptions = {
  /** Sheet row for each data row the grid body shows, in grid order. */
  readonly bodyRowIndices: ReadonlyArray<number>
  readonly colCount: number
  readonly isFiltering: boolean
  /** Sheet rows rendered outside the grid body above it (frozen rows). */
  readonly rowOffset?: number
}

export type GridBlankMargin = {
  readonly gridRowCount: number
  readonly gridColCount: number
  readonly sheetRowForGridRow: (gridRow: number) => number | undefined
}

export function useGridBlankMargin({
  bodyRowIndices,
  colCount,
  isFiltering,
  rowOffset = 0,
}: UseGridBlankMarginOptions): GridBlankMargin {
  const sheetRowForGridRow = useCallback(
    (gridRow: number): number | undefined =>
      bodyRowIndices[gridRow] ?? (isFiltering || gridRow < 0 ? undefined : rowOffset + gridRow),
    [bodyRowIndices, isFiltering, rowOffset],
  )
  return {
    gridRowCount: isFiltering
      ? bodyRowIndices.length
      : Math.max(bodyRowIndices.length + BLANK_MARGIN_ROWS, MIN_GRID_ROWS - rowOffset),
    gridColCount: Math.max(colCount + BLANK_MARGIN_COLS, MIN_GRID_COLS),
    sheetRowForGridRow,
  }
}

type CellEditor = {
  readonly setCellValue: (sheetIndex: number, row: number, col: number, rawInput: string) => void
  readonly pasteRange: (sheetIndex: number, row: number, col: number, values: ReadonlyArray<ReadonlyArray<string>>) => void
}

/** Commits a grid edit: in-bounds cells are set, a non-empty value typed into the margin grows the sheet to reach it. */
export function commitGridEdit(
  editor: CellEditor,
  sheetIndex: number,
  sheet: EditableSheet,
  sheetRow: number,
  col: number,
  rawText: string,
): void {
  if (sheetRow < sheet.rows.length && col < sheet.colCount) {
    editor.setCellValue(sheetIndex, sheetRow, col, rawText)
  } else if (rawText !== '') {
    editor.pasteRange(sheetIndex, sheetRow, col, [[rawText]])
  }
}
