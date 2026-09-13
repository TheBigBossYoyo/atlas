/**
 * Shared glide-data-grid rendering logic for SpreadsheetViewer and CsvViewer
 * (T5/DAT-12 — the two viewers were ~90% duplicated cell/column-builder code).
 *
 * Consolidated here now that T1/T3/T4's fixes (formatted values, merged-cell
 * backfill, column widths) live in one place: both viewers hand this hook
 * already-formatted `rows` and get back the same `getCellContent`/`columns`/
 * `onColumnResize`/`onItemHovered` glide-data-grid needs.
 */
import { useCallback, useMemo, useState } from 'react'
import type { GridCell, GridColumn, GridMouseEventArgs } from '@glideapps/glide-data-grid'

import { useGridTheme, type CustomGridTheme } from './useGridTheme'

const DEFAULT_COLUMN_WIDTH_PX = 120
/** Sample at most this many rows per column to guess whether it's numeric. */
const NUMERIC_SAMPLE_SIZE = 50
/** A column is treated as numeric (right-aligned) once this share of its sampled values parse as numbers. */
const NUMERIC_SAMPLE_THRESHOLD = 0.8

export type UseSpreadsheetGridOptions = {
  readonly rows: ReadonlyArray<ReadonlyArray<string>>
  readonly colCount: number
  /** Initial pixel width per column index (e.g. from a workbook's `!cols`), overridden once the user drags a column. */
  readonly colWidthsPx?: ReadonlyArray<number | undefined>
  /**
   * Changing this value (e.g. the active sheet's name) clears any
   * user-resized column widths so a newly-selected sheet starts from its own
   * `colWidthsPx` again instead of the previous sheet's drag state.
   */
  readonly resetKey?: string
}

export type UseSpreadsheetGridResult = {
  readonly columns: GridColumn[]
  readonly getCellContent: (item: readonly [number, number]) => GridCell
  readonly onColumnResize: (column: GridColumn, newSize: number) => void
  readonly onItemHovered: (args: GridMouseEventArgs) => void
  readonly hoveredRow: number | undefined
  readonly theme: CustomGridTheme
}

function columnTitle(index: number): string {
  let title = ''
  let n = index
  while (n >= 0) {
    title = String.fromCharCode(65 + (n % 26)) + title
    n = Math.floor(n / 26) - 1
  }
  return title
}

function isColumnNumeric(rows: ReadonlyArray<ReadonlyArray<string>>, col: number): boolean {
  let numericCount = 0
  let totalSampled = 0
  for (const row of rows) {
    const value = row[col]
    if (value !== undefined && value !== null && value.trim() !== '') {
      totalSampled++
      if (!isNaN(Number(value))) {
        numericCount++
      }
    }
    if (totalSampled >= NUMERIC_SAMPLE_SIZE) break
  }
  return totalSampled > 0 && numericCount / totalSampled >= NUMERIC_SAMPLE_THRESHOLD
}

export function useSpreadsheetGrid({
  rows,
  colCount,
  colWidthsPx,
  resetKey,
}: UseSpreadsheetGridOptions): UseSpreadsheetGridResult {
  const theme = useGridTheme()
  const [hoveredRow, setHoveredRow] = useState<number | undefined>()
  const [userColWidths, setUserColWidths] = useState<Record<string, number>>({})

  // A newly-selected sheet (or a freshly-loaded file) should start from its
  // own column widths, not whatever the previously-active sheet's columns
  // were dragged to. Pure derivation of `resetKey`, so this uses the
  // render-time "adjust state" idiom (see ViewerContext.tsx) rather than an effect.
  const [lastResetKey, setLastResetKey] = useState(resetKey)
  if (lastResetKey !== resetKey) {
    setLastResetKey(resetKey)
    setUserColWidths({})
  }

  const getCellContent = useCallback(
    ([col, row]: readonly [number, number]): GridCell => {
      const cellValue = rows[row]?.[col] ?? ''
      const isOdd = row % 2 !== 0
      const isHovered = row === hoveredRow
      let bgCell = isOdd ? theme.bgRowOdd : theme.bgCell
      if (isHovered) bgCell = theme.bgRowHover

      return {
        kind: 'text' as const,
        data: cellValue,
        displayData: cellValue,
        // Spreadsheet editing is out of scope (DAT-06) — allowOverlay:true
        // opened an edit box whose typed input was silently discarded, since
        // no onCellEdited was ever wired up. Keep the grid read-only and honest.
        allowOverlay: false,
        themeOverride: { bgCell },
      } as GridCell
    },
    [rows, theme, hoveredRow],
  )

  const columns = useMemo<GridColumn[]>(() => {
    return Array.from({ length: colCount }, (_, i) => {
      const userWidth = userColWidths[i]
      const fileWidth = colWidthsPx?.[i]
      const width = userWidth ?? fileWidth ?? DEFAULT_COLUMN_WIDTH_PX
      return {
        title: columnTitle(i),
        id: String(i),
        width,
        grow: userWidth === undefined && fileWidth === undefined ? 1 : undefined,
        contentAlign: isColumnNumeric(rows, i) ? ('right' as const) : undefined,
      }
    })
  }, [colCount, colWidthsPx, rows, userColWidths])

  const onColumnResize = useCallback((column: GridColumn, newSize: number) => {
    setUserColWidths((prev) => ({ ...prev, [column.id ?? '']: newSize }))
  }, [])

  const onItemHovered = useCallback((args: GridMouseEventArgs) => {
    setHoveredRow(args.location[1] >= 0 ? args.location[1] : undefined)
  }, [])

  return { columns, getCellContent, onColumnResize, onItemHovered, hoveredRow, theme }
}
