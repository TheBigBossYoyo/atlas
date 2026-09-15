/**
 * Shared glide-data-grid rendering logic for SpreadsheetViewer and CsvViewer
 * (T5/DAT-12 — the two viewers were ~90% duplicated cell/column-builder code).
 *
 * Consolidated here now that T1/T3/T4's fixes (formatted values, merged-cell
 * backfill, column widths) live in one place: both viewers hand this hook
 * already-formatted `rows` and get back the same `getCellContent`/`columns`/
 * `onColumnResize`/`onItemHovered` glide-data-grid needs.
 *
 * Editing (wave 3, DAT-06 closed for real): passing `onCellEdited` makes
 * every cell editable (`allowOverlay:true`) and returns a ready-to-spread
 * `onCellEdited` handler that translates glide-data-grid's `Item`/
 * `EditableGridCell` shape into the plain `(row, col, rawText)` the caller's
 * callback wants — the caller (SpreadsheetViewer/CsvViewer, via
 * `useSpreadsheetEditor`) owns turning that into an actual document edit.
 * Omitting `onCellEdited` keeps the grid honestly read-only, exactly like
 * before wave 3.
 *
 * `formulas` (wave 3 follow-up — confirmed data-loss bug, not a style nit):
 * a `TextCell`'s edit overlay seeds its input from `data`, NOT `displayData`
 * (glide-data-grid's own `text-cell.js`: `provideEditor` reads `value.data`).
 * Before this option existed, `getCellContent` set both fields to the same
 * already-computed display text, so double-clicking an EXISTING formula
 * cell (e.g. one showing "15" for `=SUM(A1:A5)`) opened an editor pre-filled
 * with "15", not "=SUM(A1:A5)" — and committing that overlay with NO changes
 * at all (open, press Enter) fed "15" back through `onCellEdited`, which
 * `spreadsheetDocument.ts`'s `setCellValue` — correctly, given what it was
 * told — stores as a plain value, silently deleting the formula. `formulas`
 * lets `getCellContent` seed `data` with the literal `=<formula>` text
 * instead whenever the cell has one, so re-opening and committing a formula
 * cell unchanged round-trips it as the same formula. `displayData` (what's
 * actually drawn, and what Ctrl+C copies — see glide-data-grid's
 * `copy-paste.js`, which reads `displayData` for `GridCellKind.Text`) is
 * unaffected either way, so this changes nothing about how a formula cell
 * looks or copies, only what's pre-filled when you start editing it.
 */
import { useCallback, useMemo, useState } from 'react'
import type { EditableGridCell, GridCell, GridColumn, Item, GridMouseEventArgs } from '@glideapps/glide-data-grid'

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
  /**
   * When provided, cells render with `allowOverlay:true` (editable) and a
   * committed edit calls this with the raw text the user typed — including a
   * leading `=` for a formula, unmodified, so the caller's document layer
   * (not this rendering hook) decides what that means. Omit to keep the
   * grid read-only.
   */
  readonly onCellEdited?: (row: number, col: number, rawText: string) => void
  /**
   * Per-cell formula text (no leading `=`), the same shape as `rows` —
   * `EditableSheet.formulas`, threaded through by the caller. Optional (CSV
   * sheets before any formula has ever been typed into them have no need to
   * pass this), but see the module header: omitting it for a sheet that DOES
   * have formula cells reintroduces the formula-deleting edit bug this option
   * exists to fix.
   */
  readonly formulas?: ReadonlyArray<ReadonlyArray<string | undefined>>
  /** Grid-space cells drawn as a header (bold on the header background) — Excel table header rows (USR-17). */
  readonly isHeaderCell?: (row: number, col: number) => boolean
}

export type UseSpreadsheetGridResult = {
  readonly columns: GridColumn[]
  readonly getCellContent: (item: readonly [number, number]) => GridCell
  readonly onColumnResize: (column: GridColumn, newSize: number) => void
  readonly onItemHovered: (args: GridMouseEventArgs) => void
  readonly hoveredRow: number | undefined
  readonly theme: CustomGridTheme
  /** Ready to pass straight to `DataEditor`'s `onCellEdited` prop; `undefined` when the caller didn't ask for editing. */
  readonly onCellEdited: ((cell: Item, newValue: EditableGridCell) => void) | undefined
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
  onCellEdited: onCellEditedOption,
  formulas,
  isHeaderCell,
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
      const formula = formulas?.[row]?.[col]
      // See module header: the edit overlay seeds itself from `data`, so a
      // formula cell must offer its literal `=<formula>` text there, not the
      // already-computed `cellValue` — only `displayData` (what's drawn/
      // copied) stays the computed value.
      const editableValue = formula !== undefined ? `=${formula}` : cellValue
      const isOdd = row % 2 !== 0
      const isHovered = row === hoveredRow
      let bgCell = isOdd ? theme.bgRowOdd : theme.bgCell
      if (isHovered) bgCell = theme.bgRowHover
      const isHeader = isHeaderCell?.(row, col) ?? false

      return {
        kind: 'text' as const,
        data: editableValue,
        displayData: cellValue,
        // DAT-06, closed for real (wave 3): an overlay is only opened when
        // the caller actually wired up `onCellEdited` below — otherwise the
        // grid stays honestly read-only exactly as before.
        allowOverlay: onCellEditedOption !== undefined,
        themeOverride: isHeader
          ? { bgCell: theme.bgHeader, baseFontStyle: `600 ${theme.baseFontStyle}` }
          : { bgCell },
      } as GridCell
    },
    [rows, formulas, theme, hoveredRow, onCellEditedOption, isHeaderCell],
  )

  const onCellEdited = useMemo(() => {
    if (!onCellEditedOption) return undefined
    return (cell: Item, newValue: EditableGridCell): void => {
      // Every cell this hook renders is `kind: 'text'` (see getCellContent
      // above), so the overlay glide-data-grid opens for it is always the
      // plain text editor — `newValue` is only ever a `TextCell` in
      // practice, but the guard keeps this correct if that ever changes.
      if (newValue.kind !== 'text') return
      const [col, row] = cell
      onCellEditedOption(row, col, newValue.data)
    }
  }, [onCellEditedOption])

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

  return { columns, getCellContent, onColumnResize, onItemHovered, hoveredRow, theme, onCellEdited }
}
