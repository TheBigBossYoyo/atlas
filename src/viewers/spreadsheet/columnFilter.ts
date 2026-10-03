/**
 * SHEET-FILTER-1 — per-column filtering, Excel's AutoFilter.
 *
 * A filter is a set of ALLOWED values for one column; a row is shown when
 * every active column's filter admits its value. This is a VIEW concern, not a
 * document one: nothing here edits the workbook, and a filtered-out row is
 * still saved. The viewer maps grid rows back to sheet rows through
 * `filteredRowIndices`, which already exists for the free-text row search, so
 * an edit made while a filter is active still lands on the row the user is
 * looking at.
 *
 * Values are grouped by `cellOrder`'s fold, so the checklist offers "Paper"
 * once rather than once per spelling, and offers exactly the groups that a
 * sort on the same column would keep together.
 */
import { compareForSort, foldText } from './cellOrder'

/**
 * The checklist key for an empty cell — Excel's "(Blanks)" entry.
 *
 * The empty string cannot collide with a real value's key, because a real
 * value's key is non-empty by construction (a cell that trims to nothing IS a
 * blank).
 */
export const BLANK_FILTER_KEY = ''

/** Column index → the values that column admits. A column absent from the map is unfiltered. */
export type ColumnFilters = ReadonlyMap<number, ReadonlySet<string>>

export type FilterOption = {
  /** The key stored in a filter set. `BLANK_FILTER_KEY` for empty cells. */
  readonly key: string
  /** The first spelling encountered, for display. Empty for the blanks entry. */
  readonly label: string
  /** How many rows hold this value — Excel shows no count, but it tells you what you are about to hide. */
  readonly count: number
  readonly blank: boolean
}

/** The filter key for one cell. Trimmed, folded, and blank-aware. */
export function filterKeyFor(cell: string): string {
  const trimmed = cell.trim()
  return trimmed === '' ? BLANK_FILTER_KEY : foldText(trimmed)
}

/**
 * The distinct values of one column, in the order a sort would put them.
 *
 * Rows are read to the end of the sheet, including rows currently hidden by
 * ANOTHER column's filter. That is Excel's behaviour and it is the useful one:
 * if choosing a value in column B could only ever narrow what column A already
 * showed, a filter could never be widened without clearing the other one
 * first.
 */
export function distinctColumnValues(
  rows: ReadonlyArray<ReadonlyArray<string>>,
  col: number,
  headerRows = 0,
): FilterOption[] {
  const byKey = new Map<string, { label: string; count: number }>()
  for (let r = headerRows; r < rows.length; r++) {
    const cell = rows[r]?.[col] ?? ''
    const key = filterKeyFor(cell)
    const existing = byKey.get(key)
    if (existing) existing.count++
    else byKey.set(key, { label: key === BLANK_FILTER_KEY ? '' : cell.trim(), count: 1 })
  }
  return [...byKey.entries()]
    .map(([key, v]) => ({ key, label: v.label, count: v.count, blank: key === BLANK_FILTER_KEY }))
    .sort((a, b) => compareForSort(a.blank ? '' : a.label, b.blank ? '' : b.label, 'asc'))
}

/** True when every active filter admits this row. */
export function rowMatchesFilters(row: ReadonlyArray<string>, filters: ColumnFilters): boolean {
  for (const [col, allowed] of filters) {
    if (!allowed.has(filterKeyFor(row[col] ?? ''))) return false
  }
  return true
}

/**
 * Narrows a filter set to a column's actual values, dropping it entirely when
 * it admits all of them.
 *
 * Two reasons this is not just "store what was ticked". An all-ticked filter is
 * the same as no filter, and storing it would light up the "filtered" indicator
 * and disable frozen rows for nothing. And a filter whose values no longer
 * exist — the rows were deleted, or the column was re-typed — would hide every
 * row with no way to tell from the UI why.
 */
export function normalizeFilter(
  selected: ReadonlySet<string>,
  options: ReadonlyArray<FilterOption>,
): ReadonlySet<string> | null {
  const present = new Set<string>()
  for (const option of options) if (selected.has(option.key)) present.add(option.key)
  if (present.size === options.length) return null
  return present
}

/** Sets or clears one column's filter, returning a new map. `null` clears. */
export function withColumnFilter(
  filters: ColumnFilters,
  col: number,
  allowed: ReadonlySet<string> | null,
): ColumnFilters {
  const next = new Map(filters)
  if (allowed === null) next.delete(col)
  else next.set(col, allowed)
  return next
}

/**
 * Re-aims the filters after rows or columns move.
 *
 * Column filters are keyed by column INDEX, so inserting or deleting a column
 * shifts which column a filter applies to. Leaving them put would silently
 * filter the wrong column — the data equivalent of the row-mapping bug the
 * free-text search once had. A filter on the deleted column is dropped.
 */
export function shiftFiltersForColumnChange(
  filters: ColumnFilters,
  at: number,
  change: 'insert' | 'delete',
): ColumnFilters {
  const next = new Map<number, ReadonlySet<string>>()
  for (const [col, allowed] of filters) {
    if (change === 'insert') next.set(col >= at ? col + 1 : col, allowed)
    else if (col !== at) next.set(col > at ? col - 1 : col, allowed)
  }
  return next
}
