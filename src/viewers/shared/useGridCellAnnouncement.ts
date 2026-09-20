/**
 * A11Y pass 3 — glide-data-grid (SpreadsheetViewer/CsvViewer's grid) renders
 * entirely to `<canvas>`: there is no DOM cell for a screen reader to land
 * on, and this library version exposes no accessibility hooks at all (no
 * `role`, no per-cell text, nothing — confirmed against the installed
 * `@glideapps/glide-data-grid` build). Arrow-key/Tab navigation inside the
 * grid already works natively (glide-data-grid owns its own keyboard
 * handling and reports every resulting selection through the controlled
 * `gridSelection`/`onGridSelectionChange` props both viewers already wire
 * through `useGridFind`), so this doesn't reimplement navigation — it only
 * gives a screen reader user the one thing the canvas can't: knowing which
 * cell is now selected and what's in it, mirroring what a native spreadsheet
 * announces on every arrow-key move.
 *
 * Deliberately NOT gated to only fire while `useGridFind`'s find bar is
 * open: `gridSelection` is the SAME controlled value the grid reports every
 * ordinary arrow-key/click move through too, so this covers both.
 */
import { useMemo } from 'react'
import type { GridSelection } from '@glideapps/glide-data-grid'

import { useTranslate } from '../../i18n'

/** 0-based column index -> spreadsheet-style letters (0 -> A, 25 -> Z, 26 -> AA, ...). */
function columnLetters(index: number): string {
  let n = index + 1
  let letters = ''
  while (n > 0) {
    const remainder = (n - 1) % 26
    letters = String.fromCharCode(65 + remainder) + letters
    n = Math.floor((n - 1) / 26)
  }
  return letters
}

export function cellReference(col: number, row: number): string {
  return `${columnLetters(col)}${row + 1}`
}

/**
 * Returns the text for a visually-hidden `aria-live="polite"` region: the
 * currently-selected cell's reference plus its value (or a "blank" notice),
 * empty when nothing is selected. Callers render this in a `role="status"`
 * element next to the grid — see SpreadsheetViewer.tsx/CsvViewer.tsx.
 */
export function useGridCellAnnouncement(
  rows: ReadonlyArray<ReadonlyArray<string>>,
  gridSelection: GridSelection | undefined,
): string {
  const t = useTranslate()
  const cell = gridSelection?.current?.cell

  return useMemo(() => {
    if (!cell) return ''
    const [col, row] = cell
    const ref = cellReference(col, row)
    const value = rows[row]?.[col]
    return value && value.length > 0
      ? t('grid.cellAnnouncement', { ref, value })
      : t('grid.cellAnnouncementEmpty', { ref })
    // `cell` is a fresh tuple/array identity from glide-data-grid on every
    // selection event, so depend on its two primitive members instead —
    // otherwise this would recompute (and the live region would re-announce)
    // on every render even when the selected cell hasn't actually changed.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [cell?.[0], cell?.[1], rows, t])
}
