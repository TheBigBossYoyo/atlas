/**
 * In-grid find/jump for SpreadsheetViewer and CsvViewer.
 *
 * Distinct from the viewers' existing "Search rows..." box (which *filters*
 * out non-matching rows): this scans every visible cell for a query and lets
 * the user jump to and select each match in turn, mirroring markdown's
 * Ctrl+F `SearchOverlay` (match count, next/prev) but against grid cells
 * instead of DOM text nodes. Exposed to the shell dispatcher as `openFind()`
 * via the ViewerContext capability contract (see `useRegisterViewerFind`).
 */
import { useCallback, useDeferredValue, useMemo, useState } from 'react'
import { CompactSelection, type GridSelection } from '@glideapps/glide-data-grid'

export type GridFindMatch = {
  readonly row: number
  readonly col: number
}

export type UseGridFindResult = {
  readonly isOpen: boolean
  readonly query: string
  readonly setQuery: (query: string) => void
  readonly matchCount: number
  readonly currentMatch: number
  readonly goToMatch: (direction: 'next' | 'prev') => void
  readonly open: () => void
  readonly close: () => void
  /** The currently-selected match, as a controlled glide-data-grid selection. */
  readonly gridSelection: GridSelection | undefined
  readonly onGridSelectionChange: (selection: GridSelection) => void
}

const EMPTY_SELECTION_SETS = { columns: CompactSelection.empty(), rows: CompactSelection.empty() } as const

function selectionForMatch(match: GridFindMatch): GridSelection {
  return {
    current: {
      cell: [match.col, match.row],
      range: { x: match.col, y: match.row, width: 1, height: 1 },
      rangeStack: [],
    },
    ...EMPTY_SELECTION_SETS,
  }
}

export function useGridFind(rows: ReadonlyArray<ReadonlyArray<string>>): UseGridFindResult {
  const [isOpen, setIsOpen] = useState(false)
  const [query, setQuery] = useState('')
  const [currentMatch, setCurrentMatch] = useState(0)
  const [gridSelection, setGridSelection] = useState<GridSelection | undefined>(undefined)
  const deferredQuery = useDeferredValue(query)

  const matches = useMemo<GridFindMatch[]>(() => {
    const q = deferredQuery.trim().toLowerCase()
    if (!q) return []
    const found: GridFindMatch[] = []
    for (let row = 0; row < rows.length; row++) {
      const cells = rows[row]
      for (let col = 0; col < cells.length; col++) {
        if (cells[col]?.toLowerCase().includes(q)) {
          found.push({ row, col })
        }
      }
    }
    return found
  }, [rows, deferredQuery])

  const selectMatch = useCallback((index: number, list: readonly GridFindMatch[]) => {
    setCurrentMatch(index)
    const match = list[index]
    setGridSelection(match ? selectionForMatch(match) : undefined)
  }, [])

  const goToMatch = useCallback(
    (direction: 'next' | 'prev') => {
      if (matches.length === 0) return
      const next =
        direction === 'next'
          ? (currentMatch + 1) % matches.length
          : (currentMatch - 1 + matches.length) % matches.length
      selectMatch(next, matches)
    },
    [currentMatch, matches, selectMatch],
  )

  // Re-select the first match whenever the query (or the underlying data)
  // produces a fresh match list, mirroring useSearch's re-highlight-on-change.
  const matchesKey = matches.length > 0 ? `${matches[0].row}:${matches[0].col}:${matches.length}` : ''
  const [lastMatchesKey, setLastMatchesKey] = useState('')
  if (matchesKey !== lastMatchesKey) {
    setLastMatchesKey(matchesKey)
    if (matches.length > 0) {
      selectMatch(0, matches)
    } else {
      setCurrentMatch(0)
      setGridSelection(undefined)
    }
  }

  const open = useCallback(() => setIsOpen(true), [])

  const close = useCallback(() => {
    setIsOpen(false)
    setQuery('')
    setGridSelection(undefined)
  }, [])

  const onGridSelectionChange = useCallback((selection: GridSelection) => {
    setGridSelection(selection)
  }, [])

  return {
    isOpen,
    query,
    setQuery,
    matchCount: matches.length,
    currentMatch,
    goToMatch,
    open,
    close,
    gridSelection,
    onGridSelectionChange,
  }
}
