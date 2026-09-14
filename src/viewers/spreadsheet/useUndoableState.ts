/**
 * Generic undo/redo history for spreadsheet editing (wave 3).
 *
 * A plain `{past, present, future}` history stack (the standard React
 * pattern for this), kept generic over `T` so both SpreadsheetViewer
 * (`SpreadsheetDocument`, multi-sheet) and CsvViewer (also a
 * `SpreadsheetDocument`, always one sheet) can share the same hook.
 *
 * `set()` is one call per discrete, already-committed edit — a completed
 * cell edit, one row/column insert, one sheet rename, one paste — not per
 * keystroke: glide-data-grid's `onCellEdited` only fires once an edit is
 * committed (Enter/Tab/blur), so each call already corresponds to one
 * undoable step with no extra debouncing needed here.
 *
 * `maxHistory` bounds memory: past a certain number of edits the oldest
 * `past` entries are dropped rather than kept forever. Each entry is a
 * `SpreadsheetDocument`, i.e. mostly-shared-reference row/column arrays
 * (only the touched sheet/row/column is a fresh copy — see
 * `spreadsheetDocument.ts`), so this is far cheaper than it looks even for a
 * large sheet.
 */
import { useCallback, useMemo, useState } from 'react'

const DEFAULT_MAX_HISTORY = 100

type History<T> = {
  readonly past: ReadonlyArray<T>
  readonly present: T
  readonly future: ReadonlyArray<T>
}

export type UndoableState<T> = {
  readonly present: T
  readonly canUndo: boolean
  readonly canRedo: boolean
  /** Pushes `next` as a new present, clearing redo history — one call per committed edit. */
  readonly set: (next: T) => void
  readonly undo: () => void
  readonly redo: () => void
  /** Replaces the present AND clears all undo/redo history (e.g. right after a successful save, or on load). */
  readonly reset: (next: T) => void
}

export function useUndoableState<T>(initial: T, maxHistory: number = DEFAULT_MAX_HISTORY): UndoableState<T> {
  const [history, setHistory] = useState<History<T>>({ past: [], present: initial, future: [] })

  const set = useCallback(
    (next: T) => {
      setHistory((prev) => {
        const past = [...prev.past, prev.present]
        return {
          past: past.length > maxHistory ? past.slice(past.length - maxHistory) : past,
          present: next,
          future: [],
        }
      })
    },
    [maxHistory],
  )

  const undo = useCallback(() => {
    setHistory((prev) => {
      if (prev.past.length === 0) return prev
      const previous = prev.past[prev.past.length - 1]
      return {
        past: prev.past.slice(0, -1),
        present: previous,
        future: [prev.present, ...prev.future],
      }
    })
  }, [])

  const redo = useCallback(() => {
    setHistory((prev) => {
      if (prev.future.length === 0) return prev
      const [next, ...rest] = prev.future
      return {
        past: [...prev.past, prev.present],
        present: next,
        future: rest,
      }
    })
  }, [])

  const reset = useCallback((next: T) => {
    setHistory({ past: [], present: next, future: [] })
  }, [])

  return useMemo(
    () => ({
      present: history.present,
      canUndo: history.past.length > 0,
      canRedo: history.future.length > 0,
      set,
      undo,
      redo,
      reset,
    }),
    [history, set, undo, redo, reset],
  )
}
