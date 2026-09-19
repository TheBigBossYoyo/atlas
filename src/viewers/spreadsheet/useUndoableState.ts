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
import { useCallback, useMemo, useRef, useState } from 'react'

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
  /** Steps back; returns the new present right away (callers that must not wait for a render use it). */
  readonly undo: () => T
  /** Steps forward; returns the new present right away. */
  readonly redo: () => T
  /** Replaces the present AND clears all undo/redo history (e.g. right after a successful save, or on load). */
  readonly reset: (next: T) => void
}

export function useUndoableState<T>(initial: T, maxHistory: number = DEFAULT_MAX_HISTORY): UndoableState<T> {
  const [history, setHistory] = useState<History<T>>({ past: [], present: initial, future: [] })
  // The authoritative history, updated synchronously: several calls within one
  // tick (or an undo followed at once by a save) must see each other's result
  // before React re-renders.
  const historyRef = useRef(history)
  const commit = useCallback((next: History<T>) => {
    historyRef.current = next
    setHistory(next)
  }, [])

  const set = useCallback(
    (next: T) => {
      const prev = historyRef.current
      const past = [...prev.past, prev.present]
      commit({
        past: past.length > maxHistory ? past.slice(past.length - maxHistory) : past,
        present: next,
        future: [],
      })
    },
    [commit, maxHistory],
  )

  const undo = useCallback((): T => {
    const prev = historyRef.current
    if (prev.past.length === 0) return prev.present
    const previous = prev.past[prev.past.length - 1]
    commit({ past: prev.past.slice(0, -1), present: previous, future: [prev.present, ...prev.future] })
    return previous
  }, [commit])

  const redo = useCallback((): T => {
    const prev = historyRef.current
    if (prev.future.length === 0) return prev.present
    const [next, ...rest] = prev.future
    commit({ past: [...prev.past, prev.present], present: next, future: rest })
    return next
  }, [commit])

  const reset = useCallback(
    (next: T) => {
      commit({ past: [], present: next, future: [] })
    },
    [commit],
  )

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
