/**
 * ViewerProvider — Nav/Stats/document-session context provider component (W2.X2, P1.1).
 *
 * Lets viewers publish nav items + stats and lets the app shell
 * (Sidebar/StatusBar) consume them. Also carries the document-session
 * capability contract (P1.1): the active viewer reports `isDirty` and
 * registers its own `save()` implementation here, so App.tsx can combine it
 * with markdown's own dirty state and route a global Ctrl+S/Save to whichever
 * viewer is actually active.
 *
 * Everything resets when `filePath` changes — a new file means a fresh
 * session: no nav items/stats yet, nothing dirty, and no stale save handler
 * left registered by whichever viewer was just unmounted.
 *
 * Context value lives in `viewerContextValue.ts` and consumer hooks live in
 * `useViewerContext.ts` so this file only exports a React component
 * (satisfies `react-refresh/only-export-components`).
 *
 * The reset-on-filePath-change uses the React 19 "store the prop in state"
 * idiom (https://react.dev/reference/react/useState#storing-information-from-previous-renders),
 * which avoids both `react-hooks/refs` (no render-time ref mutation) and
 * `react-hooks/set-state-in-effect` (no cascading-render effect).
 */

import React, { useCallback, useMemo, useRef, useState } from 'react'

import type { NavItem, ViewerStats } from '../../formats/types'
import { ViewerContext, type ExportableContent, type ViewerContextValue } from './viewerContextValue'

type InternalState = {
  filePath: string | null
  navItems: ReadonlyArray<NavItem>
  stats: ViewerStats | null
  isDirty: boolean
  canFind: boolean
}

const INITIAL_STATE_EXCEPT_PATH = {
  navItems: [] as ReadonlyArray<NavItem>,
  stats: null,
  isDirty: false,
  canFind: false,
} as const

export function ViewerProvider({
  filePath,
  children,
}: {
  filePath: string | null
  children: React.ReactNode
}): React.ReactElement {
  const [state, setState] = useState<InternalState>({
    filePath,
    ...INITIAL_STATE_EXCEPT_PATH,
  })

  // The registered save implementation lives in a ref, not state: it is an
  // implementation detail the active viewer swaps in/out (via `registerSave`)
  // and does not itself need to trigger a re-render. It does not need an
  // explicit reset here when `filePath` changes: the viewer that registered
  // it unmounts as part of that same file switch (ViewerRouter remounts on
  // `file.path`) and its own cleanup effect calls `registerSave(null)` before
  // any new save() call could reach it.
  const saveRef = useRef<(() => Promise<boolean>) | null>(null)

  // Same reasoning as `saveRef` above: the active viewer's find-overlay
  // opener is an implementation detail swapped in/out via `registerFind`,
  // not something that itself needs to trigger a re-render. `canFind`
  // (whether *something* is registered) does live in state, since Toolbar's
  // search button needs to reactively enable/disable.
  const findRef = useRef<(() => void) | null>(null)

  // setState-during-render reset: when the incoming filePath differs from the
  // stored one, schedule fresh InternalState. React discards the in-progress
  // render and re-runs with the new state — the lint-clean way to "reset
  // state when a prop changes" (no render-time ref mutation, no effect).
  if (state.filePath !== filePath) {
    setState({ filePath, ...INITIAL_STATE_EXCEPT_PATH })
  }

  const setNavItems = useCallback((items: ReadonlyArray<NavItem>) => {
    setState(prev => ({ ...prev, navItems: items }))
  }, [])

  const setStats = useCallback((stats: ViewerStats | null) => {
    setState(prev => ({ ...prev, stats }))
  }, [])

  const setDirty = useCallback((dirty: boolean) => {
    setState(prev => (prev.isDirty === dirty ? prev : { ...prev, isDirty: dirty }))
  }, [])

  const registerSave = useCallback((save: (() => Promise<boolean>) | null) => {
    saveRef.current = save
  }, [])

  const save = useCallback(async (): Promise<boolean> => {
    if (saveRef.current === null) {
      return false
    }
    return saveRef.current()
  }, [])

  // P1.1 placeholder — no viewer registers export content yet; real
  // implementation deferred to the Phase 3 per-format export work.
  const getExportableContent = useCallback((): ExportableContent | null => null, [])

  const registerFind = useCallback((find: (() => void) | null) => {
    findRef.current = find
    setState(prev => (prev.canFind === (find !== null) ? prev : { ...prev, canFind: find !== null }))
  }, [])

  const openFind = useCallback(() => {
    findRef.current?.()
  }, [])

  const value = useMemo<ViewerContextValue>(
    () => ({
      navItems: state.navItems,
      setNavItems,
      stats: state.stats,
      setStats,
      isDirty: state.isDirty,
      setDirty,
      registerSave,
      save,
      getExportableContent,
      canFind: state.canFind,
      registerFind,
      openFind,
    }),
    [
      state.navItems,
      state.stats,
      state.isDirty,
      state.canFind,
      setNavItems,
      setStats,
      setDirty,
      registerSave,
      save,
      getExportableContent,
      registerFind,
      openFind,
    ],
  )

  return (
    <ViewerContext.Provider value={value}>{children}</ViewerContext.Provider>
  )
}
