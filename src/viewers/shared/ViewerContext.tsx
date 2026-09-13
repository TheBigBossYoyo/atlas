/**
 * ViewerProvider — Nav/Stats context provider component (W2.X2).
 *
 * Lets viewers publish nav items + stats and lets the app shell
 * (Sidebar/StatusBar) consume them. Auto-clears when `filePath` changes.
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

import React, { useCallback, useMemo, useState } from 'react'

import type { NavItem, ViewerStats } from '../../formats/types'
import { ViewerContext, type ViewerContextValue } from './viewerContextValue'

type InternalState = {
  filePath: string | null
  navItems: ReadonlyArray<NavItem>
  stats: ViewerStats | null
}

export function ViewerProvider({
  filePath,
  children,
}: {
  filePath: string | null
  children: React.ReactNode
}): React.ReactElement {
  const [state, setState] = useState<InternalState>({
    filePath,
    navItems: [],
    stats: null,
  })

  // setState-during-render reset: when the incoming filePath differs from the
  // stored one, schedule fresh InternalState. React discards the in-progress
  // render and re-runs with the new state — the lint-clean way to "reset state
  // when a prop changes" (no render-time ref mutation, no effect).
  if (state.filePath !== filePath) {
    setState({ filePath, navItems: [], stats: null })
  }

  const setNavItems = useCallback((items: ReadonlyArray<NavItem>) => {
    setState(prev => ({ ...prev, navItems: items }))
  }, [])

  const setStats = useCallback((stats: ViewerStats | null) => {
    setState(prev => ({ ...prev, stats }))
  }, [])

  const value = useMemo<ViewerContextValue>(
    () => ({
      navItems: state.navItems,
      setNavItems,
      stats: state.stats,
      setStats,
    }),
    [state.navItems, state.stats, setNavItems, setStats],
  )

  return (
    <ViewerContext.Provider value={value}>{children}</ViewerContext.Provider>
  )
}
