/**
 * ViewerContext value type + React context object (W2.X2 — split for fast-refresh)
 *
 * Kept separate from the Provider component and the consumer hooks so that:
 *   - `ViewerContext.tsx` only exports a React component (fast-refresh friendly)
 *   - `useViewerContext.ts` only exports hooks (fast-refresh friendly)
 */

import { createContext } from 'react'
import type { NavItem, ViewerStats } from '../../formats/types'

export type ViewerContextValue = {
  navItems: ReadonlyArray<NavItem>
  setNavItems: (items: ReadonlyArray<NavItem>) => void
  stats: ViewerStats | null
  setStats: (stats: ViewerStats | null) => void
}

export const ViewerContext = createContext<ViewerContextValue | null>(null)
