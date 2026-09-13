/**
 * Consumer hooks for the viewer nav/stats context (W2.X2).
 *
 * Lives in its own file so the Provider component file only exports a component
 * (satisfies `react-refresh/only-export-components`).
 */

import { useContext } from 'react'

import type { NavItem, ViewerStats } from '../../formats/types'
import { ViewerContext, type ViewerContextValue } from './viewerContextValue'

function useViewerCtx(): ViewerContextValue {
  const ctx = useContext(ViewerContext)
  if (!ctx) throw new Error('ViewerContext: missing <ViewerProvider>')
  return ctx
}

export function useNavItems(): ReadonlyArray<NavItem> {
  return useViewerCtx().navItems
}

export function useSetNavItems(): (items: ReadonlyArray<NavItem>) => void {
  return useViewerCtx().setNavItems
}

export function useViewerStats(): ViewerStats | null {
  return useViewerCtx().stats
}

export function useSetViewerStats(): (stats: ViewerStats | null) => void {
  return useViewerCtx().setStats
}
