/**
 * Consumer hooks for the viewer nav/stats/document-session context (W2.X2, P1.1).
 *
 * Lives in its own file so the Provider component file only exports a component
 * (satisfies `react-refresh/only-export-components`).
 */

import { useContext } from 'react'

import type { NavItem, ViewerStats } from '../../formats/types'
import { ViewerContext, type ExportableContent, type ViewerContextValue } from './viewerContextValue'

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

/** Whether the active viewer has unsaved changes (P1.1 document-session contract). */
export function useViewerIsDirty(): boolean {
  return useViewerCtx().isDirty
}

export function useSetViewerDirty(): (dirty: boolean) => void {
  return useViewerCtx().setDirty
}

/** The active viewer calls this to plug its own save implementation into the shared contract. */
export function useRegisterViewerSave(): (save: (() => Promise<boolean>) | null) => void {
  return useViewerCtx().registerSave
}

/** Invokes whichever save implementation the active viewer has registered (or resolves `false` if none has). */
export function useViewerSave(): () => Promise<boolean> {
  return useViewerCtx().save
}

/** The active viewer calls this to plug its own "Save As" implementation into the shared contract. */
export function useRegisterViewerSaveAs(): (saveAs: (() => Promise<boolean>) | null) => void {
  return useViewerCtx().registerSaveAs
}

/** Invokes whichever "Save As" implementation the active viewer has registered (or resolves `false` if none has). */
export function useViewerSaveAs(): () => Promise<boolean> {
  return useViewerCtx().saveAs
}

/** Reports the path a Save As landed on, so the shell moves this document's tab there. */
export function useReportSavedPath(): (path: string) => void {
  return useViewerCtx().reportSavedPath
}

/** Phase-3 placeholder — always resolves `null` until a viewer registers real export content. */
export function useGetExportableContent(): () => ExportableContent | null {
  return useViewerCtx().getExportableContent
}

/** Whether the active viewer has registered an in-viewer find implementation. */
export function useCanFindViewer(): boolean {
  return useViewerCtx().canFind
}

/** The active viewer calls this to plug its own "open find" implementation into the shared contract. */
export function useRegisterViewerFind(): (find: (() => void) | null) => void {
  return useViewerCtx().registerFind
}

/** Invokes whichever find implementation the active viewer has registered (a no-op if none has). */
export function useOpenViewerFind(): () => void {
  return useViewerCtx().openFind
}
