/**
 * ViewerContext value type + React context object (W2.X2 — split for fast-refresh)
 *
 * Kept separate from the Provider component and the consumer hooks so that:
 *   - `ViewerContext.tsx` only exports a React component (fast-refresh friendly)
 *   - `useViewerContext.ts` only exports hooks (fast-refresh friendly)
 *
 * P1.1 — document-session capability contract. `setNavItems`/`setStats` publish
 * chrome data; `isDirty`/`setDirty`/`registerSave`/`save` let the active viewer
 * (currently DocxViewer) plug into the same dirty/save lifecycle markdown has
 * always had, so App.tsx can combine both into one dirty signal and route a
 * global Ctrl+S/Save to whichever viewer is actually active. `getExportableContent`
 * is a typed placeholder for the per-format export work in Phase 3 — no viewer
 * registers one yet, so it always resolves to `null` today.
 *
 * `canFind`/`registerFind`/`openFind` (Wave 3-T) mirror the same
 * register-from-the-active-viewer pattern for in-viewer find: SpreadsheetViewer
 * and CsvViewer register their own find-overlay opener here, so App.tsx's
 * Ctrl+F/Toolbar-search dispatcher can reach whichever viewer is active
 * without knowing anything about grids. Markdown's own Ctrl+F handling
 * (`useSearch`) is untouched — this contract only ever applies to non-markdown
 * viewers that opt in.
 */

import { createContext } from 'react'
import type { NavItem, ViewerStats } from '../../formats/types'

/** Placeholder shape for a future format-aware export (Phase 3). Unused today. */
export type ExportableContent = {
  readonly format: string
  readonly suggestedName: string
  readonly data: string | ArrayBuffer
}

export type ViewerContextValue = {
  navItems: ReadonlyArray<NavItem>
  setNavItems: (items: ReadonlyArray<NavItem>) => void
  stats: ViewerStats | null
  setStats: (stats: ViewerStats | null) => void
  /** Whether the active viewer has unsaved changes since its last load/save. */
  isDirty: boolean
  setDirty: (dirty: boolean) => void
  /**
   * The active viewer registers its own save implementation here (or `null` to
   * unregister, e.g. on unmount). `save()` calls whatever is currently
   * registered and resolves `false` when nothing is.
   */
  registerSave: (save: (() => Promise<boolean>) | null) => void
  save: () => Promise<boolean>
  /** Phase-3 placeholder — no viewer registers this yet. */
  getExportableContent: () => ExportableContent | null
  /** Whether the active viewer has registered an in-viewer find implementation. */
  canFind: boolean
  /**
   * The active viewer registers its own "open find" implementation here (or
   * `null` to unregister, e.g. on unmount). `openFind()` calls whatever is
   * currently registered and is a no-op when nothing is.
   */
  registerFind: (find: (() => void) | null) => void
  openFind: () => void
}

export const ViewerContext = createContext<ViewerContextValue | null>(null)
