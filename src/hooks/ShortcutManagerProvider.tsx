/**
 * ShortcutManagerProvider — the single top-level keydown listener behind the
 * centralized shortcut dispatcher (P2.1 / SHELL-08, SHELL-09, SHELL-19,
 * RUN-02, UX-03, UX-24).
 *
 * Before this, six independent `window.addEventListener('keydown', ...)`
 * sites (`useUniversalShortcuts`, `useFontSize`, `useSearch`, `ThemeMenu`,
 * `ExportMenu`, `ShortcutsModal`) — plus `UnsavedChangesDialog`'s own Escape
 * listener and DocxViewer's element-level combo handling — all raced for the
 * same keys with no shared precedence, which is the direct architectural
 * cause of SHELL-08/09 (Ctrl+B/Ctrl+E while editing a DOCX also toggling the
 * sidebar / opening the Export menu) and UX-03 (Ctrl+P double-firing print
 * AND export).
 *
 * This provider replaces all of them with one `window` listener and an
 * explicit two-tier precedence, matching Decision D5:
 *
 *   1. "active viewer" handlers (`useViewerShortcuts`) — e.g. DocxViewer's
 *      own Ctrl+B/I/U/E/L/R/J/P/S/... combos. These win unconditionally
 *      over shell-global handlers, regardless of exactly where focus is
 *      within the viewer (not just "is the contentEditable literally
 *      focused") — the active *viewer*, not the literal DOM focus target, is
 *      what gets first refusal.
 *   2. "shell-global" handlers (`useShellShortcut`) — app-chrome shortcuts
 *      (open/save/export/theme/sidebar/search/shortcuts-modal/close-file/
 *      etc.), reached only once no viewer handler has claimed the event.
 *
 * Each handler also receives `{ inPlainField }` — whether the keydown target
 * is a plain input/textarea or any `isContentEditable` element — so a
 * shell-global handler can reproduce the old "don't fire while the user is
 * typing" guard without every hook re-deriving it (and, per SHELL-08's own
 * recommendation, that guard now correctly covers `contentEditable` too,
 * where the old per-hook checks only covered `<input>`/`<textarea>`).
 *
 * Handlers within a tier run most-recently-registered first (LIFO) — e.g. if
 * a modal opens on top of another, its own Escape handler (registered when
 * it mounted, later than the background one) gets first refusal.
 *
 * Registration order across renders never matters for correctness: each
 * `useShellShortcut`/`useViewerShortcuts` call keeps its own latest handler
 * closure in a ref (see `useShortcutManager.ts`) and only registers/
 * unregisters the *wrapper* when `enabled` changes, so callers can pass a
 * fresh inline closure on every render with no memoization burden.
 */

import { useCallback, useEffect, useMemo, useRef, type ReactNode } from 'react'

import {
  ShortcutManagerContext,
  type ShortcutDispatchContext,
  type ShortcutHandler,
  type ShortcutManagerContextValue,
} from './shortcutManagerContext'

/**
 * Whether `target` is a plain text-entry surface: an `<input>`/`<textarea>`,
 * or an element inside a `contenteditable` region (DocxViewer's editor
 * surface). Real browsers expose this directly as `Element.isContentEditable`,
 * but jsdom (this project's Vitest environment) implements neither
 * `isContentEditable` nor even the `contentEditable` string reflection (both
 * come back `undefined`) — falling back to an explicit `contenteditable`
 * attribute lookup up the ancestor chain keeps this correct under test while
 * `isContentEditable` continues to be used first wherever it's reliable.
 */
function isPlainFieldTarget(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) return false
  if (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) return true
  if (target.isContentEditable) return true
  return target.closest('[contenteditable=""], [contenteditable="true"]') !== null
}

export function ShortcutManagerProvider({ children }: { children: ReactNode }): React.ReactElement {
  // Plain mutable arrays behind refs, not state — a keydown can happen many
  // times a second and none of this should ever trigger a React re-render;
  // only the one native listener below observes it.
  const viewerHandlersRef = useRef<ReadonlyArray<ShortcutHandler>>([])
  const shellHandlersRef = useRef<ReadonlyArray<ShortcutHandler>>([])

  const registerViewerHandler = useCallback((handler: ShortcutHandler): (() => void) => {
    viewerHandlersRef.current = [...viewerHandlersRef.current, handler]
    return () => {
      viewerHandlersRef.current = viewerHandlersRef.current.filter((h) => h !== handler)
    }
  }, [])

  const registerShellHandler = useCallback((handler: ShortcutHandler): (() => void) => {
    shellHandlersRef.current = [...shellHandlersRef.current, handler]
    return () => {
      shellHandlersRef.current = shellHandlersRef.current.filter((h) => h !== handler)
    }
  }, [])

  useEffect(() => {
    function dispatch(event: KeyboardEvent): void {
      const ctx: ShortcutDispatchContext = { inPlainField: isPlainFieldTarget(event.target) }

      const viewerHandlers = viewerHandlersRef.current
      for (let i = viewerHandlers.length - 1; i >= 0; i -= 1) {
        if (viewerHandlers[i](event, ctx)) return
      }

      const shellHandlers = shellHandlersRef.current
      for (let i = shellHandlers.length - 1; i >= 0; i -= 1) {
        if (shellHandlers[i](event, ctx)) return
      }
    }

    window.addEventListener('keydown', dispatch)
    return () => window.removeEventListener('keydown', dispatch)
  }, [])

  const value = useMemo<ShortcutManagerContextValue>(
    () => ({ registerViewerHandler, registerShellHandler }),
    [registerViewerHandler, registerShellHandler],
  )

  return <ShortcutManagerContext.Provider value={value}>{children}</ShortcutManagerContext.Provider>
}
