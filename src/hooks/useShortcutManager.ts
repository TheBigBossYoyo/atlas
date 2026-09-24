/**
 * Consumer hooks for the centralized shortcut dispatcher (P2.1 / SHELL-19 / D5).
 *
 * See `ShortcutManagerProvider.tsx` for the full precedence design. Lives in
 * its own file so it only exports hooks (satisfies
 * `react-refresh/only-export-components`), matching the split already used
 * for `ViewerContext.tsx`/`useViewerContext.ts`.
 *
 * Both hooks below tolerate rendering without a `<ShortcutManagerProvider>`
 * ancestor (silently registering nothing) rather than throwing, unlike
 * `useViewerContext.ts`'s strict `useViewerCtx()` — many existing component
 * tests (`Toolbar.test.tsx`, `DocxViewer.editor.test.tsx`, etc.) render a
 * single component tree without the app's full provider stack, and a
 * shortcut registration that silently no-ops there is the correct behavior
 * for a test harness, not a bug to work around.
 */

import { useContext, useLayoutEffect, useRef } from 'react'

import { ShortcutManagerContext, type ShortcutHandler } from './shortcutManagerContext'

function useRegisteredShortcut(
  scope: 'viewer' | 'shell',
  handler: ShortcutHandler,
  enabled: boolean,
): void {
  const manager = useContext(ShortcutManagerContext)

  // "Latest ref" pattern: keep the newest handler closure available to the
  // registered wrapper without re-registering (and thus reordering the
  // precedence stack) on every render just because the caller passed a new
  // inline closure. Assigned in an effect, never during render, per the
  // project's `react-hooks/refs` lint rule.
  //
  // Layout effects, not passive ones (REF-EFFECT-1): a key pressed after a
  // render has committed but before its passive effects ran would otherwise
  // be dispatched to the PREVIOUS render's handler — e.g. Ctrl+2 rejected by
  // a closure from before the document loaded (`isMarkdown` still false), or
  // Ctrl+S calling a `saveFile` that captured the text one keystroke ago.
  // Layout effects run in the same commit, so what is on screen and what
  // handles the next key can never disagree.
  const handlerRef = useRef(handler)
  useLayoutEffect(() => {
    handlerRef.current = handler
  });

  useLayoutEffect(() => {
    if (!enabled || !manager) return undefined
    const stableHandler: ShortcutHandler = (event, ctx) => handlerRef.current(event, ctx)
    return scope === 'viewer'
      ? manager.registerViewerHandler(stableHandler)
      : manager.registerShellHandler(stableHandler)
  }, [enabled, manager, scope]);
}

/**
 * Registers a shell-global (app-chrome) shortcut handler — the lowest
 * precedence tier, reached only once no active-viewer handler has claimed
 * the event. Pass `enabled: false` to temporarily stop listening (e.g. a
 * menu's own Escape handler that should only listen while open) instead of
 * conditionally calling the hook, which would violate the rules of hooks.
 */
export function useShellShortcut(handler: ShortcutHandler, enabled = true): void {
  useRegisteredShortcut('shell', handler, enabled)
}

/**
 * Registers an active-viewer shortcut handler (e.g. DocxViewer's own
 * Ctrl+B/I/U/E/L/R/J/P/S/... combos) — wins over every shell-global handler
 * regardless of exactly where focus is within the viewer. Exported so a
 * future viewer (PDF/PPTX/ODP internal navigation, per P2.1's scope note)
 * can adopt the same precedence without this wave rewriting their existing
 * internal handlers.
 */
export function useViewerShortcuts(handler: ShortcutHandler, enabled = true): void {
  useRegisteredShortcut('viewer', handler, enabled)
}
