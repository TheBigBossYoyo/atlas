/**
 * ShortcutManager context value + React context object (P2.1 / SHELL-19 / D5).
 *
 * Kept separate from the Provider component and the consumer hooks so that:
 *   - `ShortcutManagerProvider.tsx` only exports a React component (fast-refresh friendly)
 *   - `useShortcutManager.ts` only exports hooks (fast-refresh friendly)
 *
 * See `ShortcutManagerProvider.tsx` for the full design rationale.
 */

import { createContext } from 'react'

/**
 * Extra context handed to every registered handler alongside the raw
 * `KeyboardEvent`. `inPlainField` mirrors the pre-dispatcher `inField` guard
 * every hook used to compute for itself, widened to also cover
 * `isContentEditable` (SHELL-08's own recommendation) — a shell-global
 * handler generally wants to skip command shortcuts while the user is
 * typing into a plain input/textarea or a contentEditable region (e.g.
 * DocxViewer's editor surface), but Ctrl+S/Ctrl+P intentionally do not check
 * it (Save/Print should work even mid-edit).
 */
export interface ShortcutDispatchContext {
  readonly inPlainField: boolean
}

/**
 * A registered shortcut handler. Return `true` to claim the event (the
 * handler is expected to have called `event.preventDefault()` itself when it
 * does) and stop the dispatcher from trying any lower-precedence handler.
 * Returning `false`/`undefined` lets the dispatcher keep going.
 */
export type ShortcutHandler = (event: KeyboardEvent, ctx: ShortcutDispatchContext) => boolean | void

export interface ShortcutManagerContextValue {
  /**
   * Registers a handler in the "active viewer" precedence tier — wins over
   * every shell-global handler regardless of where focus literally is within
   * the viewer (e.g. DocxViewer's own Ctrl+B/E/P/S combos). Returns an
   * unregister function; call it from the registering effect's cleanup.
   */
  registerViewerHandler: (handler: ShortcutHandler) => () => void
  /**
   * Registers a handler in the "shell-global" precedence tier (app chrome:
   * open/save/export/theme/sidebar/search/shortcuts-modal/etc.) — the lowest
   * precedence, only reached once no viewer handler has claimed the event.
   */
  registerShellHandler: (handler: ShortcutHandler) => () => void
}

export const ShortcutManagerContext = createContext<ShortcutManagerContextValue | null>(null)
