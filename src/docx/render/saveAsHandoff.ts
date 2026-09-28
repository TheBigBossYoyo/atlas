/**
 * DOCX-3 (second half) — carries the caret and the scroll position across the
 * remount a Save As causes.
 *
 * `useDocxSave` already restores focus and the native selection once the OS save
 * dialog closes, which is the half of the problem that belongs to the component.
 * It is not enough on its own: `reportSavedPath` tells the shell the document now
 * lives at a new path, `ViewerRouter` keys the whole viewer subtree by
 * `file.path`, and a render or two later a brand-new `DocxEditor` replaces the
 * one whose focus was just restored. The user is left with `document.activeElement`
 * on `<body>` and no model selection, so typing silently does nothing until they
 * click back into the page.
 *
 * Not fixed by keeping the viewer mounted instead. That means `ViewerRouter` no
 * longer remounting on a path change, which is what makes `ViewerErrorBoundary`
 * clear a cached crash (LOAD-08/RUN-09) and what lets every viewer seed state
 * from `file.path`/`bundle` with no resync effect — see `ViewerRouter.tsx`'s own
 * comment on the per-viewer audit that change needs. That audit is worth doing
 * and is still open; it is not worth doing as the price of restoring a caret.
 *
 * So: a single-slot handoff, in a module rather than in React state, because the
 * entire point is that no part of the tree survives to hold it. It is written
 * only by the Save As path, only when the path actually changed (the one case
 * that remounts), and claimed by the next mount whose own path matches. A slot
 * that is never claimed — Save As onto the same path, a viewer that never
 * finishes rendering, a switch to a different document first — ages out and is
 * discarded rather than being applied to the wrong document later.
 */
import { useEffect, type RefObject } from 'react'

import type { Range } from '../editor'

type SaveAsHandoff = {
  /** The path the remounted viewer will be showing; the slot is claimed by matching it. */
  readonly targetPath: string
  readonly range: Range | null
  readonly scrollTop: number
  readonly scrollLeft: number
  readonly stashedAt: number
}

/**
 * Generous: the remount was measured at ~300ms after the save resolves, but the
 * document has to re-parse and paginate first, and a large one on a slow machine
 * is the case this most needs to work for. The bound only has to be short enough
 * that a slot which was never going to be claimed cannot survive until the user
 * opens some other document.
 */
const MAX_AGE_MS = 15_000

let pending: SaveAsHandoff | null = null

export function stashSaveAsHandoff(handoff: Omit<SaveAsHandoff, 'stashedAt'>): void {
  pending = { ...handoff, stashedAt: Date.now() }
}

/**
 * Returns and clears the pending handoff when it was stashed for `path`.
 *
 * Clears a stale or foreign slot too, rather than leaving it for a later mount:
 * if the next viewer to mount is not the one the Save As renamed, restoring a
 * caret recorded in a different document is worse than restoring nothing.
 */
export function claimSaveAsHandoff(path: string): Omit<SaveAsHandoff, 'stashedAt'> | null {
  const claimed = pending
  if (claimed === null) return null

  pending = null
  if (claimed.targetPath !== path) return null
  if (Date.now() - claimed.stashedAt > MAX_AGE_MS) return null

  return {
    targetPath: claimed.targetPath,
    range: claimed.range,
    scrollTop: claimed.scrollTop,
    scrollLeft: claimed.scrollLeft,
  }
}

/** Test-only: drops any pending slot so one test's stash cannot leak into the next. */
export function resetSaveAsHandoff(): void {
  pending = null
}

/**
 * Claims the handoff on behalf of a freshly mounted viewer and puts the caret,
 * focus and scroll position back.
 *
 * `ready` must mean "the document is laid out", not "the component mounted":
 * restoring a model selection before there is any `[data-paragraph-path]` DOM to
 * map it onto leaves the native selection unset, and the effect that syncs the
 * two only re-runs when `pages` or `range` change — neither of which would
 * change again on its own. So the claim waits for the first render that has
 * pages. An unclaimed slot ages out; it is never applied to a later document.
 *
 * Focus is taken before the selection is set on purpose: the selection-painting
 * effect checks whether focus sits in a field OUTSIDE the editor and, if it
 * does, paints a highlight instead of moving the DOM selection — which would
 * leave typing still doing nothing, the exact symptom this exists to fix.
 */
export function useSaveAsFocusRestore(
  path: string,
  ready: boolean,
  editorRootRef: RefObject<HTMLElement | null>,
  setRange: (range: Range) => void,
): void {
  useEffect(() => {
    if (!ready) return
    const root = editorRootRef.current
    if (root === null) return

    const handoff = claimSaveAsHandoff(path)
    if (handoff === null) return

    root.focus({ preventScroll: true })
    // Best-effort, and secondary to the caret: setting the DOM selection can
    // scroll the scrollport itself in Chromium. The caret being back where the
    // user left it usually puts the view back with it.
    root.scrollTop = handoff.scrollTop
    root.scrollLeft = handoff.scrollLeft
    if (handoff.range !== null) setRange(handoff.range)
  }, [editorRootRef, path, ready, setRange])
}
