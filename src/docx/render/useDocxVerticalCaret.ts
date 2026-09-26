/**
 * DOCX-SMALL-WINDOW-2 — retrying an Up/Down/PageUp/PageDown caret move that had
 * to wait for its target page to mount.
 *
 * Extracted from `src/viewers/DocxViewer.tsx`. `PageStack` virtualizes pages, so
 * only the page(s) the current selection sits on are guaranteed to be in the DOM.
 * An adjacent page that has not scrolled into view is a bare placeholder with no
 * `.docx-page__line` elements at all — and to `positionOnAdjacentLine`'s geometry
 * search that is indistinguishable from "there is no more document". On a small
 * window, where fewer pages are mounted at once, that made Down at a page
 * boundary look like the end of the file.
 *
 * So `handleKeyDownEvent` defers instead of guessing: it calls
 * `requestVerticalMove` with the page index that needs mounting, the viewer pins
 * that page (via `pendingMountPageIndex`), and the layout effect below retries
 * the move once the page has real DOM.
 *
 * It is a LAYOUT effect, not a passive one, deliberately: it must land in the same
 * paint as the newly mounted page, or the user sees the caret sitting still for a
 * frame before it jumps. That constraint also rules out the workarounds this
 * codebase applies elsewhere to keep a state reset out of an effect body — see the
 * note above the effect.
 */
import { useCallback, useLayoutEffect, useState, type RefObject } from 'react'

import { extendOrCollapse, moveCursorToLineEnd, moveCursorToLineStart, type Range } from '../editor'
import { positionOnAdjacentLine, positionOnePageVertically, type VerticalDirection } from '../editor/Cursor'
import type { Document as DocxDocument } from '../model'

/** Everything the retry needs to reproduce the move the keydown handler deferred. */
export type PendingVerticalMove = {
  readonly direction: VerticalDirection
  readonly goalX: number
  readonly isPageMove: boolean
  readonly fromRect: DOMRect
  readonly viewportHeight: number
  readonly shift: boolean
  readonly baseRange: Range
  readonly mountPageIndex: number
}

export type DocxVerticalCaret = {
  /**
   * The page index the viewer must pin so the retry has real DOM to search, or
   * null when no move is pending. Feeds `pinnedPageIndices`.
   */
  readonly pendingMountPageIndex: number | null
  /** Defers a vertical move until `move.mountPageIndex` has mounted. */
  readonly requestVerticalMove: (move: PendingVerticalMove) => void
}

/**
 * @param documentModel  - the live model the geometry search runs against.
 * @param editorRootRef  - the contentEditable surface. A null ref means the
 *                         viewer went away mid-flight; the request is dropped.
 * @param setRange       - where the retried move lands.
 */
export function useDocxVerticalCaret(
  documentModel: DocxDocument,
  editorRootRef: RefObject<HTMLDivElement | null>,
  setRange: (range: Range | null) => void,
): DocxVerticalCaret {
  const [pendingVerticalMove, setPendingVerticalMove] = useState<PendingVerticalMove | null>(null)

  const requestVerticalMove = useCallback((move: PendingVerticalMove) => {
    setPendingVerticalMove(move)
  }, [])

  // Clearing `pendingVerticalMove` from inside this effect is the effect's job —
  // it consumes a ONE-SHOT request — and is self-terminating: the next render
  // sees `null` and returns at the top, so there is no cascade. Neither idiom this
  // codebase uses elsewhere to keep such a reset out of an effect body would work
  // here anyway: an async IIFE (see `useDocxPagination`) cannot read layout
  // geometry synchronously, and a setState-during-render reset (see
  // `useDocxComments`) cannot run after a paint.
  useLayoutEffect(() => {
    if (pendingVerticalMove === null) {
      return
    }
    const root = editorRootRef.current
    if (root === null) {
      setPendingVerticalMove(null)
      return
    }

    const { direction, goalX, isPageMove, fromRect, viewportHeight, shift, baseRange } = pendingVerticalMove

    const target = isPageMove
      ? positionOnePageVertically(fromRect, goalX, direction, viewportHeight, root, documentModel)
      : positionOnAdjacentLine(baseRange.focus, goalX, direction, root, documentModel)

    // If the target page's content still doesn't yield a line to land on (a
    // genuinely short/empty page, or content that failed to mount for some other
    // reason), fall back to the same line-start/end clamp `handleKeyDownEvent`
    // itself uses — this must never leave the caret stuck with no visible
    // response to the keypress.
    const newFocus =
      target ??
      (direction === 'up'
        ? moveCursorToLineStart(baseRange.focus, documentModel)
        : moveCursorToLineEnd(baseRange.focus, documentModel))

    setRange(extendOrCollapse(baseRange, newFocus, shift))
    setPendingVerticalMove(null)
    // `editorRootRef` is a ref and `setRange` a stable setter, so neither is a
    // dependency; re-running on either would defeat the one-shot.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pendingVerticalMove, documentModel])

  return {
    pendingMountPageIndex: pendingVerticalMove?.mountPageIndex ?? null,
    requestVerticalMove,
  }
}
