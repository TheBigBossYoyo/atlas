/**
 * Keeping what the user SEES selected in step with the DOCX editor's model
 * `range`, after every render that could have moved either one.
 *
 * Extracted from `src/viewers/DocxViewer.tsx` as part of breaking up that
 * file's one very large component. It is a layout effect rather than a passive
 * one because it reads and writes DOM selection/scroll geometry, which must
 * settle in the same commit the new `pages` painted in — a passive effect would
 * let the browser show one frame with the caret in the old place.
 *
 * It also owns the "scroll the next painted selection into view" flag. That
 * flag is set by whoever MOVED the selection programmatically (Find next/prev,
 * outline navigation) and consumed here, because only this effect knows when
 * the target paragraph actually exists in the DOM: `PageStack` virtualizes
 * pages, so a match on a page that was not mounted has nothing to scroll to
 * until the render that mounts it. Exposing it as `revealSelectionOnNextPaint()`
 * rather than handing the ref around keeps that one-way — callers announce an
 * intent, they do not reach into this effect's state.
 */
import { useCallback, useEffect, useLayoutEffect, useRef, type RefObject } from 'react'

import { positionToDomRange } from '../editor/Cursor'
import type { Range } from '../editor'
import type { Document as DocxDocument } from '../model'
import type { Page } from '../layout'
import { paintSelectionHighlight, syncSelectionToDom } from './selectionDom'

export type DocxSelectionPainting = {
  /**
   * Ask the next painted selection to be scrolled into view. Call it right
   * before setting a range you moved programmatically; it is consumed once.
   */
  readonly revealSelectionOnNextPaint: () => void
}

/**
 * @param editorRootRef   - the contentEditable surface. A null ref (not yet
 *                          mounted, or an error state rendering instead) is a
 *                          no-op, not an error.
 * @param range           - the model selection to paint.
 * @param documentModelRef - read through a ref because this effect must use the
 *                          model as of THIS commit without taking `documentModel`
 *                          as a dependency, which would re-run it on a keystroke
 *                          that did not move the selection.
 * @param pages           - a dependency, not an input: a new pagination pass
 *                          replaces the DOM nodes the selection points at, so
 *                          the selection has to be re-applied afterwards.
 */
export function useDocxSelectionPainting(
  editorRootRef: RefObject<HTMLDivElement | null>,
  range: Range | null,
  documentModelRef: RefObject<DocxDocument>,
  pages: ReadonlyArray<Page> | null,
): DocxSelectionPainting {
  // USR-06 — set by Find next/prev so the post-render selection sync scrolls
  // the match into view.
  const revealSelectionRef = useRef(false)

  const revealSelectionOnNextPaint = useCallback(() => {
    revealSelectionRef.current = true
  }, [])

  useLayoutEffect(() => {
    const root = editorRootRef.current
    if (root === null) {
      return
    }

    // USR-06 — placing the DOM selection inside a contentEditable moves
    // keyboard focus into it in Chromium, which stole focus from the Find
    // box after the first match (Enter then edited the document). While a
    // form field outside the editor has focus, paint the model selection with
    // the CSS Custom Highlight API instead of moving the DOM selection.
    const activeElement = root.ownerDocument.activeElement
    const fieldHasFocus =
      activeElement !== null &&
      !root.contains(activeElement) &&
      activeElement.matches('input, textarea, select, [contenteditable="true"]')
    if (fieldHasFocus) {
      paintSelectionHighlight(root, range, documentModelRef.current)
    } else {
      paintSelectionHighlight(root, null)
      syncSelectionToDom(root, range, documentModelRef.current)
    }

    if (revealSelectionRef.current && range !== null) {
      revealSelectionRef.current = false
      const point = positionToDomRange(range.focus, root, documentModelRef.current)
      const node = point?.node ?? null
      const element = node === null ? null : node.nodeType === Node.ELEMENT_NODE ? (node as Element) : node.parentElement
      if (element !== null && typeof element.scrollIntoView === 'function') {
        element.scrollIntoView({ block: 'center', inline: 'nearest' })
      }
    }
    // `editorRootRef`/`documentModelRef` are refs, deliberately not dependencies
    // — see the param docs above for why the model is read through one.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pages, range])

  // The highlight registry is global to the document, so a viewer that unmounts
  // while its selection is painted would leave that highlight behind for
  // whatever mounts next.
  useEffect(() => () => paintSelectionHighlight(null, null), [])

  return { revealSelectionOnNextPaint }
}
