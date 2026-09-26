/**
 * Reading and writing the DOM selection for the DOCX editor's contentEditable
 * surface, and painting a model range without touching that selection.
 *
 * Extracted from `src/viewers/DocxViewer.tsx`. Three functions, one subject: the
 * translation between a model `Range` (paragraph/run/offset addresses) and what
 * the browser's own selection API holds. The viewer decides WHEN to call them;
 * everything here is stateless and takes its root element as an argument, which
 * is also what makes it testable without mounting the whole viewer.
 *
 * `paintSelectionHighlight` deliberately uses the CSS Custom Highlight API
 * instead of the DOM selection - see its own comment (USR-06).
 */
import { domPointToPosition, positionToDomRange } from '../editor/Cursor'
import type { Range } from '../editor'
import type { Document as DocxDocument } from '../model'

const SELECTION_HIGHLIGHT_NAME = 'docx-selection'

// The CSS Custom Highlight API is not in TypeScript's DOM lib yet, so the two
// shapes this module touches are declared structurally and feature-detected at
// the call site rather than asserted to exist.
type HighlightRegistry = { set: (name: string, highlight: unknown) => void; delete: (name: string) => void }
type HighlightConstructor = new (...ranges: globalThis.Range[]) => unknown

export function getSelectionFromDom(root: HTMLElement, docModel?: DocxDocument): Range | null {
  const domSelection = root.ownerDocument.getSelection()
  if (domSelection === null || domSelection.rangeCount === 0) {
    return null
  }

  const anchorNode = domSelection.anchorNode
  const focusNode = domSelection.focusNode
  if (anchorNode === null || focusNode === null) {
    return null
  }

  if (!root.contains(anchorNode) || !root.contains(focusNode)) {
    return null
  }

  const anchor = domPointToPosition(anchorNode, domSelection.anchorOffset, root.ownerDocument, docModel)
  const focus = domPointToPosition(focusNode, domSelection.focusOffset, root.ownerDocument, docModel)

  return anchor && focus ? { anchor, focus } : null
}

/**
 * USR-06 — paints a model range without touching the DOM selection (and so
 * without stealing focus), using the CSS Custom Highlight API. `null` clears
 * it. A no-op where the API is unavailable (e.g. jsdom).
 */
export function paintSelectionHighlight(root: HTMLElement | null, range: Range | null, docModel?: DocxDocument): void {
  const registry = (globalThis.CSS as unknown as { highlights?: HighlightRegistry } | undefined)?.highlights
  const HighlightCtor = (globalThis as unknown as { Highlight?: HighlightConstructor }).Highlight
  if (registry === undefined || HighlightCtor === undefined) {
    return
  }
  if (root === null || range === null) {
    registry.delete(SELECTION_HIGHLIGHT_NAME)
    return
  }
  const anchor = positionToDomRange(range.anchor, root, docModel)
  const focus = positionToDomRange(range.focus, root, docModel)
  if (anchor === null || focus === null) {
    registry.delete(SELECTION_HIGHLIGHT_NAME)
    return
  }
  const domRange = root.ownerDocument.createRange()
  domRange.setStart(anchor.node, anchor.offset)
  domRange.setEnd(focus.node, focus.offset)
  if (domRange.collapsed) {
    domRange.setStart(focus.node, focus.offset)
    domRange.setEnd(anchor.node, anchor.offset)
  }
  registry.set(SELECTION_HIGHLIGHT_NAME, new HighlightCtor(domRange))
}

export function syncSelectionToDom(root: HTMLElement, range: Range | null, docModel?: DocxDocument): void {
  if (range === null) {
    return
  }

  const anchor = positionToDomRange(range.anchor, root, docModel)
  const focus = positionToDomRange(range.focus, root, docModel)
  if (anchor === null || focus === null) {
    return
  }

  const selection = root.ownerDocument.getSelection()
  if (selection === null) {
    return
  }

  const domRange = root.ownerDocument.createRange()
  domRange.setStart(anchor.node, anchor.offset)
  domRange.setEnd(focus.node, focus.offset)
  selection.removeAllRanges()
  selection.addRange(domRange)
}
