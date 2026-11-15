/**
 * Pure operations on the PDF edit model, plus the undo/redo history around it
 * (PDF-EDIT-1). No React, no pdf.js, no pdf-lib — every function here returns
 * a new state and leaves its input untouched, which is what lets the history
 * keep earlier states by reference.
 */

import { normalizeRotation } from '../rotation'
import type {
  PageSlot,
  PdfAnnotationDraft,
  PdfEditState,
  PdfPoint,
  PdfRect,
  TextEdit,
} from './types'

export function createInitialEditState(pageCount: number): PdfEditState {
  return {
    pages: Array.from({ length: pageCount }, (_, index) => ({ source: index + 1, rotation: 0 })),
    annotations: [],
    textEdits: [],
  }
}

/** Whether the state differs from a freshly loaded document's — i.e. whether writing it changes anything. */
export function hasDocumentEdits(state: PdfEditState, sourcePageCount: number): boolean {
  if (state.annotations.length > 0 || state.textEdits.length > 0) return true
  if (state.pages.length !== sourcePageCount) return true
  return state.pages.some((slot, index) => slot.source !== index + 1 || slot.rotation !== 0)
}

// ---- Pages ------------------------------------------------------------------

export function rotatePage(state: PdfEditState, displayIndex: number, delta: 90 | -90): PdfEditState {
  const slot = state.pages[displayIndex]
  if (!slot) return state
  const pages = state.pages.slice()
  pages[displayIndex] = { ...slot, rotation: normalizeRotation(slot.rotation + delta) }
  return { ...state, pages }
}

/** Moves the page at `displayIndex` by `offset` places; out-of-range moves are a no-op. */
export function movePage(state: PdfEditState, displayIndex: number, offset: number): PdfEditState {
  const target = displayIndex + offset
  if (!state.pages[displayIndex] || target < 0 || target >= state.pages.length) return state
  const pages = state.pages.slice()
  const [slot] = pages.splice(displayIndex, 1)
  pages.splice(target, 0, slot)
  return { ...state, pages }
}

/**
 * Removes a page from the output. The last remaining page cannot be deleted —
 * a PDF with no pages is not a document any reader opens. The page's pending
 * annotations and text edits go with it, so restoring it via undo restores
 * them too (they live in the same state snapshot).
 */
export function deletePage(state: PdfEditState, displayIndex: number): PdfEditState {
  const slot = state.pages[displayIndex]
  if (!slot || state.pages.length <= 1) return state
  return {
    pages: state.pages.filter((_, index) => index !== displayIndex),
    annotations: state.annotations.filter((annotation) => annotation.page !== slot.source),
    textEdits: state.textEdits.filter((edit) => edit.page !== slot.source),
  }
}

export function findSlotIndexForSource(pages: ReadonlyArray<PageSlot>, source: number): number {
  return pages.findIndex((slot) => slot.source === source)
}

// ---- Annotations --------------------------------------------------------------

export function addAnnotation(state: PdfEditState, annotation: PdfAnnotationDraft): PdfEditState {
  return { ...state, annotations: [...state.annotations, annotation] }
}

export function updateAnnotation(state: PdfEditState, annotation: PdfAnnotationDraft): PdfEditState {
  const index = state.annotations.findIndex((existing) => existing.id === annotation.id)
  if (index === -1) return state
  const annotations = state.annotations.slice()
  annotations[index] = annotation
  return { ...state, annotations }
}

/** Removes an annotation or a text replacement by id — the select tool deletes either. */
export function removeEdit(state: PdfEditState, id: string): PdfEditState {
  const annotations = state.annotations.filter((annotation) => annotation.id !== id)
  const textEdits = state.textEdits.filter((edit) => edit.id !== id)
  if (annotations.length === state.annotations.length && textEdits.length === state.textEdits.length) {
    return state
  }
  return { ...state, annotations, textEdits }
}

function translatePoint(point: PdfPoint, dx: number, dy: number): PdfPoint {
  return { x: point.x + dx, y: point.y + dy }
}

function translateRect(rect: PdfRect, dx: number, dy: number): PdfRect {
  return { ...rect, x: rect.x + dx, y: rect.y + dy }
}

export function translateAnnotation(annotation: PdfAnnotationDraft, dx: number, dy: number): PdfAnnotationDraft {
  switch (annotation.kind) {
    case 'text':
      return { ...annotation, anchor: translatePoint(annotation.anchor, dx, dy) }
    case 'highlight':
      return { ...annotation, rects: annotation.rects.map((rect) => translateRect(rect, dx, dy)) }
    case 'ink':
      return { ...annotation, points: annotation.points.map((point) => translatePoint(point, dx, dy)) }
  }
}

// ---- Text edits ------------------------------------------------------------------

function sameTarget(a: TextEdit, b: TextEdit): boolean {
  if (a.page !== b.page) return false
  return a.paragraphId !== null ? a.paragraphId === b.paragraphId : a.id === b.id
}

/**
 * Adds or replaces the edit of one paragraph (or one piece of new text).
 * `unchanged` drops it instead: editing a paragraph back to exactly what the
 * file says leaves the file's own bytes alone rather than rewriting them into
 * an identical copy.
 */
export function setTextEdit(state: PdfEditState, edit: TextEdit, unchanged: boolean): PdfEditState {
  const others = state.textEdits.filter((existing) => !sameTarget(existing, edit))
  return { ...state, textEdits: unchanged ? others : [...others, edit] }
}

export function findTextEdit(state: PdfEditState, page: number, paragraphId: string): TextEdit | undefined {
  return state.textEdits.find((edit) => edit.page === page && edit.paragraphId === paragraphId)
}

// ---- History ------------------------------------------------------------------------

/** Generous for a document editor, and bounded: each entry is a shallow-shared snapshot. */
export const MAX_HISTORY = 200

export type EditHistory = {
  readonly past: ReadonlyArray<PdfEditState>
  readonly present: PdfEditState
  readonly future: ReadonlyArray<PdfEditState>
}

/**
 * Dirty is an identity compare against the state that was last written. States
 * are never mutated and undo/redo hand back the very objects they stored, so
 * undoing back to the saved state reads as clean again — while "undo, then make
 * a different edit" does not, which a revision counter would get wrong (both
 * land on the same count).
 */
export function isEditStateDirty(history: EditHistory, savedState: PdfEditState): boolean {
  return history.present !== savedState
}

export function createHistory(initial: PdfEditState): EditHistory {
  return { past: [], present: initial, future: [] }
}

export function pushHistory(history: EditHistory, next: PdfEditState): EditHistory {
  if (next === history.present) return history
  const past = [...history.past, history.present]
  return {
    past: past.length > MAX_HISTORY ? past.slice(past.length - MAX_HISTORY) : past,
    present: next,
    future: [],
  }
}

export function undoHistory(history: EditHistory): EditHistory {
  const previous = history.past[history.past.length - 1]
  if (!previous) return history
  return {
    past: history.past.slice(0, -1),
    present: previous,
    future: [history.present, ...history.future],
  }
}

export function redoHistory(history: EditHistory): EditHistory {
  const [next, ...future] = history.future
  if (!next) return history
  return {
    past: [...history.past, history.present],
    present: next,
    future,
  }
}
