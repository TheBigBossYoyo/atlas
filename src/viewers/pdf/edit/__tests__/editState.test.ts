import { describe, expect, it } from 'vitest'

import {
  addAnnotation,
  createHistory,
  createInitialEditState,
  deletePage,
  hasDocumentEdits,
  isEditStateDirty,
  MAX_HISTORY,
  movePage,
  pushHistory,
  redoHistory,
  removeEdit,
  rotatePage,
  translateAnnotation,
  undoHistory,
  setTextEdit,
} from '../editState'
import type { InkAnnotation, TextEdit } from '../types'

const ink = (id: string, page: number): InkAnnotation => ({
  kind: 'ink',
  id,
  page,
  points: [{ x: 1, y: 2 }],
  color: { r: 0, g: 0, b: 0 },
  width: 1,
})

const textEdit = (id: string, paragraphId: string | null, text: string): TextEdit => ({
  kind: 'text-edit',
  id,
  page: 1,
  paragraphId,
  streamKey: 'page',
  pieces: [{ opIndex: 3, element: 0 }],
  insertAfterOp: 4,
  origin: { x: 0, y: 0 },
  direction: { x: 1, y: 0 },
  leading: 12,
  width: null,
  align: 'left',
  firstLineIndent: 0,
  tabStops: [],
  styles: [],
  spans: [{ text, style: 0 }],
  bounds: [],
})

describe('page operations', () => {
  it('a fresh state has no edits', () => {
    expect(hasDocumentEdits(createInitialEditState(3), 3)).toBe(false)
  })

  it('rotation accumulates and wraps, and turning back to 0 is no edit at all', () => {
    let state = rotatePage(createInitialEditState(2), 1, -90)
    expect(state.pages[1].rotation).toBe(270)
    state = rotatePage(state, 1, 90)
    expect(state.pages[1].rotation).toBe(0)
    expect(hasDocumentEdits(state, 2)).toBe(false)
  })

  it('moves pages and ignores moves off either end', () => {
    const state = createInitialEditState(3)
    expect(movePage(state, 0, 2).pages.map((slot) => slot.source)).toEqual([2, 3, 1])
    expect(movePage(state, 0, -1)).toBe(state)
    expect(movePage(state, 2, 1)).toBe(state)
  })

  it('deleting a page drops its pending edits with it, and the last page cannot go', () => {
    let state = addAnnotation(createInitialEditState(2), ink('a', 1))
    state = addAnnotation(state, ink('b', 2))
    state = deletePage(state, 0)
    expect(state.pages.map((slot) => slot.source)).toEqual([2])
    expect(state.annotations.map((annotation) => annotation.id)).toEqual(['b'])
    expect(deletePage(state, 0)).toBe(state)
  })
})

describe('annotations and text replacements', () => {
  it('removeEdit finds either kind by id and is a no-op for an unknown one', () => {
    let state = addAnnotation(createInitialEditState(1), ink('a', 1))
    state = setTextEdit(state, textEdit('r', 'page:3:0', 'after'), false)
    expect(removeEdit(state, 'missing')).toBe(state)
    expect(removeEdit(removeEdit(state, 'a'), 'r')).toEqual(createInitialEditState(1))
  })

  it('editing a paragraph back to what the file says drops its edit', () => {
    const edited = setTextEdit(createInitialEditState(1), textEdit('r', 'page:3:0', 'after'), false)
    expect(edited.textEdits).toHaveLength(1)
    expect(setTextEdit(edited, textEdit('r', 'page:3:0', 'before'), true).textEdits).toHaveLength(0)
  })

  it('a second edit of the same paragraph replaces the first instead of stacking', () => {
    let state = setTextEdit(createInitialEditState(1), textEdit('a', 'page:3:0', 'one'), false)
    state = setTextEdit(state, textEdit('b', 'page:3:0', 'two'), false)
    expect(state.textEdits.map((edit) => edit.spans[0].text)).toEqual(['two'])
  })

  it('new text (no paragraph) is identified by its own id', () => {
    let state = setTextEdit(createInitialEditState(1), textEdit('n1', null, 'first'), false)
    state = setTextEdit(state, textEdit('n2', null, 'second'), false)
    expect(state.textEdits).toHaveLength(2)
    state = setTextEdit(state, textEdit('n1', null, 'first, revised'), false)
    expect(state.textEdits.map((edit) => edit.spans[0].text).sort()).toEqual(['first, revised', 'second'])
  })

  it('deleting a page drops its text edits too', () => {
    const state = setTextEdit(createInitialEditState(2), textEdit('r', 'page:3:0', 'after'), false)
    expect(deletePage(state, 0).textEdits).toHaveLength(0)
  })

  it('translates every kind of annotation', () => {
    expect(translateAnnotation(ink('a', 1), 5, -1)).toMatchObject({ points: [{ x: 6, y: 1 }] })
    expect(
      translateAnnotation(
        { kind: 'highlight', id: 'h', page: 1, rects: [{ x: 0, y: 0, width: 2, height: 2 }], color: { r: 1, g: 1, b: 0 } },
        1,
        1,
      ),
    ).toMatchObject({ rects: [{ x: 1, y: 1 }] })
  })
})

describe('history', () => {
  it('undo back to the saved state reads clean; undo then a different edit does not', () => {
    const initial = createInitialEditState(2)
    let history = createHistory(initial)
    history = pushHistory(history, rotatePage(history.present, 0, 90))
    expect(isEditStateDirty(history, initial)).toBe(true)
    history = undoHistory(history)
    expect(isEditStateDirty(history, initial)).toBe(false)
    // Same number of steps from the saved state as before, different content.
    history = pushHistory(history, rotatePage(history.present, 1, 90))
    expect(isEditStateDirty(history, initial)).toBe(true)
  })

  it('redo replays, and a new edit clears the redo stack', () => {
    let history = createHistory(createInitialEditState(1))
    const rotated = rotatePage(history.present, 0, 90)
    history = pushHistory(history, rotated)
    history = undoHistory(history)
    expect(redoHistory(history).present).toBe(rotated)
    history = pushHistory(history, addAnnotation(history.present, ink('a', 1)))
    expect(history.future).toHaveLength(0)
    expect(redoHistory(history)).toBe(history)
  })

  it('pushing an unchanged state is not an undo step', () => {
    const history = createHistory(createInitialEditState(1))
    expect(pushHistory(history, history.present)).toBe(history)
  })

  it('keeps at most MAX_HISTORY undo steps', () => {
    let history = createHistory(createInitialEditState(1))
    for (let index = 0; index < MAX_HISTORY + 10; index += 1) {
      history = pushHistory(history, addAnnotation(history.present, ink(`a${index}`, 1)))
    }
    expect(history.past).toHaveLength(MAX_HISTORY)
  })
})
