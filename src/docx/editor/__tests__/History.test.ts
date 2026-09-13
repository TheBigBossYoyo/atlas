import { afterEach, describe, expect, it, vi } from 'vitest'

import { applyCommand, History } from '../index'
import type { Position } from '../index'
import type {
  Comment,
  Document,
  Endnote,
  Footer,
  Footnote,
  Header,
  NumberingDef,
  Paragraph,
  Run,
  Section,
  Style,
} from '../../model'

describe('History', () => {
  afterEach(() => {
    vi.useRealTimers()
  })

  it('returns null when undo is empty', () => {
    const history = new History()
    const document = createDocument('Atlas')

    expect(history.undo(document)).toBeNull()
  })

  it('returns null when redo is empty', () => {
    const history = new History()
    const document = createDocument('Atlas')

    expect(history.redo(document)).toBeNull()
  })

  it('undo applies the stored inverse and returns a redo command that replays it', () => {
    const history = new History()
    const original = createDocument('Atlas')
    const command = { kind: 'insert-text', at: position([0], 0, 5), text: '!' } as const
    const applied = applyCommand(original, command)

    history.push(applied.inverse)
    const undone = history.undo(applied.document)

    expect(undone?.document).toEqual(original)

    // DXE-02/DXE-04 — delete's own inverse is now the general replace-blocks
    // primitive (needed so cross-paragraph/cross-run delete also inverts
    // exactly), so the redo command is no longer byte-identical to the
    // original insert-text — what matters is that replaying it reproduces the
    // post-insert document.
    const redone = applyCommand(undone!.document, undone!.redoCommand)
    expect(redone.document).toEqual(applied.document)
  })

  it('redo reapplies the last undone command', () => {
    const history = new History()
    const original = createDocument('Atlas')
    const command = { kind: 'insert-text', at: position([0], 0, 5), text: '!' } as const
    const applied = applyCommand(original, command)

    history.push(applied.inverse)
    const undone = history.undo(applied.document)
    const redone = history.redo(undone!.document)

    expect(redone?.document).toEqual(applied.document)

    const undoneAgain = applyCommand(redone!.document, redone!.undoCommand)
    expect(undoneAgain.document).toEqual(original)
  })

  it('undo restores the selection active before the edit it undoes (DXE-16)', () => {
    const history = new History()
    const original = createDocument('Atlas')
    const command = { kind: 'insert-text', at: position([0], 0, 5), text: '!' } as const
    const applied = applyCommand(original, command)

    history.push(applied.inverse)
    const undone = history.undo(applied.document)

    // Undoing an insert collapses the cursor back to where the insertion
    // started — exactly where the user was typing from.
    expect(undone?.range).toEqual({
      anchor: position([0], 0, 5),
      focus: position([0], 0, 5),
    })
  })

  it('redo restores the selection at the end of the redone edit', () => {
    const history = new History()
    const original = createDocument('Atlas')
    const command = { kind: 'insert-text', at: position([0], 0, 5), text: '!' } as const
    const applied = applyCommand(original, command)

    history.push(applied.inverse)
    const undone = history.undo(applied.document)
    const redone = history.redo(undone!.document)

    // The redo command is the (undone insert's) replace-blocks inverse,
    // whose `cursor` re-selects exactly the span it is restoring — arguably
    // better UX than a bare collapsed cursor (it mirrors how undoing a
    // delete re-selects the just-restored text) and still satisfies "the
    // cursor moves to where that edit was".
    expect(redone?.range).toEqual({
      anchor: position([0], 0, 5),
      focus: position([0], 0, 6),
    })
  })

  it('push clears the redo stack', () => {
    const history = new History()
    const original = createDocument('Atlas')
    const insert = applyCommand(original, {
      kind: 'insert-text',
      at: position([0], 0, 5),
      text: '!',
    })

    history.push(insert.inverse)
    const undone = history.undo(insert.document)
    expect(undone).not.toBeNull()

    history.push({
      kind: 'apply-style',
      paragraphPath: [0],
      styleId: 'Heading1',
    })

    expect(history.redo(undone!.document)).toBeNull()
  })

  it('coalesces adjacent InsertText commands within one second', () => {
    vi.useFakeTimers()
    vi.setSystemTime(new Date('2026-05-15T12:00:00Z'))

    const history = new History()
    const original = createDocument('')
    const firstCommand = { kind: 'insert-text', at: position([0], 0, 0), text: 'A' } as const
    const first = applyCommand(original, firstCommand)

    expect(history.coalesceWithLast(firstCommand, first.inverse)).toBe(false)
    history.push(first.inverse)

    vi.advanceTimersByTime(300)

    const secondCommand = { kind: 'insert-text', at: position([0], 0, 1), text: 't' } as const
    const second = applyCommand(first.document, secondCommand)

    expect(history.coalesceWithLast(secondCommand, second.inverse)).toBe(true)

    const undone = history.undo(second.document)
    expect(undone?.document).toEqual(original)
    // The coalesced undo entry is still the rewritten delete-range described
    // above; only its own inverse (what redo replays) now comes back as a
    // replace-blocks snapshot rather than a plain insert-text — replaying it
    // must still reproduce "At".
    const redone = applyCommand(undone!.document, undone!.redoCommand)
    expect(runText(redone.document)).toBe('At')
  })

  it('does not coalesce non-text commands', () => {
    vi.useFakeTimers()
    vi.setSystemTime(new Date('2026-05-15T12:00:00Z'))

    const history = new History()
    const original = createDocument('Atlas')
    const firstCommand = { kind: 'insert-text', at: position([0], 0, 5), text: '!' } as const
    const first = applyCommand(original, firstCommand)

    expect(history.coalesceWithLast(firstCommand, first.inverse)).toBe(false)
    history.push(first.inverse)

    vi.advanceTimersByTime(100)

    const styleCommand = {
      kind: 'apply-style',
      paragraphPath: [0],
      styleId: 'Heading1',
    } as const
    const style = applyCommand(first.document, styleCommand)

    expect(history.coalesceWithLast(styleCommand, style.inverse)).toBe(false)
  })

  it('does not coalesce distant InsertText commands after the window expires', () => {
    vi.useFakeTimers()
    vi.setSystemTime(new Date('2026-05-15T12:00:00Z'))

    const history = new History()
    const original = createDocument('')
    const firstCommand = { kind: 'insert-text', at: position([0], 0, 0), text: 'A' } as const
    const first = applyCommand(original, firstCommand)

    expect(history.coalesceWithLast(firstCommand, first.inverse)).toBe(false)
    history.push(first.inverse)

    vi.advanceTimersByTime(1200)

    const secondCommand = { kind: 'insert-text', at: position([0], 0, 1), text: 't' } as const
    const second = applyCommand(first.document, secondCommand)

    expect(history.coalesceWithLast(secondCommand, second.inverse)).toBe(false)
    history.push(second.inverse)

    const undone = history.undo(second.document)
    expect(runText(undone!.document)).toBe('A')
  })

  it('does not coalesce InsertText commands at non-adjacent positions', () => {
    vi.useFakeTimers()
    vi.setSystemTime(new Date('2026-05-15T12:00:00Z'))

    const history = new History()
    const original = createDocument('AB')
    const firstCommand = { kind: 'insert-text', at: position([0], 0, 1), text: 'x' } as const
    const first = applyCommand(original, firstCommand)

    expect(history.coalesceWithLast(firstCommand, first.inverse)).toBe(false)
    history.push(first.inverse)

    vi.advanceTimersByTime(300)

    const secondCommand = { kind: 'insert-text', at: position([0], 0, 0), text: 'y' } as const
    const second = applyCommand(first.document, secondCommand)

    expect(history.coalesceWithLast(secondCommand, second.inverse)).toBe(false)
    history.push(second.inverse)

    const undone = history.undo(second.document)
    expect(runText(undone!.document)).toBe('AxB')
  })
})

function position(paragraphPath: ReadonlyArray<number>, runIndex: number, charOffset: number): Position {
  return {
    paragraphPath: Object.freeze([...paragraphPath]),
    runIndex,
    charOffset,
  }
}

function runText(document: Document): string {
  const paragraph = document.sections[0].blocks[0] as Paragraph
  const run = paragraph.children[0] as Run
  return run.children.map((child) => child.kind === 'text' ? child.value : '').join('')
}

function createDocument(text: string): Document {
  const paragraph = Object.freeze({
    kind: 'paragraph',
    children: Object.freeze([
      Object.freeze({
        kind: 'run',
        children: Object.freeze([
          Object.freeze({
            kind: 'text',
            value: text,
          }),
        ]),
      }),
    ]),
  }) satisfies Paragraph

  const section = Object.freeze({
    kind: 'section',
    props: {},
    blocks: Object.freeze([paragraph]),
  }) satisfies Section

  return Object.freeze({
    kind: 'document',
    sections: Object.freeze([section]),
    styles: new Map<string, Style>(),
    numbering: new Map<string, NumberingDef>(),
    comments: new Map<string, Comment>(),
    footnotes: new Map<string, Footnote>(),
    endnotes: new Map<string, Endnote>(),
    headers: new Map<string, Header>(),
    footers: new Map<string, Footer>(),
  }) satisfies Document
}
