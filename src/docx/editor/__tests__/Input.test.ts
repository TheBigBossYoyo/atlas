import { describe, expect, it } from 'vitest'

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
  RunProps,
  Section,
  Style,
} from '../../model'
import type { Position, Range } from '../commandTypes'
import { History } from '../History'
import {
  extendOrCollapse,
  handleBeforeInput,
  handleKeyDown,
  moveCursorLeft,
  moveCursorRight,
  moveCursorToDocEnd,
  moveCursorToDocStart,
  moveCursorToLineEnd,
  moveCursorToLineStart,
} from '../Input'

// ─── Helpers ──────────────────────────────────────────────────────────────────

function pos(paragraphPath: ReadonlyArray<number>, runIndex: number, charOffset: number): Position {
  return {
    paragraphPath: Object.freeze([...paragraphPath]),
    runIndex,
    charOffset,
  }
}

function range(anchor: Position, focus: Position): Range {
  return { anchor, focus }
}

function collapsedRange(p: Position): Range {
  return { anchor: p, focus: p }
}

function createDocument(paragraphs: ReadonlyArray<Paragraph>): Document {
  const section = Object.freeze({
    kind: 'section',
    props: {},
    blocks: Object.freeze([...paragraphs]),
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

function createParagraph(texts: ReadonlyArray<string>, runProps?: ReadonlyArray<RunProps | undefined>): Paragraph {
  const children = texts.map((text, i) => createRun(text, runProps?.[i]))
  return Object.freeze({
    kind: 'paragraph',
    children: Object.freeze(children),
  }) satisfies Paragraph
}

function createRun(text: string, props?: RunProps): Run {
  return Object.freeze({
    kind: 'run',
    ...(props !== undefined ? { props } : {}),
    children: Object.freeze([Object.freeze({ kind: 'text', value: text })]),
  }) satisfies Run
}

function getText(doc: Document): string[][] {
  return doc.sections[0].blocks.map((b) => {
    const para = b as Paragraph
    return para.children.map((c) => {
      const run = c as Run
      return run.children.map((ch) => (ch.kind === 'text' ? ch.value : '')).join('')
    })
  })
}

function makeInputEvent(inputType: string, data?: string | null): InputEvent {
  return new InputEvent('beforeinput', { inputType, data: data ?? null })
}

function makeKeyEvent(key: string, opts?: { ctrl?: boolean; shift?: boolean; meta?: boolean }): KeyboardEvent {
  return new KeyboardEvent('keydown', {
    key,
    ctrlKey: opts?.ctrl ?? false,
    shiftKey: opts?.shift ?? false,
    metaKey: opts?.meta ?? false,
  })
}

// ─── moveCursorLeft ───────────────────────────────────────────────────────────

describe('moveCursorLeft', () => {
  it('moves left within the same run', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const result = moveCursorLeft(pos([0, 0], 0, 3), doc)
    expect(result).toEqual(pos([0, 0], 0, 2))
  })

  it('clamps at run start (offset 0) by jumping to prev run end', () => {
    const doc = createDocument([createParagraph(['ab', 'cd'])])
    const result = moveCursorLeft(pos([0, 0], 1, 0), doc)
    expect(result).toEqual(pos([0, 0], 0, 2))
  })

  it('jumps to end of previous paragraph', () => {
    const doc = createDocument([createParagraph(['hello']), createParagraph(['world'])])
    const result = moveCursorLeft(pos([0, 1], 0, 0), doc)
    expect(result).toEqual(pos([0, 0], 0, 5))
  })

  it('clamps at document start', () => {
    const doc = createDocument([createParagraph(['hi'])])
    const startPos = pos([0, 0], 0, 0)
    const result = moveCursorLeft(startPos, doc)
    expect(result).toEqual(startPos)
  })

  it('decrements charOffset by 1 normally', () => {
    const doc = createDocument([createParagraph(['abc'])])
    const result = moveCursorLeft(pos([0, 0], 0, 2), doc)
    expect(result.charOffset).toBe(1)
  })
})

// ─── moveCursorRight ──────────────────────────────────────────────────────────

describe('moveCursorRight', () => {
  it('moves right within the same run', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const result = moveCursorRight(pos([0, 0], 0, 2), doc)
    expect(result).toEqual(pos([0, 0], 0, 3))
  })

  it('wraps to next run in same paragraph', () => {
    const doc = createDocument([createParagraph(['ab', 'cd'])])
    const result = moveCursorRight(pos([0, 0], 0, 2), doc)
    expect(result).toEqual(pos([0, 0], 1, 0))
  })

  it('jumps to start of next paragraph', () => {
    const doc = createDocument([createParagraph(['hello']), createParagraph(['world'])])
    const result = moveCursorRight(pos([0, 0], 0, 5), doc)
    expect(result).toEqual(pos([0, 1], 0, 0))
  })

  it('clamps at document end', () => {
    const doc = createDocument([createParagraph(['hi'])])
    const endPos = pos([0, 0], 0, 2)
    const result = moveCursorRight(endPos, doc)
    expect(result).toEqual(endPos)
  })

  it('increments charOffset by 1 normally', () => {
    const doc = createDocument([createParagraph(['abc'])])
    const result = moveCursorRight(pos([0, 0], 0, 1), doc)
    expect(result.charOffset).toBe(2)
  })
})

// ─── moveCursorToLineStart / moveCursorToLineEnd ──────────────────────────────

describe('moveCursorToLineStart', () => {
  it('returns run 0 offset 0 for any position', () => {
    const doc = createDocument([createParagraph(['hello world'])])
    const result = moveCursorToLineStart(pos([0, 0], 0, 5), doc)
    expect(result).toEqual(pos([0, 0], 0, 0))
  })
})

describe('moveCursorToLineEnd', () => {
  it('returns position after last character in paragraph', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const result = moveCursorToLineEnd(pos([0, 0], 0, 0), doc)
    expect(result).toEqual(pos([0, 0], 0, 5))
  })

  it('returns end of last run in multi-run paragraph', () => {
    const doc = createDocument([createParagraph(['ab', 'cde'])])
    const result = moveCursorToLineEnd(pos([0, 0], 0, 0), doc)
    expect(result).toEqual(pos([0, 0], 1, 3))
  })
})

// ─── moveCursorToDocStart / moveCursorToDocEnd ────────────────────────────────

describe('moveCursorToDocStart', () => {
  it('returns position at start of first paragraph', () => {
    const doc = createDocument([createParagraph(['hello']), createParagraph(['world'])])
    const result = moveCursorToDocStart(doc)
    expect(result.runIndex).toBe(0)
    expect(result.charOffset).toBe(0)
  })
})

describe('moveCursorToDocEnd', () => {
  it('returns position at end of last paragraph', () => {
    const doc = createDocument([createParagraph(['hello']), createParagraph(['world'])])
    const result = moveCursorToDocEnd(doc)
    expect(result.charOffset).toBe(5)
  })
})

// ─── extendOrCollapse ────────────────────────────────────────────────────────

describe('extendOrCollapse', () => {
  it('collapses when shift not held', () => {
    const current = range(pos([0, 0], 0, 0), pos([0, 0], 0, 3))
    const newFocus = pos([0, 0], 0, 5)
    const result = extendOrCollapse(current, newFocus, false)
    expect(result.anchor).toEqual(newFocus)
    expect(result.focus).toEqual(newFocus)
  })

  it('extends when shift is held', () => {
    const anchor = pos([0, 0], 0, 0)
    const current = range(anchor, pos([0, 0], 0, 3))
    const newFocus = pos([0, 0], 0, 7)
    const result = extendOrCollapse(current, newFocus, true)
    expect(result.anchor).toEqual(anchor)
    expect(result.focus).toEqual(newFocus)
  })

  it('creates collapsed range when current is null and shift not held', () => {
    const newFocus = pos([0, 0], 0, 2)
    const result = extendOrCollapse(null, newFocus, false)
    expect(result.anchor).toEqual(newFocus)
    expect(result.focus).toEqual(newFocus)
  })
})

// ─── handleBeforeInput ───────────────────────────────────────────────────────

describe('handleBeforeInput', () => {
  it('insertText appends text and updates range', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const history = new History()
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 5)), history }
    const result = handleBeforeInput(makeInputEvent('insertText', ' world'), ctx)
    expect(result).not.toBeNull()
    expect(getText(result!.document)).toEqual([['hello world']])
    expect(result!.range?.focus.charOffset).toBe(11)
  })

  it('insertText returns null when range is null', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const ctx = { document: doc, range: null, history: new History() }
    const result = handleBeforeInput(makeInputEvent('insertText', 'x'), ctx)
    expect(result).toBeNull()
  })

  it('insertParagraph splits paragraph', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const history = new History()
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 2)), history }
    const result = handleBeforeInput(makeInputEvent('insertParagraph'), ctx)
    expect(result).not.toBeNull()
    expect(result!.document.sections[0].blocks.length).toBe(2)
  })

  it('deleteContentBackward removes character before cursor', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const history = new History()
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 3)), history }
    const result = handleBeforeInput(makeInputEvent('deleteContentBackward'), ctx)
    expect(result).not.toBeNull()
    expect(getText(result!.document)).toEqual([['helo']])
    expect(result!.range?.focus.charOffset).toBe(2)
  })

  it('deleteContentForward removes character after cursor', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const history = new History()
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 2)), history }
    const result = handleBeforeInput(makeInputEvent('deleteContentForward'), ctx)
    expect(result).not.toBeNull()
    expect(getText(result!.document)).toEqual([['helo']])
  })

  it('unknown inputType returns null', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 0)), history: new History() }
    const result = handleBeforeInput(makeInputEvent('insertOrderedList'), ctx)
    expect(result).toBeNull()
  })
})

// ─── handleKeyDown ────────────────────────────────────────────────────────────

describe('handleKeyDown', () => {
  it('Ctrl+Z undoes last insert', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const history = new History()
    const ctx1 = { document: doc, range: collapsedRange(pos([0, 0], 0, 5)), history }
    const afterInsert = handleBeforeInput(makeInputEvent('insertText', ' world'), ctx1)!
    const ctx2 = { document: afterInsert.document, range: afterInsert.range, history }
    const result = handleKeyDown(makeKeyEvent('z', { ctrl: true }), ctx2)
    expect(result).not.toBeNull()
    expect(getText(result!.document)).toEqual([['hello']])
  })

  it('Ctrl+Y redoes after undo', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const history = new History()
    const ctx1 = { document: doc, range: collapsedRange(pos([0, 0], 0, 5)), history }
    const afterInsert = handleBeforeInput(makeInputEvent('insertText', ' world'), ctx1)!
    const ctx2 = { document: afterInsert.document, range: afterInsert.range, history }
    const afterUndo = handleKeyDown(makeKeyEvent('z', { ctrl: true }), ctx2)!
    const ctx3 = { document: afterUndo.document, range: afterUndo.range, history }
    const result = handleKeyDown(makeKeyEvent('y', { ctrl: true }), ctx3)
    expect(result).not.toBeNull()
    expect(getText(result!.document)).toEqual([['hello world']])
  })

  it('Ctrl+Shift+Z redoes', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const history = new History()
    const ctx1 = { document: doc, range: collapsedRange(pos([0, 0], 0, 5)), history }
    const afterInsert = handleBeforeInput(makeInputEvent('insertText', '!'), ctx1)!
    const ctx2 = { document: afterInsert.document, range: afterInsert.range, history }
    const afterUndo = handleKeyDown(makeKeyEvent('z', { ctrl: true }), ctx2)!
    const ctx3 = { document: afterUndo.document, range: afterUndo.range, history }
    const result = handleKeyDown(makeKeyEvent('z', { ctrl: true, shift: true }), ctx3)
    expect(result).not.toBeNull()
    expect(getText(result!.document)).toEqual([['hello!']])
  })

  it('Ctrl+B emits ApplyRunFormat with bold toggle and returns updated document', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const history = new History()
    const sel = range(pos([0, 0], 0, 0), pos([0, 0], 0, 5))
    const ctx = { document: doc, range: sel, history }
    const result = handleKeyDown(makeKeyEvent('b', { ctrl: true }), ctx)
    expect(result).not.toBeNull()
    // After applying bold, the run should have bold: true
    const run = (result!.document.sections[0].blocks[0] as Paragraph).children[0] as Run
    expect(run.props?.bold).toBe(true)
  })

  it('Ctrl+B with no selection returns null', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 2)), history: new History() }
    const result = handleKeyDown(makeKeyEvent('b', { ctrl: true }), ctx)
    expect(result).toBeNull()
  })

  it('ArrowLeft moves cursor left', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 3)), history: new History() }
    const result = handleKeyDown(makeKeyEvent('ArrowLeft'), ctx)
    expect(result).not.toBeNull()
    expect(result!.range?.focus.charOffset).toBe(2)
  })

  it('ArrowRight moves cursor right', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 2)), history: new History() }
    const result = handleKeyDown(makeKeyEvent('ArrowRight'), ctx)
    expect(result).not.toBeNull()
    expect(result!.range?.focus.charOffset).toBe(3)
  })

  it('Shift+ArrowRight extends selection', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const initial = collapsedRange(pos([0, 0], 0, 2))
    const ctx = { document: doc, range: initial, history: new History() }
    const result = handleKeyDown(makeKeyEvent('ArrowRight', { shift: true }), ctx)
    expect(result).not.toBeNull()
    expect(result!.range?.anchor.charOffset).toBe(2)
    expect(result!.range?.focus.charOffset).toBe(3)
  })

  it('Home moves cursor to line start', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 3)), history: new History() }
    const result = handleKeyDown(makeKeyEvent('Home'), ctx)
    expect(result).not.toBeNull()
    expect(result!.range?.focus.charOffset).toBe(0)
  })

  it('End moves cursor to line end', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 0)), history: new History() }
    const result = handleKeyDown(makeKeyEvent('End'), ctx)
    expect(result).not.toBeNull()
    expect(result!.range?.focus.charOffset).toBe(5)
  })

  it('Ctrl+A selects all content', () => {
    const doc = createDocument([createParagraph(['hello']), createParagraph(['world'])])
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 0)), history: new History() }
    const result = handleKeyDown(makeKeyEvent('a', { ctrl: true }), ctx)
    expect(result).not.toBeNull()
    const r = result!.range!
    expect(r.anchor.charOffset).toBe(0)
    expect(r.focus.charOffset).toBe(5)
  })

  it('unhandled key returns null', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 0)), history: new History() }
    const result = handleKeyDown(makeKeyEvent('F5'), ctx)
    expect(result).toBeNull()
  })

  it('ArrowUp returns null (deferred)', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 2)), history: new History() }
    expect(handleKeyDown(makeKeyEvent('ArrowUp'), ctx)).toBeNull()
  })

  it('ArrowDown returns null (deferred)', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 2)), history: new History() }
    expect(handleKeyDown(makeKeyEvent('ArrowDown'), ctx)).toBeNull()
  })

  it('Backspace deletes character before cursor', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const history = new History()
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 3)), history }
    const result = handleKeyDown(makeKeyEvent('Backspace'), ctx)
    expect(result).not.toBeNull()
    expect(getText(result!.document)).toEqual([['helo']])
  })

  it('Enter inserts paragraph break', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const history = new History()
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 2)), history }
    const result = handleKeyDown(makeKeyEvent('Enter'), ctx)
    expect(result).not.toBeNull()
    expect(result!.document.sections[0].blocks.length).toBe(2)
  })

  // ─── D18 — Tab/Shift+Tab list level ──────────────────────────────────────

  it('Tab at the start of a list paragraph increases its level instead of inserting a tab', () => {
    const listParagraph: Paragraph = Object.freeze({
      kind: 'paragraph',
      props: { numPr: { numId: '1', ilvl: 0 } },
      children: Object.freeze([createRun('Item')]),
    }) as Paragraph
    const doc = createDocument([listParagraph])
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 0)), history: new History() }

    const result = handleKeyDown(makeKeyEvent('Tab'), ctx)

    expect(result).not.toBeNull()
    const updated = result!.document.sections[0].blocks[0] as Paragraph
    expect(updated.props?.numPr).toEqual({ numId: '1', ilvl: 1 })
  })

  it('Shift+Tab at the start of a list paragraph decreases its level', () => {
    const listParagraph: Paragraph = Object.freeze({
      kind: 'paragraph',
      props: { numPr: { numId: '1', ilvl: 1 } },
      children: Object.freeze([createRun('Item')]),
    }) as Paragraph
    const doc = createDocument([listParagraph])
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 0)), history: new History() }

    const result = handleKeyDown(makeKeyEvent('Tab', { shift: true }), ctx)

    expect(result).not.toBeNull()
    const updated = result!.document.sections[0].blocks[0] as Paragraph
    expect(updated.props?.numPr).toEqual({ numId: '1', ilvl: 0 })
  })

  it('Tab still inserts a literal tab character outside a list', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const history = new History()
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 0)), history }

    const result = handleKeyDown(makeKeyEvent('Tab'), ctx)

    expect(result).not.toBeNull()
    expect(getText(result!.document)).toEqual([['\thello']])
  })

  it('Tab mid-paragraph in a list still inserts a literal tab (only line-start Tab changes level)', () => {
    const listParagraph: Paragraph = Object.freeze({
      kind: 'paragraph',
      props: { numPr: { numId: '1', ilvl: 0 } },
      children: Object.freeze([createRun('Item')]),
    }) as Paragraph
    const doc = createDocument([listParagraph])
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 2)), history: new History() }

    const result = handleKeyDown(makeKeyEvent('Tab'), ctx)

    expect(result).not.toBeNull()
    expect(getText(result!.document)).toEqual([['It\tem']])
  })

  it('Shift+Tab outside a list is a no-op', () => {
    const doc = createDocument([createParagraph(['hello'])])
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 0)), history: new History() }

    expect(handleKeyDown(makeKeyEvent('Tab', { shift: true }), ctx)).toBeNull()
  })
})

// ─── DXE-21 — word-boundary delete crosses runs and paragraphs ──────────────

describe('word-boundary delete (DXE-21)', () => {
  it('Ctrl+Backspace deletes a word that spans a run boundary', () => {
    const doc = createDocument([createParagraph(['foo', 'bar baz'])])
    const history = new History()
    // Cursor right after "foobar" (run0="foo", run1="bar baz" — offset 3
    // within run1, i.e. right after "bar").
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 1, 3)), history }

    const result = handleBeforeInput(makeInputEvent('deleteWordBackward'), ctx)

    expect(result).not.toBeNull()
    expect(getText(result!.document).flat().join('')).toBe(' baz')
  })

  it('Ctrl+Backspace at the start of a paragraph merges into the previous paragraph', () => {
    const doc = createDocument([createParagraph(['hello']), createParagraph(['world'])])
    const history = new History()
    const ctx = { document: doc, range: collapsedRange(pos([0, 1], 0, 0)), history }

    const result = handleBeforeInput(makeInputEvent('deleteWordBackward'), ctx)

    expect(result).not.toBeNull()
    expect(result!.document.sections[0].blocks.length).toBe(1)
    expect(getText(result!.document)).toEqual([['helloworld']])
  })

  it('Ctrl+Delete at the end of a paragraph merges the next paragraph in', () => {
    const doc = createDocument([createParagraph(['hello']), createParagraph(['world'])])
    const history = new History()
    const ctx = { document: doc, range: collapsedRange(pos([0, 0], 0, 5)), history }

    const result = handleBeforeInput(makeInputEvent('deleteWordForward'), ctx)

    expect(result).not.toBeNull()
    expect(result!.document.sections[0].blocks.length).toBe(1)
    expect(getText(result!.document)).toEqual([['helloworld']])
  })
})
