import { describe, expect, it } from 'vitest'

import { applyCommand } from '../commands'
import type { Command, Position, Range, TrackChangesContext } from '../commandTypes'
import { writeDocumentXml } from '../../serializer/documentWriter'
import type {
  Comment,
  Document,
  Endnote,
  Footer,
  Footnote,
  Header,
  NumberingDef,
  ParagraphChild,
  Paragraph,
  Run,
  Section,
  Style,
} from '../../model'

// DXE-11 — command-level coverage for recording typed insertions/deletions as
// w:ins/w:del when Track Changes is on. See commands.ts's applyInsertText/
// applyDeleteSpan/buildTrackedDeleteParagraph doc comments for the design
// this exercises.

const TRACK: TrackChangesContext = { enabled: true, author: 'Atlas', date: '2024-01-01T00:00:00.000Z' }
const TRACK_BOB: TrackChangesContext = { enabled: true, author: 'Bob', date: '2024-01-02T00:00:00.000Z' }

describe('DXE-11 — insertions recorded as w:ins', () => {
  it('wraps freshly typed text in a new ins-revision, leaving the surrounding plain text untouched', () => {
    const original = createDocument([createParagraph(['Atlas'])])

    const result = applyCommand(
      original,
      { kind: 'insert-text', at: position([0], 0, 2), text: '++' },
      TRACK,
    )

    const children = paragraphChildren(result.document, [0])
    expect(children).toHaveLength(3)
    expect(runText(children[0])).toBe('At')
    expect(children[1].kind).toBe('ins-revision')
    expect(children[1]).toMatchObject({ author: 'Atlas', date: TRACK.date })
    expect(revisionText(children[1])).toBe('++')
    expect(runText(children[2])).toBe('las')

    // The original document is never mutated in place.
    expect(paragraphChildren(original, [0])).toHaveLength(1)
  })

  it('extends a still-pending insertion by the same author instead of nesting a new w:ins', () => {
    const original = createDocument([createParagraph([''])])

    const first = applyCommand(
      original,
      { kind: 'insert-text', at: position([0], 0, 0), text: 'AB' },
      TRACK,
    )
    const firstChildren = paragraphChildren(first.document, [0])
    expect(firstChildren).toHaveLength(1)
    expect(firstChildren[0].kind).toBe('ins-revision')
    const firstId = (firstChildren[0] as { id: string }).id

    const second = applyCommand(first.document, { kind: 'insert-text', at: requireRange(first).focus, text: 'C' }, TRACK)

    const secondChildren = paragraphChildren(second.document, [0])
    expect(secondChildren).toHaveLength(1)
    expect(secondChildren[0].kind).toBe('ins-revision')
    expect((secondChildren[0] as { id: string }).id).toBe(firstId)
    expect(revisionText(secondChildren[0])).toBe('ABC')
  })

  it('starts a separate w:ins when typing next to a tracked insertion by a different author', () => {
    const original = createDocument([
      createParagraphWithChildren([createInsRevision('5', 'Bob', 'Bob added')]),
    ])

    const result = applyCommand(
      original,
      { kind: 'insert-text', at: position([0], 0, 'Bob added'.length), text: 'X' },
      TRACK,
    )

    const children = paragraphChildren(result.document, [0])
    expect(children).toHaveLength(2)
    expect(children[0]).toMatchObject({ kind: 'ins-revision', id: '5', author: 'Bob' })
    expect(children[1]).toMatchObject({ kind: 'ins-revision', author: 'Atlas' })
    // Fresh ids never collide with one already used in the document.
    expect((children[1] as { id: string }).id).not.toBe('5')
    expect(revisionText(children[1])).toBe('X')
  })

  it('does not wrap insertions when Track Changes is off', () => {
    const original = createDocument([createParagraph(['Atlas'])])

    const result = applyCommand(original, { kind: 'insert-text', at: position([0], 0, 2), text: '++' })

    const children = paragraphChildren(result.document, [0])
    expect(children).toHaveLength(1)
    expect(children[0].kind).toBe('run')
    expect(runText(children[0])).toBe('At++las')
  })
})

describe('DXE-11 — deletions recorded as w:del', () => {
  it('keeps deleted text as a del-revision instead of removing it', () => {
    const original = createDocument([createParagraph(['Atlas'])])

    const cmd: Command = {
      kind: 'delete-range',
      range: { anchor: position([0], 0, 1), focus: position([0], 0, 4) },
    }
    const result = applyCommand(original, cmd, TRACK)

    const children = paragraphChildren(result.document, [0])
    expect(children).toHaveLength(3)
    expect(runText(children[0])).toBe('A')
    expect(children[1]).toMatchObject({ kind: 'del-revision', author: 'Atlas' })
    expect(revisionText(children[1])).toBe('tla')
    expect(runText(children[2])).toBe('s')

    // Undo restores the exact original document.
    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('hard-deletes a still-pending insertion by the same author instead of nesting a w:del', () => {
    const original = createDocument([
      createParagraphWithChildren([createInsRevision('9', 'Atlas', 'Hello')]),
    ])

    const cmd: Command = {
      kind: 'delete-range',
      range: { anchor: position([0], 0, 0), focus: position([0], 0, 'Hello'.length) },
    }
    const result = applyCommand(original, cmd, TRACK)

    const children = paragraphChildren(result.document, [0])
    expect(children).toHaveLength(1)
    expect(children[0].kind).toBe('run')
    expect(runText(children[0])).toBe('')

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('splits a mixed selection: your own pending insertion is dropped, plain text becomes a w:del', () => {
    const original = createDocument([
      createParagraphWithChildren([
        createRun('before '),
        createInsRevision('9', 'Atlas', 'pending'),
        createRun(' after'),
      ]),
    ])

    // Deletes "re " (tail of the plain run) + "pend" (prefix of the pending
    // insertion): flattened offsets 4..11 across "before "(7) + "pending"(7).
    const cmd: Command = {
      kind: 'delete-range',
      range: { anchor: position([0], 0, 4), focus: position([0], 1, 4) },
    }
    const result = applyCommand(original, cmd, TRACK)

    const children = paragraphChildren(result.document, [0])
    expect(children).toHaveLength(4)
    expect(runText(children[0])).toBe('befo')
    expect(children[1]).toMatchObject({ kind: 'del-revision' })
    expect(revisionText(children[1])).toBe('re ')
    // The remainder of the original pending insertion survives, unmodified.
    expect(children[2]).toMatchObject({ kind: 'ins-revision', id: '9', author: 'Atlas' })
    expect(revisionText(children[2])).toBe('ing')
    expect(runText(children[3])).toBe(' after')
  })

  it('falls back to an untracked delete across paragraph boundaries even when Track Changes is on', () => {
    const original = createDocument([createParagraph(['Hello']), createParagraph(['World'])])

    const cmd: Command = {
      kind: 'delete-range',
      range: { anchor: position([0], 0, 3), focus: position([1], 0, 2) },
    }
    const result = applyCommand(original, cmd, TRACK)

    expect(result.document.sections[0].blocks).toHaveLength(1)
    const children = paragraphChildren(result.document, [0])
    expect(children).toHaveLength(1)
    expect(children[0].kind).toBe('run')
    expect(runText(children[0])).toBe('Helrld')
  })

  it('does not wrap deletions when Track Changes is off', () => {
    const original = createDocument([createParagraph(['Atlas'])])
    const cmd: Command = {
      kind: 'delete-range',
      range: { anchor: position([0], 0, 1), focus: position([0], 0, 4) },
    }

    const result = applyCommand(original, cmd)

    const children = paragraphChildren(result.document, [0])
    expect(children).toHaveLength(1)
    expect(runText(children[0])).toBe('As')
  })
})

describe('DXE-11 — accept/reject still resolve tracked edits recorded by typing', () => {
  it('accept keeps a typed insertion as plain text; reject removes it', () => {
    const original = createDocument([createParagraph(['Atlas'])])
    const inserted = applyCommand(
      original,
      { kind: 'insert-text', at: position([0], 0, 2), text: '++' },
      TRACK,
    )

    const accepted = applyCommand(inserted.document, { kind: 'accept-revision', paragraphPath: [0, 0], childIndex: 1 })
    const acceptedChildren = paragraphChildren(accepted.document, [0])
    expect(acceptedChildren.every((child) => child.kind === 'run')).toBe(true)
    expect(acceptedChildren.map(runText).join('')).toBe('At++las')

    const rejected = applyCommand(inserted.document, { kind: 'reject-revision', paragraphPath: [0, 0], childIndex: 1 })
    const rejectedChildren = paragraphChildren(rejected.document, [0])
    expect(rejectedChildren.map(runText).join('')).toBe('Atlas')
  })

  it('reject restores text recorded by a tracked deletion', () => {
    const original = createDocument([createParagraph(['Atlas'])])
    const deleted = applyCommand(
      original,
      { kind: 'delete-range', range: { anchor: position([0], 0, 1), focus: position([0], 0, 4) } },
      TRACK,
    )

    const rejected = applyCommand(deleted.document, { kind: 'reject-revision', paragraphPath: [0, 0], childIndex: 1 })
    const rejectedChildren = paragraphChildren(rejected.document, [0])
    expect(rejectedChildren.map(runText).join('')).toBe('Atlas')

    const accepted = applyCommand(deleted.document, { kind: 'accept-revision', paragraphPath: [0, 0], childIndex: 1 })
    const acceptedChildren = paragraphChildren(accepted.document, [0])
    expect(acceptedChildren.map(runText).join('')).toBe('As')
  })
})

describe('DXE-11 — getEditableRuns flattens through existing revisions', () => {
  it('lets a paragraph containing an existing tracked insertion stay editable elsewhere', () => {
    const original = createDocument([
      createParagraphWithChildren([
        createRun('Before '),
        createInsRevision('2', 'Bob', 'Bob added'),
        createRun(' after'),
      ]),
    ])

    // Untracked insert appended at the very end — previously threw because
    // getEditableRuns rejected the whole paragraph on the first ins-revision
    // child it saw.
    const result = applyCommand(original, {
      kind: 'insert-text',
      at: position([0], 2, ' after'.length),
      text: '!',
    })

    const children = paragraphChildren(result.document, [0])
    expect(children).toHaveLength(3)
    expect(children[1]).toMatchObject({ kind: 'ins-revision', id: '2', author: 'Bob' })
    expect(runText(children[2])).toBe(' after!')
  })
})

describe('DXE-11 — serializer round-trip', () => {
  it('serializes a typed tracked insertion as a valid w:ins element', () => {
    const original = createDocument([createParagraph(['Atlas'])])
    const result = applyCommand(
      original,
      { kind: 'insert-text', at: position([0], 0, 2), text: '++' },
      TRACK,
    )

    const xml = writeDocumentXml(result.document)
    expect(xml).toContain('<w:ins')
    expect(xml).toContain('w:author="Atlas"')
    expect(xml).toMatch(/<w:t[^>]*>\+\+<\/w:t>/)
  })

  it('serializes a tracked deletion as a valid w:del/w:delText element', () => {
    const original = createDocument([createParagraph(['Atlas'])])
    const result = applyCommand(
      original,
      { kind: 'delete-range', range: { anchor: position([0], 0, 1), focus: position([0], 0, 4) } },
      TRACK,
    )

    const xml = writeDocumentXml(result.document)
    expect(xml).toContain('<w:del')
    expect(xml).toMatch(/<w:delText[^>]*>tla<\/w:delText>/)
  })
})

describe('DXE-11 — same-author tracking is scoped to a single trackChanges call', () => {
  it('starts a new w:ins when the caret returns after an intervening plain (untracked) edit', () => {
    const original = createDocument([createParagraph([''])])

    const tracked = applyCommand(original, { kind: 'insert-text', at: position([0], 0, 0), text: 'A' }, TRACK)
    // A different author's session typing next to it should not extend "A"'s revision.
    const other = applyCommand(tracked.document, { kind: 'insert-text', at: requireRange(tracked).focus, text: 'B' }, TRACK_BOB)

    const children = paragraphChildren(other.document, [0])
    expect(children).toHaveLength(2)
    expect(children[0]).toMatchObject({ kind: 'ins-revision', author: 'Atlas' })
    expect(children[1]).toMatchObject({ kind: 'ins-revision', author: 'Bob' })
  })
})

describe('DXE-11 — a composite command forwards trackChanges to every sub-command', () => {
  // Regression: `applyEditorCommands`/`handleRichPaste` in DocxViewer.tsx
  // wrap paste (and Replace All) in a single `composite` command so the
  // whole batch undoes in one step. `applyComposite` used to call
  // `applyCommand` on each sub-command with no `trackChanges` argument at
  // all, so pasting while Track Changes was on silently inserted plain,
  // untracked text — defeating the point of turning it on for exactly the
  // edit a reviewer most needs visibility into.
  it('records every insert-text in a paste-shaped composite as its own w:ins', () => {
    const original = createDocument([createParagraph(['Hello'])])

    const result = applyCommand(
      original,
      {
        kind: 'composite',
        commands: [
          { kind: 'insert-paragraph-break', at: position([0], 0, 5) },
          { kind: 'insert-text', at: position([1], 0, 0), text: 'World' },
        ],
      },
      TRACK,
    )

    const secondParagraphChildren = paragraphChildren(result.document, [1])
    expect(secondParagraphChildren).toHaveLength(1)
    expect(secondParagraphChildren[0]).toMatchObject({ kind: 'ins-revision', author: 'Atlas' })
    expect(revisionText(secondParagraphChildren[0])).toBe('World')
  })

  it('leaves a composite untracked when no trackChanges context is given, matching a single command', () => {
    const original = createDocument([createParagraph([''])])

    const result = applyCommand(original, {
      kind: 'composite',
      commands: [{ kind: 'insert-text', at: position([0], 0, 0), text: 'World' }],
    })

    const children = paragraphChildren(result.document, [0])
    expect(children).toHaveLength(1)
    expect(children[0].kind).toBe('run')
  })
})

// ─── helpers ──────────────────────────────────────────────────────────────

/** `applyCommand`'s general return type can't express that `insert-text`
 * always yields a `range` — narrow it once here rather than at every call
 * site. */
function requireRange(result: { readonly range?: Range }): Range {
  if (result.range === undefined) {
    throw new Error('Expected applyCommand to return a range')
  }
  return result.range
}

function position(paragraphPath: ReadonlyArray<number>, runIndex: number, charOffset: number): Position {
  return {
    paragraphPath: Object.freeze([...paragraphPath]),
    runIndex,
    charOffset,
  }
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

function createRun(text: string): Run {
  return Object.freeze({
    kind: 'run',
    children: Object.freeze([Object.freeze({ kind: 'text', value: text })]),
  }) satisfies Run
}

function createParagraph(texts: ReadonlyArray<string>): Paragraph {
  return createParagraphWithChildren(texts.map((text) => createRun(text)))
}

function createParagraphWithChildren(children: ReadonlyArray<ParagraphChild>): Paragraph {
  return Object.freeze({
    kind: 'paragraph',
    children: Object.freeze([...children]),
  }) satisfies Paragraph
}

function createInsRevision(id: string, author: string, text: string): ParagraphChild {
  return Object.freeze({
    kind: 'ins-revision',
    id,
    author,
    children: Object.freeze([createRun(text)]) as ReadonlyArray<ParagraphChild> & ReadonlyArray<Run>,
  })
}

/** Every test in this file uses a single section with top-level paragraphs
 * only, so `[blockIndex]` is enough addressing. */
function paragraphChildren(document: Document, paragraphPath: readonly [number]): ReadonlyArray<ParagraphChild> {
  const [blockIndex] = paragraphPath
  const paragraph = document.sections[0].blocks[blockIndex] as Paragraph
  return paragraph.children
}

function runText(child: ParagraphChild): string {
  if (child.kind !== 'run') {
    return ''
  }
  return child.children.map((grandchild) => (grandchild.kind === 'text' ? grandchild.value : '')).join('')
}

function revisionText(child: ParagraphChild): string {
  if (child.kind !== 'ins-revision' && child.kind !== 'del-revision') {
    return ''
  }
  return (child.children as ReadonlyArray<ParagraphChild>).map(runText).join('')
}
