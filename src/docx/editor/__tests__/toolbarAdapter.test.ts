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
  ParagraphChild,
  Run,
  Section,
  Style,
} from '../../model'
import type { Position, Range } from '../commandTypes'
import { toolbarToCommand } from '../toolbarAdapter'

function pos(paragraphPath: ReadonlyArray<number>, runIndex: number, charOffset: number): Position {
  return { paragraphPath: Object.freeze([...paragraphPath]), runIndex, charOffset }
}

function collapsed(p: Position): Range {
  return { anchor: p, focus: p }
}

function createRun(text: string): Run {
  return Object.freeze({
    kind: 'run',
    children: Object.freeze([Object.freeze({ kind: 'text', value: text })]),
  }) as Run
}

function createRevisionParagraph(kind: 'ins-revision' | 'del-revision'): Paragraph {
  const revision: ParagraphChild = Object.freeze({
    kind,
    id: 'r1',
    children: Object.freeze([createRun('tracked')]),
  }) as ParagraphChild

  return Object.freeze({ kind: 'paragraph', children: Object.freeze([revision]) }) as Paragraph
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

describe('toolbarToCommand', () => {
  // ---------------------------------------------------------------------------
  // D17 — accept/reject wiring
  // ---------------------------------------------------------------------------

  it('accept-change resolves the ins-revision at the selection paragraph', () => {
    const document = createDocument([createRevisionParagraph('ins-revision')])
    const command = toolbarToCommand({ kind: 'accept-change' }, collapsed(pos([0], 0, 0)), document)

    expect(command).toEqual({ kind: 'accept-revision', paragraphPath: [0], childIndex: 0 })
  })

  it('reject-change resolves the del-revision at the selection paragraph', () => {
    const document = createDocument([createRevisionParagraph('del-revision')])
    const command = toolbarToCommand({ kind: 'reject-change' }, collapsed(pos([0], 0, 0)), document)

    expect(command).toEqual({ kind: 'reject-revision', paragraphPath: [0], childIndex: 0 })
  })

  it('accept-change returns null when there is no selection', () => {
    const document = createDocument([createRevisionParagraph('ins-revision')])
    expect(toolbarToCommand({ kind: 'accept-change' }, null, document)).toBeNull()
  })

  it('accept-change returns null when the selected paragraph has no revision', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('plain')]) }) as Paragraph,
    ])
    expect(toolbarToCommand({ kind: 'accept-change' }, collapsed(pos([0], 0, 0)), document)).toBeNull()
  })

  it('reject-change returns null when the paragraph path does not resolve', () => {
    const document = createDocument([createRevisionParagraph('ins-revision')])
    expect(toolbarToCommand({ kind: 'reject-change' }, collapsed(pos([9], 0, 0)), document)).toBeNull()
  })

  it('accept-all-changes builds an accept-all-revisions command regardless of selection', () => {
    const document = createDocument([createRevisionParagraph('ins-revision')])
    expect(toolbarToCommand({ kind: 'accept-all-changes' }, null, document)).toEqual({
      kind: 'accept-all-revisions',
    })
  })

  it('reject-all-changes builds a reject-all-revisions command regardless of selection', () => {
    const document = createDocument([createRevisionParagraph('del-revision')])
    expect(toolbarToCommand({ kind: 'reject-all-changes' }, collapsed(pos([0], 0, 0)), document)).toEqual({
      kind: 'reject-all-revisions',
    })
  })

  // ---------------------------------------------------------------------------
  // Existing wiring stays intact after the signature change
  // ---------------------------------------------------------------------------

  it('toggle-bold still builds an apply-run-format command', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('hello')]) }) as Paragraph,
    ])
    const range: Range = { anchor: pos([0], 0, 0), focus: pos([0], 0, 5) }

    const command = toolbarToCommand({ kind: 'toggle-bold' }, range, document)

    expect(command).toEqual({ kind: 'apply-run-format', range, format: { bold: true } })
  })

  it('insert-table builds a command from the selection focus', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('hello')]) }) as Paragraph,
    ])
    const command = toolbarToCommand(
      { kind: 'insert-table', rows: 2, cols: 3 },
      collapsed(pos([0], 0, 2)),
      document,
    )

    expect(command).toEqual({ kind: 'insert-table', at: pos([0], 0, 2), rows: 2, cols: 3 })
  })

  it('insert-page-break builds an insert-inline command with a page BreakNode at the caret', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('hello')]) }) as Paragraph,
    ])
    const command = toolbarToCommand(
      { kind: 'insert-page-break' },
      collapsed(pos([0], 0, 2)),
      document,
    )

    expect(command).toEqual({
      kind: 'insert-inline',
      at: pos([0], 0, 2),
      child: { kind: 'break', breakType: 'page' },
    })
  })

  it('insert-page-break returns null when there is no selection', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('hello')]) }) as Paragraph,
    ])
    expect(toolbarToCommand({ kind: 'insert-page-break' }, null, document)).toBeNull()
  })

  it('insert-hyperlink is still handled by the caller (bundle-aware), not this pure mapper', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('hello')]) }) as Paragraph,
    ])
    expect(toolbarToCommand({ kind: 'insert-hyperlink' }, collapsed(pos([0], 0, 0)), document)).toBeNull()
  })

  // ---------------------------------------------------------------------------
  // D18 — a multi-paragraph selection includes every paragraph in between,
  // not just its two endpoints
  // ---------------------------------------------------------------------------

  it('toggle-bullet-list targets every paragraph a multi-paragraph selection spans', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('One')]) }) as Paragraph,
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('Two')]) }) as Paragraph,
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('Three')]) }) as Paragraph,
    ])
    const range: Range = { anchor: pos([0, 0], 0, 0), focus: pos([0, 2], 0, 2) }

    const command = toolbarToCommand({ kind: 'toggle-bullet-list' }, range, document)

    expect(command).toEqual({
      kind: 'insert-list',
      paragraphPaths: [[0, 0], [0, 1], [0, 2]],
      numId: 1,
      level: 0,
    })
  })

  it('set-alignment targets every paragraph regardless of anchor/focus order', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('One')]) }) as Paragraph,
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('Two')]) }) as Paragraph,
    ])
    // Focus before anchor (user selected upward).
    const range: Range = { anchor: pos([0, 1], 0, 3), focus: pos([0, 0], 0, 0) }

    const command = toolbarToCommand({ kind: 'set-alignment', align: 'center' }, range, document)

    expect(command).toMatchObject({ kind: 'apply-para-format', paragraphPaths: [[0, 0], [0, 1]] })
  })
})
