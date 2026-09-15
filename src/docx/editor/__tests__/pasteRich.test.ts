import { describe, expect, it } from 'vitest'

import { applyCommand } from '../commands'
import type { Command, Position } from '../commandTypes'
import { htmlToPasteBlocks } from '../pasteBlocks'
import { buildRichPasteCommands, type RichPasteBundleContext } from '../pasteRich'
import type {
  Comment,
  Document,
  Endnote,
  Footer,
  Footnote,
  Header,
  NumberingDef,
  Paragraph,
  Section,
  Style,
} from '../../model'

// DXE-19 — command-level coverage for turning parsed paste blocks into the
// editor's own Command pipeline (`pasteBlocks.ts` covers the HTML -> block
// parsing step separately).

function position(paragraphPath: ReadonlyArray<number>, runIndex: number, charOffset: number): Position {
  return { paragraphPath: Object.freeze([...paragraphPath]), runIndex, charOffset }
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

function createParagraph(text: string): Paragraph {
  return Object.freeze({
    kind: 'paragraph',
    children: Object.freeze([{ kind: 'run' as const, children: Object.freeze([{ kind: 'text' as const, value: text }]) }]),
  }) satisfies Paragraph
}

function emptyBundleContext(document: Document): RichPasteBundleContext {
  return {
    relationships: [],
    rawArchive: new Map(),
    contentTypes: { defaults: [], overrides: [] },
    numberingPart: undefined,
    numbering: document.numbering,
  }
}

function flattenText(doc: Document): string {
  const paragraph = doc.sections[0].blocks[0]
  if (paragraph.kind !== 'paragraph') {
    throw new Error('expected a paragraph')
  }
  return paragraph.children
    .flatMap((child) => {
      if (child.kind === 'run') return child.children
      if (child.kind === 'hyperlink') return child.children.flatMap((run) => (run.kind === 'run' ? run.children : []))
      return []
    })
    .filter((node): node is { kind: 'text'; value: string } => node.kind === 'text')
    .map((node) => node.value)
    .join('')
}

async function pasteAt(document: Document, html: string, at: Position): Promise<Document> {
  const blocks = htmlToPasteBlocks(html)
  const result = await buildRichPasteCommands(document, emptyBundleContext(document), blocks, at, {
    anchor: at,
    focus: at,
  })
  expect(result).not.toBeNull()
  const batch: Command =
    result!.commands.length === 1 ? result!.commands[0] : { kind: 'composite', commands: result!.commands }
  return applyCommand(document, batch).document
}

describe('DXE-19 — buildRichPasteCommands', () => {
  it('returns null for an empty block list', async () => {
    const document = createDocument([createParagraph('Hello')])
    const result = await buildRichPasteCommands(document, emptyBundleContext(document), [], position([0], 0, 0), null)
    expect(result).toBeNull()
  })

  it('continues correctly after a formatted run splits the run it was inserted into', async () => {
    // Regression test: `apply-run-format` re-splits the single run
    // `insert-text` just created (isolating "bold" from its neighbors),
    // which shifts every run index after the split. Failing to follow that
    // command's own returned range before addressing the *next* run used to
    // throw "InsertText position is outside the paragraph" as soon as a
    // paste mixed formatted and plain text in the same paragraph.
    const document = createDocument([createParagraph('Howdy DOCX')])
    const cursor = position([0], 0, 5) // right after "Howdy", before " DOCX"

    const next = await pasteAt(document, '<p>Pasted <b>bold</b> and plain</p>', cursor)

    expect(flattenText(next)).toBe('HowdyPasted bold and plain DOCX')
  })

  it('continues correctly after a hyperlink wraps a formatted run', async () => {
    const document = createDocument([createParagraph('Start End')])
    const cursor = position([0], 0, 5) // right after "Start", before " End"

    const next = await pasteAt(
      document,
      '<p>before <a href="https://example.com"><b>link</b></a> after</p>',
      cursor,
    )

    expect(flattenText(next)).toBe('Startbefore link after End')

    const paragraph = next.sections[0].blocks[0]
    if (paragraph.kind !== 'paragraph') throw new Error('expected a paragraph')
    expect(paragraph.children.some((child) => child.kind === 'hyperlink')).toBe(true)
  })

  it('inserts a pasted table verbatim via insert-table, then resumes the streamed body afterward', async () => {
    const document = createDocument([createParagraph('')])
    const cursor = position([0], 0, 0)

    const html = '<table><tr><td>A1</td><td>B1</td></tr><tr><td>A2</td><td>B2</td></tr></table>'
    const blocks = htmlToPasteBlocks(html)
    const result = await buildRichPasteCommands(document, emptyBundleContext(document), blocks, cursor, {
      anchor: cursor,
      focus: cursor,
    })

    expect(result).not.toBeNull()
    expect(result!.commands.some((command) => command.kind === 'insert-table')).toBe(true)

    const batch: Command =
      result!.commands.length === 1 ? result!.commands[0] : { kind: 'composite', commands: result!.commands }
    const applied = applyCommand(document, batch)
    // `insert-table` splits the paragraph the cursor was in around the
    // table (before/table/after) — pasting at the very start of an empty
    // paragraph puts the (empty) "before" half at index 0 and the table
    // itself at index 1.
    const tableBlock = applied.document.sections[0].blocks[1]
    expect(tableBlock.kind).toBe('table')
    if (tableBlock.kind === 'table') {
      expect(tableBlock.rows).toHaveLength(2)
      const firstRow = tableBlock.rows[0]
      if (firstRow.kind === 'table-row') {
        expect(firstRow.cells).toHaveLength(2)
      } else {
        throw new Error('expected the first table row')
      }
    }
  })

  it('allocates one shared numId for a pasted bullet list and applies it after every item exists', async () => {
    const document = createDocument([createParagraph('')])
    const cursor = position([0], 0, 0)

    const next = await pasteAt(document, '<ul><li>One</li><li>Two</li></ul>', cursor)

    const numIds = next.sections[0].blocks
      .filter((block): block is Paragraph => block.kind === 'paragraph')
      .map((paragraph) => paragraph.props?.numPr?.numId)
      .filter((numId): numId is string => numId !== undefined)

    expect(numIds.length).toBeGreaterThanOrEqual(2)
    expect(new Set(numIds).size).toBe(1)
  })
})
