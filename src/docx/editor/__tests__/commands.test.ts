import { describe, expect, it } from 'vitest'

import type { Position } from '../index'
import { applyCommand } from '../index'
import { findParagraph, replaceParagraph } from '../commands'
import type {
  Comment,
  Document,
  Endnote,
  Footer,
  Footnote,
  Header,
  NumberingDef,
  ParaProps,
  Paragraph,
  Run,
  RunProps,
  Section,
  Style,
  Table,
  TableRow,
} from '../../model'
import { twip } from '../../model'

describe('docx editor commands', () => {
  it('findParagraph resolves a section-explicit paragraph path', () => {
    const document = createDocument([createParagraph(['Atlas'])])

    expect(findParagraph(document, [0, 0])).toEqual(document.sections[0].blocks[0])
  })

  it('replaceParagraph swaps only the targeted paragraph', () => {
    const original = createDocument([createParagraph(['One']), createParagraph(['Two'])])
    const replacement = createParagraph(['Updated'])

    const next = replaceParagraph(original, [1], replacement)

    expect(next.sections[0].blocks[0]).toBe(original.sections[0].blocks[0])
    expect(next.sections[0].blocks[1]).toEqual(replacement)
  })

  it('InsertText updates the target run text', () => {
    const original = createDocument([createParagraph(['Atlas'])])

    const result = applyCommand(original, {
      kind: 'insert-text',
      at: position([0], 0, 2),
      text: '++',
    })

    expect(paragraphTexts(result.document)).toEqual([['At++las']])
    expect(result.inverse).toEqual({
      kind: 'delete-range',
      range: {
        anchor: position([0], 0, 2),
        focus: position([0], 0, 4),
      },
    })
  })

  it('InsertText round-trips through its inverse on an empty paragraph', () => {
    const original = createDocument([createParagraph([''])])

    const applied = applyCommand(original, {
      kind: 'insert-text',
      at: position([0], 0, 0),
      text: 'Atlas',
    })
    const reverted = applyCommand(applied.document, applied.inverse)

    expect(reverted.document).toEqual(original)
  })

  it('DeleteRange removes text inside a single run', () => {
    const original = createDocument([createParagraph(['Atlas'])])

    const result = applyCommand(original, {
      kind: 'delete-range',
      range: {
        anchor: position([0], 0, 1),
        focus: position([0], 0, 4),
      },
    })

    expect(paragraphTexts(result.document)).toEqual([['As']])
    // DXE-04/DXE-05 — delete now uses the general replace-blocks primitive
    // (the same one cross-paragraph delete needs) so its inverse is an exact
    // paragraph snapshot rather than a single insert-text; what matters is
    // that applying it round-trips back to the original document.
    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('DeleteRange spanning two paragraphs merges them and round-trips through its inverse', () => {
    const original = createDocument([
      createParagraph(['Alpha']),
      createParagraph(['Beta']),
    ])

    const result = applyCommand(original, {
      kind: 'delete-range',
      range: {
        anchor: position([0], 0, 3),
        focus: position([1], 0, 2),
      },
    })

    expect(result.document.sections[0].blocks.length).toBe(1)
    expect(paragraphTexts(result.document)).toEqual([['Alpta']])

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('DeleteRange spanning three paragraphs drops the fully-enclosed middle one', () => {
    const original = createDocument([
      createParagraph(['One']),
      createParagraph(['Two']),
      createParagraph(['Three']),
    ])

    const result = applyCommand(original, {
      kind: 'delete-range',
      range: {
        anchor: position([0], 0, 1),
        focus: position([2], 0, 2),
      },
    })

    expect(paragraphTexts(result.document)).toEqual([['Oree']])

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('DeleteRange across a multi-run selection within one paragraph slices every intersected run', () => {
    const original = createDocument([
      createParagraph(['Al', 'pha', 'Bet'], undefined, [undefined, { bold: true }, undefined]),
    ])

    const result = applyCommand(original, {
      kind: 'delete-range',
      range: {
        anchor: position([0], 0, 1),
        focus: position([0], 2, 2),
      },
    })

    expect(paragraphTexts(result.document)).toEqual([['At']])

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('DeleteRange at an out-of-range paragraph path throws rather than silently no-opping', () => {
    // Callers (Input.ts/DocxViewer) rely on a thrown error — not a silent
    // no-op — to decide whether to preventDefault() so the DOM never diverges
    // from the model (DXE-03).
    const document = createDocument([createParagraph(['Alpha'])])
    expect(() =>
      applyCommand(document, {
        kind: 'delete-range',
        range: { anchor: position([5], 0, 0), focus: position([5], 0, 1) },
      }),
    ).toThrow()
  })

  it('typing/deleting inside a hyperlink works via the flattened run list (DXE-03)', () => {
    const hyperlinkRun: Run = Object.freeze({
      kind: 'run',
      children: Object.freeze([Object.freeze({ kind: 'text', value: 'link text' })]),
    }) as Run
    const hyperlink = Object.freeze({
      kind: 'hyperlink' as const,
      relationshipId: 'rId5',
      children: Object.freeze([hyperlinkRun]),
    })
    const paragraph: Paragraph = Object.freeze({
      kind: 'paragraph',
      children: Object.freeze([createRun('Before '), hyperlink, createRun(' after')]),
    }) as Paragraph
    const original = createDocument([paragraph])

    const result = applyCommand(original, {
      kind: 'insert-text',
      at: position([0], 1, 4),
      text: 'XX',
    })

    const updatedParagraph = result.document.sections[0].blocks[0] as Paragraph
    expect(updatedParagraph.children[1]).toMatchObject({ kind: 'hyperlink', relationshipId: 'rId5' })
    expect(flattenedRunTexts(updatedParagraph)).toEqual(['Before ', 'linkXX text', ' after'])

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('DeleteRange round-trips a paragraph-break delete returned by InsertParagraphBreak', () => {
    const original = createDocument([createParagraph(['Atlas'])])

    const split = applyCommand(original, {
      kind: 'insert-paragraph-break',
      at: position([0], 0, 2),
    })
    const merged = applyCommand(split.document, split.inverse)

    expect(merged.document).toEqual(original)
  })

  it('InsertParagraphBreak splits a run into two paragraphs', () => {
    const original = createDocument([createParagraph(['Atlas'])])

    const result = applyCommand(original, {
      kind: 'insert-paragraph-break',
      at: position([0], 0, 2),
    })

    expect(paragraphTexts(result.document)).toEqual([['At'], ['las']])
    expect(result.inverse).toEqual({
      kind: 'delete-range',
      range: {
        anchor: position([0], 0, 2),
        focus: position([1], 0, 0),
      },
    })
  })

  it('InsertParagraphBreak round-trips through DeleteRange at the boundary', () => {
    const original = createDocument([
      createParagraph(['Alpha', 'Beta'], undefined, [undefined, { bold: true }]),
    ])

    const applied = applyCommand(original, {
      kind: 'insert-paragraph-break',
      at: position([0], 0, 5),
    })
    const reverted = applyCommand(applied.document, applied.inverse)

    expect(reverted.document).toEqual(original)
  })

  it('ApplyRunFormat formats the selected text slice', () => {
    const original = createDocument([createParagraph(['Atlas'])])

    const result = applyCommand(original, {
      kind: 'apply-run-format',
      range: {
        anchor: position([0], 0, 1),
        focus: position([0], 0, 4),
      },
      format: { bold: true },
    })

    const paragraph = result.document.sections[0].blocks[0] as Paragraph
    expect(runTexts(paragraph)).toEqual(['A', 'tla', 's'])
    expect((paragraph.children[1] as Run).props).toEqual({ bold: true })
  })

  it('ApplyRunFormat round-trips through its inverse', () => {
    const original = createDocument([createParagraph(['Atlas'])])

    const applied = applyCommand(original, {
      kind: 'apply-run-format',
      range: {
        anchor: position([0], 0, 1),
        focus: position([0], 0, 4),
      },
      format: { italic: true },
    })
    const reverted = applyCommand(applied.document, applied.inverse)

    expect(reverted.document).toEqual(original)
  })

  it('ApplyParaFormat updates each targeted paragraph', () => {
    const original = createDocument([
      createParagraph(['One']),
      createParagraph(['Two']),
    ])

    const result = applyCommand(original, {
      kind: 'apply-para-format',
      paragraphPaths: [[0], [1]],
      format: { jc: 'center' },
    })

    expect((result.document.sections[0].blocks[0] as Paragraph).props).toEqual({ jc: 'center' })
    expect((result.document.sections[0].blocks[1] as Paragraph).props).toEqual({ jc: 'center' })
  })

  it('ApplyParaFormat round-trips through its inverse', () => {
    const original = createDocument([
      createParagraph(['One'], { spacing: { before: twip(120) } }),
      createParagraph(['Two'], { spacing: { before: twip(120) } }),
    ])

    const applied = applyCommand(original, {
      kind: 'apply-para-format',
      paragraphPaths: [[0], [1]],
      format: { spacing: { after: twip(240) } },
    })
    const reverted = applyCommand(applied.document, applied.inverse)

    expect(reverted.document).toEqual(original)
  })

  it('ApplyStyle writes the paragraph style id', () => {
    const original = createDocument([createParagraph(['Heading'])])

    const result = applyCommand(original, {
      kind: 'apply-style',
      paragraphPath: [0],
      styleId: 'Heading1',
    })

    expect((result.document.sections[0].blocks[0] as Paragraph).props).toEqual({ pStyle: 'Heading1' })
  })

  it('ApplyStyle round-trips through its inverse', () => {
    const original = createDocument([createParagraph(['Heading'])])

    const applied = applyCommand(original, {
      kind: 'apply-style',
      paragraphPath: [0],
      styleId: 'Heading1',
    })
    const reverted = applyCommand(applied.document, applied.inverse)

    expect(reverted.document).toEqual(original)
  })

  it('AcceptRevision on ins-revision unwraps children into the paragraph', () => {
    const original = createParagraphWithRevision('ins-revision', 'kept')
    const result = applyCommand(original, {
      kind: 'accept-revision',
      paragraphPath: [0],
      childIndex: 0,
    })
    expect(paragraphTexts(result.document)).toEqual([['kept']])
  })

  it('RejectRevision on ins-revision drops children entirely', () => {
    const original = createParagraphWithRevision('ins-revision', 'gone')
    const result = applyCommand(original, {
      kind: 'reject-revision',
      paragraphPath: [0],
      childIndex: 0,
    })
    expect((result.document.sections[0].blocks[0] as Paragraph).children).toHaveLength(0)
  })

  it('AcceptRevision on del-revision drops children', () => {
    const original = createParagraphWithRevision('del-revision', 'gone')
    const result = applyCommand(original, {
      kind: 'accept-revision',
      paragraphPath: [0],
      childIndex: 0,
    })
    expect((result.document.sections[0].blocks[0] as Paragraph).children).toHaveLength(0)
  })

  it('AcceptAllRevisions resolves every revision across the document', () => {
    const original = createDocument([
      createParagraphWithRevisionRaw('ins-revision', 'A'),
      createParagraphWithRevisionRaw('del-revision', 'B'),
    ])
    const result = applyCommand(original, { kind: 'accept-all-revisions' })
    expect(paragraphTexts(result.document)).toEqual([['A'], []])
  })

  // ---------------------------------------------------------------------------
  // D13 — composite commands + replace-blocks (atomic batches, exact undo)
  // ---------------------------------------------------------------------------

  it('Composite applies every sub-command atomically and undoes as one step', () => {
    const original = createDocument([createParagraph(['Atlas'])])

    const result = applyCommand(original, {
      kind: 'composite',
      commands: [
        { kind: 'insert-text', at: position([0], 0, 5), text: '!' },
        { kind: 'insert-text', at: position([0], 0, 0), text: '>' },
      ],
    })

    expect(paragraphTexts(result.document)).toEqual([['>Atlas!']])

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('Composite rolls back cleanly (throws) when a later sub-command fails', () => {
    const original = createDocument([createParagraph(['Atlas'])])

    expect(() =>
      applyCommand(original, {
        kind: 'composite',
        commands: [
          { kind: 'insert-text', at: position([0], 0, 0), text: 'X' },
          { kind: 'insert-text', at: position([9], 0, 0), text: 'Y' },
        ],
      }),
    ).toThrow()

    // Nothing was mutated in place — `original` is a fresh document each
    // call, but this asserts the thrown call never returned a partially
    // applied document for the caller to accidentally commit.
    expect(paragraphTexts(original)).toEqual([['Atlas']])
  })

  it('ReplaceBlocks swaps N sibling blocks for a new set and inverts exactly', () => {
    const original = createDocument([createParagraph(['One']), createParagraph(['Two'])])
    const replacement = createParagraph(['Merged'])

    const result = applyCommand(original, {
      kind: 'replace-blocks',
      at: [0, 0],
      count: 2,
      blocks: [replacement],
    })

    expect(result.document.sections[0].blocks.length).toBe(1)
    expect(paragraphTexts(result.document)).toEqual([['Merged']])

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  // ---------------------------------------------------------------------------
  // D18 — lists, indent/outdent, hyperlink, page break, table
  // ---------------------------------------------------------------------------

  it('InsertList sets numPr on every targeted paragraph and undoes per-paragraph', () => {
    const original = createDocument([createParagraph(['One']), createParagraph(['Two'])])

    const result = applyCommand(original, {
      kind: 'insert-list',
      paragraphPaths: [[0], [1]],
      numId: 1,
      level: 0,
    })

    const first = result.document.sections[0].blocks[0] as Paragraph
    const second = result.document.sections[0].blocks[1] as Paragraph
    expect(first.props?.numPr).toEqual({ numId: '1', ilvl: 0 })
    expect(second.props?.numPr).toEqual({ numId: '1', ilvl: 0 })

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('InsertList toggles the list back off when every target already has it', () => {
    const original = createDocument([
      createParagraph(['One'], { numPr: { numId: '1', ilvl: 0 } }),
    ])

    const result = applyCommand(original, {
      kind: 'insert-list',
      paragraphPaths: [[0]],
      numId: 1,
      level: 0,
    })

    expect((result.document.sections[0].blocks[0] as Paragraph).props?.numPr).toBeUndefined()
  })

  it('ChangeListLevel increases ilvl on a list paragraph and inverts', () => {
    const original = createDocument([
      createParagraph(['Item'], { numPr: { numId: '1', ilvl: 0 } }),
    ])

    const result = applyCommand(original, {
      kind: 'change-list-level',
      paragraphPath: [0],
      delta: 1,
    })

    expect((result.document.sections[0].blocks[0] as Paragraph).props?.numPr).toEqual({ numId: '1', ilvl: 1 })

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('ChangeListLevel adjusts left indent for a non-list paragraph and inverts', () => {
    const original = createDocument([createParagraph(['Body'])])

    const result = applyCommand(original, {
      kind: 'change-list-level',
      paragraphPath: [0],
      delta: 1,
    })

    expect((result.document.sections[0].blocks[0] as Paragraph).props?.ind?.left).toBe(720)

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('ChangeListLevel is a true no-op at the outdent boundary (does not go negative)', () => {
    const original = createDocument([createParagraph(['Body'])])

    const result = applyCommand(original, {
      kind: 'change-list-level',
      paragraphPath: [0],
      delta: -1,
    })

    expect(result.document).toBe(original)
  })

  it('InsertInline (page break) splits the run and is undoable', () => {
    const original = createDocument([createParagraph(['Atlas'])])

    const result = applyCommand(original, {
      kind: 'insert-inline',
      at: position([0], 0, 2),
      child: { kind: 'break', breakType: 'page' },
    })

    const paragraph = result.document.sections[0].blocks[0] as Paragraph
    expect(paragraph.children.length).toBe(3)
    expect((paragraph.children[1] as Run).children[0]).toEqual({ kind: 'break', breakType: 'page' })
    expect(result.range?.anchor).toEqual(position([0], 2, 0))

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('InsertTable splits the paragraph around a new N x M table with a tblGrid', () => {
    const original = createDocument([createParagraph(['AtlasDoc'])])

    const result = applyCommand(original, {
      kind: 'insert-table',
      at: position([0], 0, 5),
      rows: 2,
      cols: 3,
    })

    const blocks = result.document.sections[0].blocks
    expect(blocks.length).toBe(3)
    expect(blocks[0]).toMatchObject({ kind: 'paragraph' })
    expect(runTexts(blocks[0] as Paragraph)).toEqual(['Atlas'])
    expect(runTexts(blocks[2] as Paragraph)).toEqual(['Doc'])

    const table = blocks[1] as Table
    expect(table.kind).toBe('table')
    expect(table.rows.length).toBe(2)
    expect(table.tblGrid?.length).toBe(3)
    expect((table.rows[0] as TableRow).cells.length).toBe(3)

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('InsertHyperlink wraps a collapsed cursor with the link text and undoes exactly', () => {
    const original = createDocument([createParagraph(['Visit  today'])])

    const result = applyCommand(original, {
      kind: 'insert-hyperlink',
      range: { anchor: position([0], 0, 6), focus: position([0], 0, 6) },
      url: 'https://example.com',
      relationshipId: 'rId9',
    })

    const paragraph = result.document.sections[0].blocks[0] as Paragraph
    expect(flattenedRunTexts(paragraph)).toEqual(['Visit ', 'https://example.com', ' today'])
    expect(paragraph.children[1]).toMatchObject({ kind: 'hyperlink', relationshipId: 'rId9' })

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })

  it('InsertHyperlink wraps a non-collapsed selection in place', () => {
    const original = createDocument([createParagraph(['Read the docs today'])])

    // "Read the docs today" — offset 9 is the 'd' starting "docs", offset 13
    // is right after its trailing 's'.
    const result = applyCommand(original, {
      kind: 'insert-hyperlink',
      range: { anchor: position([0], 0, 9), focus: position([0], 0, 13) },
      url: 'https://docs.example.com',
      relationshipId: 'rId3',
    })

    const paragraph = result.document.sections[0].blocks[0] as Paragraph
    expect(flattenedRunTexts(paragraph)).toEqual(['Read the ', 'docs', ' today'])
    expect(paragraph.children[1]).toMatchObject({ kind: 'hyperlink', relationshipId: 'rId3' })

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(original)
  })
})

function createParagraphWithRevisionRaw(
  kind: 'ins-revision' | 'del-revision',
  text: string,
): Paragraph {
  const run: Run = Object.freeze({
    kind: 'run',
    children: Object.freeze([Object.freeze({ kind: 'text', value: text })]),
  }) as Run
  const revision = Object.freeze({
    kind,
    id: '1',
    children: Object.freeze([run]),
  })
  return Object.freeze({
    kind: 'paragraph',
    children: Object.freeze([revision]),
  }) as Paragraph
}

function createParagraphWithRevision(
  kind: 'ins-revision' | 'del-revision',
  text: string,
): Document {
  return createDocument([createParagraphWithRevisionRaw(kind, text)])
}

function paragraphTexts(document: Document): string[][] {
  return document.sections[0].blocks.map((block) => runTexts(block as Paragraph))
}

function runTexts(paragraph: Paragraph): string[] {
  return paragraph.children.map((child) => getRunText(child as Run))
}

function getRunText(run: Run): string {
  return run.children.map((child) => child.kind === 'text' ? child.value : '').join('')
}

/** Flattens a paragraph's plain runs AND any runs nested inside hyperlinks
 * into one ordered list of visible text, for asserting on hyperlink-aware
 * edits (DXE-03) without depending on the exact child grouping. */
function flattenedRunTexts(paragraph: Paragraph): string[] {
  const texts: string[] = []
  for (const child of paragraph.children) {
    if (child.kind === 'run') {
      texts.push(getRunText(child))
    } else if (child.kind === 'hyperlink') {
      for (const grandchild of child.children) {
        if (grandchild.kind === 'run') {
          texts.push(getRunText(grandchild))
        }
      }
    }
  }
  return texts
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

function createParagraph(
  texts: ReadonlyArray<string>,
  props?: ParaProps,
  runProps?: ReadonlyArray<RunProps | undefined>,
): Paragraph {
  const children = texts.map((text, index) => createRun(text, runProps?.[index]))

  return Object.freeze({
    kind: 'paragraph',
    ...(props !== undefined ? { props } : {}),
    children: Object.freeze(children),
  }) satisfies Paragraph
}

function createRun(text: string, props?: RunProps): Run {
  return Object.freeze({
    kind: 'run',
    ...(props !== undefined ? { props } : {}),
    children: Object.freeze([
      Object.freeze({
        kind: 'text',
        value: text,
      }),
    ]),
  }) satisfies Run
}
