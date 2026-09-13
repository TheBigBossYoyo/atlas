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
    expect(result.inverse).toEqual({
      kind: 'insert-text',
      at: position([0], 0, 1),
      text: 'tla',
    })
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
