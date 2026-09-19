/** D29 — reading and rewriting header/footer text, paragraph by paragraph. */
import { describe, expect, it } from 'vitest'

import { applyCommand } from '../commands'
import {
  buildHeaderFooterTextEdit,
  buildInsertHeaderFooterParagraph,
  buildRemoveHeaderFooterParagraph,
  listHeaderFooterParts,
  setHeaderFooterBlockText,
} from '../headerFooter'
import type { Block, Document, Drawing, Field, Header, Footer, Paragraph, Run } from '../../model'

function paragraph(text: string, bold = false): Paragraph {
  return {
    kind: 'paragraph',
    props: { jc: 'center' },
    children: [{ kind: 'run', props: bold ? { bold: true } : {}, children: [{ kind: 'text', value: text }] }],
  } as unknown as Paragraph
}

function multiRunParagraph(...runs: ReadonlyArray<readonly [string, boolean]>): Paragraph {
  return {
    kind: 'paragraph',
    children: runs.map(([text, bold]): Run => ({ kind: 'run', props: { bold }, children: [{ kind: 'text', value: text }] })),
  } as unknown as Paragraph
}

function imageParagraph(): Paragraph {
  const drawing: Drawing = { kind: 'drawing', layout: 'inline', relationshipId: 'rId9' }
  return { kind: 'paragraph', children: [{ kind: 'run', children: [drawing] }] } as unknown as Paragraph
}

function pageFieldParagraph(): Paragraph {
  const field: Field = { kind: 'field', fieldType: 'PAGE', instruction: 'PAGE', result: [], raw: '<w:fldSimple w:instr="PAGE"/>' }
  return { kind: 'paragraph', children: [field] } as unknown as Paragraph
}

/** Concatenates a paragraph's own run text, the same way the panel's `text`
 * row does — used here to check the EDITED RESULT reads right, independent
 * of exactly how many runs the diff happened to split it into. */
function paragraphText(paragraph: Paragraph): string {
  return (paragraph.children as ReadonlyArray<Run>).map((run) => (run.children[0] as { value: string }).value).join('')
}

function documentWith(headerBlocks: ReadonlyArray<Block>, footerBlocks: ReadonlyArray<Block>): Document {
  const header: Header = { kind: 'header', id: 'rId4', blocks: headerBlocks }
  const footer: Footer = { kind: 'footer', id: 'rId5', blocks: footerBlocks }
  return {
    kind: 'document',
    sections: [
      {
        kind: 'section',
        props: {
          headerReference: [{ id: 'rId4', type: 'default' }],
          footerReference: [{ id: 'rId5', type: 'default' }],
        },
        blocks: [],
      },
    ],
    styles: new Map(),
    numbering: new Map(),
    comments: new Map(),
    footnotes: new Map(),
    endnotes: new Map(),
    headers: new Map([['rId4', header]]),
    footers: new Map([['rId5', footer]]),
  } as unknown as Document
}

describe('listHeaderFooterParts', () => {
  it('lists each referenced part once, one row per paragraph', () => {
    const document = documentWith([paragraph('Quarterly report')], [paragraph('Page 1')])
    expect(listHeaderFooterParts(document)).toEqual([
      { kind: 'header', id: 'rId4', type: 'default', rows: [{ kind: 'text', blockIndex: 0, text: 'Quarterly report' }] },
      { kind: 'footer', id: 'rId5', type: 'default', rows: [{ kind: 'text', blockIndex: 0, text: 'Page 1' }] },
    ])
  })

  it('flags a table block as a non-editable placeholder', () => {
    const table = { kind: 'table', rows: [] } as unknown as Block
    const document = documentWith([table], [paragraph('Page 1')])
    expect(listHeaderFooterParts(document)[0].rows).toEqual([{ kind: 'placeholder', blockIndex: 0, label: '[Table]' }])
  })

  it('flags a paragraph holding an image, and one holding a PAGE field, as placeholders', () => {
    const document = documentWith([imageParagraph(), pageFieldParagraph()], [paragraph('Page 1')])
    expect(listHeaderFooterParts(document)[0].rows).toEqual([
      { kind: 'placeholder', blockIndex: 0, label: '[Image]' },
      { kind: 'placeholder', blockIndex: 1, label: '[Page number]' },
    ])
  })
})

describe('buildHeaderFooterTextEdit / setHeaderFooterBlockText', () => {
  it('rewrites one paragraph in place, keeping its properties and its run formatting', () => {
    const document = documentWith([paragraph('Old title', true)], [paragraph('Page 1')])
    const next = setHeaderFooterBlockText(document, 'header', 'rId4', 0, 'New title')

    const rewritten = next.headers.get('rId4')!.blocks[0] as Paragraph
    expect(rewritten.props).toEqual({ jc: 'center' })
    expect(paragraphText(rewritten)).toBe('New title')
    // Every resulting run keeps the paragraph's original bold formatting,
    // however many runs the diff split the new text across.
    expect((rewritten.children as ReadonlyArray<Run>).every((run) => run.props?.bold === true)).toBe(true)
    // The footer, and the original header paragraph, are untouched.
    expect(next.footers).toBe(document.footers)
    expect((document.headers.get('rId4')!.blocks[0] as Paragraph).children[0]).toMatchObject({
      children: [{ kind: 'text', value: 'Old title' }],
    })
  })

  it('keeps an unedited run bit-for-bit and only rebuilds the run the edit actually touched', () => {
    // "Confidential Draft" as two runs, bold then plain — editing only the
    // second word must not disturb the bold run's formatting or identity.
    const document = documentWith([multiRunParagraph(['Confidential ', true], ['Draft', false])], [paragraph('Page 1')])
    const next = setHeaderFooterBlockText(document, 'header', 'rId4', 0, 'Confidential Final')

    const rewritten = next.headers.get('rId4')!.blocks[0] as Paragraph
    const originalFirstRun = (document.headers.get('rId4')!.blocks[0] as Paragraph).children[0]
    // The bold "Confidential " run is reused verbatim (same object).
    expect(rewritten.children[0]).toBe(originalFirstRun)
    expect(rewritten.children[1]).toMatchObject({ props: { bold: false }, children: [{ kind: 'text', value: 'Final' }] })
  })

  it('does not touch a placeholder block (image, field, or table)', () => {
    const document = documentWith([imageParagraph()], [paragraph('Page 1')])
    expect(setHeaderFooterBlockText(document, 'header', 'rId4', 0, 'New text')).toBe(document)
  })

  it('returns the same document for an unknown part, an out-of-range block, or unchanged text', () => {
    const document = documentWith([paragraph('Title')], [paragraph('Page 1')])
    expect(setHeaderFooterBlockText(document, 'header', 'nope', 0, 'x')).toBe(document)
    expect(setHeaderFooterBlockText(document, 'header', 'rId4', 5, 'x')).toBe(document)
    expect(setHeaderFooterBlockText(document, 'header', 'rId4', 0, 'Title')).toBe(document)
  })
})

describe('a header with an image, a PAGE field and bold text', () => {
  it('editing the text paragraph keeps the drawing and the field completely untouched', () => {
    const document = documentWith(
      [imageParagraph(), paragraph('Confidential', true), pageFieldParagraph()],
      [paragraph('Page 1')],
    )

    const next = setHeaderFooterBlockText(document, 'header', 'rId4', 1, 'CONFIDENTIAL')
    const blocks = next.headers.get('rId4')!.blocks

    expect(blocks[0]).toBe(document.headers.get('rId4')!.blocks[0]) // the image, untouched
    expect(blocks[2]).toBe(document.headers.get('rId4')!.blocks[2]) // the PAGE field, untouched
    const editedParagraph = blocks[1] as Paragraph
    expect(paragraphText(editedParagraph)).toBe('CONFIDENTIAL')
    expect((editedParagraph.children as ReadonlyArray<Run>).every((run) => run.props?.bold === true)).toBe(true)
  })
})

describe('buildInsertHeaderFooterParagraph / buildRemoveHeaderFooterParagraph', () => {
  it('appends a blank paragraph seeded from the last text paragraph, and removes it again', () => {
    const document = documentWith([paragraph('Title')], [paragraph('Page 1')])
    const insertCommand = buildInsertHeaderFooterParagraph(document, 'header', 'rId4')!
    const inserted = applyCommand(document, insertCommand)
    expect(inserted.document.headers.get('rId4')!.blocks).toHaveLength(2)
    expect((inserted.document.headers.get('rId4')!.blocks[1] as Paragraph).props).toEqual({ jc: 'center' })

    const removeCommand = buildRemoveHeaderFooterParagraph(inserted.document, 'header', 'rId4', 1)!
    const removed = applyCommand(inserted.document, removeCommand)
    expect(removed.document.headers.get('rId4')!.blocks).toHaveLength(1)
  })

  it('refuses to remove the only remaining block', () => {
    const document = documentWith([paragraph('Title')], [paragraph('Page 1')])
    expect(buildRemoveHeaderFooterParagraph(document, 'header', 'rId4', 0)).toBeNull()
  })
})

describe('the replace-header-footer-blocks command', () => {
  it('applies and undoes a text edit through the normal command path', () => {
    const document = documentWith([paragraph('Before')], [paragraph('Page 1')])
    const command = buildHeaderFooterTextEdit(document, 'header', 'rId4', 0, 'After')!
    const applied = applyCommand(document, command)
    expect((applied.document.headers.get('rId4')!.blocks[0] as Paragraph).children[0]).toMatchObject({
      children: [{ kind: 'text', value: 'After' }],
    })

    const undone = applyCommand(applied.document, applied.inverse)
    expect((undone.document.headers.get('rId4')!.blocks[0] as Paragraph).children[0]).toMatchObject({
      children: [{ kind: 'text', value: 'Before' }],
    })
  })

  it('applies and undoes a paragraph insert/remove as one step each', () => {
    const document = documentWith([paragraph('Title')], [paragraph('Page 1')])
    const insertCommand = buildInsertHeaderFooterParagraph(document, 'header', 'rId4')!
    const applied = applyCommand(document, insertCommand)
    expect(applied.document.headers.get('rId4')!.blocks).toHaveLength(2)

    const undone = applyCommand(applied.document, applied.inverse)
    expect(undone.document.headers.get('rId4')!.blocks).toHaveLength(1)
    expect(undone.document.headers.get('rId4')!.blocks[0]).toBe(document.headers.get('rId4')!.blocks[0])
  })
})
