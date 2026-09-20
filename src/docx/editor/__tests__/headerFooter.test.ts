/** D29 — reading and rewriting header/footer text, paragraph by paragraph. */
import { describe, expect, it } from 'vitest'

import { applyCommand } from '../commands'
import {
  buildHeaderFooterSegmentEdit,
  buildHeaderFooterTextEdit,
  buildInsertHeaderFooterParagraph,
  buildRemoveHeaderFooterParagraph,
  listHeaderFooterParts,
  setHeaderFooterBlockText,
  setHeaderFooterSegmentText,
} from '../headerFooter'
import type { Block, Bookmark, Document, Drawing, Field, Header, Footer, Hyperlink, Paragraph, Run } from '../../model'

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

function textRun(text: string, bold = false): Run {
  return { kind: 'run', props: bold ? { bold: true } : {}, children: [{ kind: 'text', value: text }] }
}

function tabRun(): Run {
  return { kind: 'run', children: [{ kind: 'tab' }] }
}

function drawingRun(relationshipId: string): Run {
  const drawing: Drawing = { kind: 'drawing', layout: 'inline', relationshipId }
  return { kind: 'run', children: [drawing] }
}

function pageField(): Field {
  return {
    kind: 'field',
    fieldType: 'PAGE',
    instruction: 'PAGE',
    result: [],
    raw: '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>',
  }
}

/** "Chapter Title[tab]" + a PAGE field, with a logo drawing in front — the
 * canonical "mixed" paragraph this feature exists for. */
function chapterTitleWithPageNumberParagraph(): Paragraph {
  return {
    kind: 'paragraph',
    children: [drawingRun('rIdLogo'), textRun('Chapter Title'), tabRun(), pageField()],
  } as unknown as Paragraph
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

// ---------------------------------------------------------------------------
// D29 follow-up 2 — mixed paragraphs (text + a drawing/field/hyperlink/etc)
// ---------------------------------------------------------------------------

describe('listHeaderFooterParts — mixed rows', () => {
  it('"Chapter Title[tab] + PAGE field", with a logo drawing in front, is one mixed row: atom, text, atom', () => {
    const document = documentWith([chapterTitleWithPageNumberParagraph()], [paragraph('Page 1')])
    const rows = listHeaderFooterParts(document)[0].rows
    expect(rows).toEqual([
      {
        kind: 'mixed',
        blockIndex: 0,
        segments: [
          { kind: 'atom', label: '[Image]' },
          { kind: 'text', segmentIndex: 0, text: 'Chapter Title\t' },
          { kind: 'atom', label: '[Page number]' },
        ],
      },
    ])
  })

  it('labels NUMPAGES/DATE fields and a hyperlink by their display text', () => {
    const numPages: Field = { kind: 'field', fieldType: 'NUMPAGES', instruction: 'NUMPAGES', result: [] }
    const date: Field = { kind: 'field', fieldType: 'DATE', instruction: 'DATE', result: [] }
    const hyperlink: Hyperlink = { kind: 'hyperlink', relationshipId: 'rIdLink', children: [textRun('our site')] }
    const mixed: Paragraph = {
      kind: 'paragraph',
      children: [textRun('Page 1 of '), numPages, textRun(' — '), date, textRun(' — '), hyperlink],
    } as unknown as Paragraph

    const document = documentWith([mixed], [paragraph('Page 1')])
    const row = listHeaderFooterParts(document)[0].rows[0]
    expect(row.kind).toBe('mixed')
    expect(row).toMatchObject({
      segments: [
        { kind: 'text', segmentIndex: 0, text: 'Page 1 of ' },
        { kind: 'atom', label: '[Total pages]' },
        { kind: 'text', segmentIndex: 1, text: ' — ' },
        { kind: 'atom', label: '[Date]' },
        { kind: 'text', segmentIndex: 2, text: ' — ' },
        { kind: 'atom', label: '[Link: our site]' },
      ],
    })
  })

  it('a bookmark splits the text either side into separate segments but is never shown as its own chip', () => {
    const bookmarkStart: Bookmark = { kind: 'bookmark', id: '0', boundary: 'start', name: '_GoBack' }
    const bookmarkEnd: Bookmark = { kind: 'bookmark', id: '0', boundary: 'end' }
    const paragraphWithBookmark: Paragraph = {
      kind: 'paragraph',
      // A lone field keeps this paragraph from being fully-plain (so it's a
      // 'mixed' row, not a 'text' row) without being what's under test here.
      children: [textRun('before '), bookmarkStart, bookmarkEnd, textRun('after'), pageField()],
    } as unknown as Paragraph

    const document = documentWith([paragraphWithBookmark], [paragraph('Page 1')])
    const row = listHeaderFooterParts(document)[0].rows[0]
    expect(row).toMatchObject({
      kind: 'mixed',
      segments: [
        { kind: 'text', segmentIndex: 0, text: 'before ' },
        { kind: 'text', segmentIndex: 1, text: 'after' },
        { kind: 'atom', label: '[Page number]' },
      ],
    })
  })

  it('a run mixing plain text with a drawing is left as one atom, not split', () => {
    const drawing: Drawing = { kind: 'drawing', layout: 'inline', relationshipId: 'rId9' }
    const oddRun: Run = { kind: 'run', children: [{ kind: 'text', value: 'Logo' }, drawing] }
    const withOddRun: Paragraph = {
      kind: 'paragraph',
      children: [oddRun, textRun('Confidential'), pageField()],
    } as unknown as Paragraph

    const document = documentWith([withOddRun], [paragraph('Page 1')])
    const row = listHeaderFooterParts(document)[0].rows[0]
    expect(row).toMatchObject({
      kind: 'mixed',
      segments: [
        { kind: 'atom', label: 'Logo[Image]' },
        { kind: 'text', segmentIndex: 0, text: 'Confidential' },
        { kind: 'atom', label: '[Page number]' },
      ],
    })
  })
})

describe('buildHeaderFooterSegmentEdit / setHeaderFooterSegmentText', () => {
  it('rewrites one text segment, keeping the drawing and the PAGE field completely untouched', () => {
    const document = documentWith([chapterTitleWithPageNumberParagraph()], [paragraph('Page 1')])
    const originalParagraph = document.headers.get('rId4')!.blocks[0] as Paragraph

    const next = setHeaderFooterSegmentText(document, 'header', 'rId4', 0, 0, 'New Title\t')
    const rewritten = next.headers.get('rId4')!.blocks[0] as Paragraph

    // The drawing run and the field are the exact same objects — never rebuilt.
    expect(rewritten.children[0]).toBe(originalParagraph.children[0])
    expect(rewritten.children[rewritten.children.length - 1]).toBe(
      originalParagraph.children[originalParagraph.children.length - 1],
    )

    const rows = listHeaderFooterParts(next)[0].rows
    expect(rows).toEqual([
      {
        kind: 'mixed',
        blockIndex: 0,
        segments: [
          { kind: 'atom', label: '[Image]' },
          { kind: 'text', segmentIndex: 0, text: 'New Title\t' },
          { kind: 'atom', label: '[Page number]' },
        ],
      },
    ])
  })

  it('edits only the targeted segment when a paragraph has more than one text segment', () => {
    const twoSegmentParagraph: Paragraph = {
      kind: 'paragraph',
      children: [textRun('Page '), pageField(), textRun(' of '), { ...pageField(), fieldType: 'NUMPAGES' as const, instruction: 'NUMPAGES' }],
    } as unknown as Paragraph
    const document = documentWith([twoSegmentParagraph], [paragraph('Page 1')])

    const next = setHeaderFooterSegmentText(document, 'header', 'rId4', 0, 1, ' out of ')
    const rows = listHeaderFooterParts(next)[0].rows
    expect(rows[0]).toMatchObject({
      segments: [
        { kind: 'text', segmentIndex: 0, text: 'Page ' },
        { kind: 'atom' },
        { kind: 'text', segmentIndex: 1, text: ' out of ' },
        { kind: 'atom' },
      ],
    })
  })

  it('returns null for a fully plain paragraph (use buildHeaderFooterTextEdit for that), an out-of-range segment, a placeholder, or unchanged text', () => {
    const document = documentWith(
      [paragraph('Plain'), chapterTitleWithPageNumberParagraph(), imageParagraph()],
      [paragraph('Page 1')],
    )
    expect(buildHeaderFooterSegmentEdit(document, 'header', 'rId4', 0, 0, 'x')).toBeNull() // fully plain
    expect(buildHeaderFooterSegmentEdit(document, 'header', 'rId4', 1, 5, 'x')).toBeNull() // no such segment
    expect(buildHeaderFooterSegmentEdit(document, 'header', 'rId4', 2, 0, 'x')).toBeNull() // placeholder (no text at all)
    expect(buildHeaderFooterSegmentEdit(document, 'header', 'rId4', 1, 0, 'Chapter Title\t')).toBeNull() // unchanged
  })

  it('emptying every text segment around an atom leaves it a placeholder on the next read (no data loss — the atom itself is untouched)', () => {
    const document = documentWith([chapterTitleWithPageNumberParagraph()], [paragraph('Page 1')])
    const emptied = setHeaderFooterSegmentText(document, 'header', 'rId4', 0, 0, '')

    const rows = listHeaderFooterParts(emptied)[0].rows
    expect(rows).toEqual([{ kind: 'placeholder', blockIndex: 0, label: '[Image][Page number]' }])
    // The field is still the exact same object — nothing about it was rewritten.
    const remaining = emptied.headers.get('rId4')!.blocks[0] as Paragraph
    expect(remaining.children.some((child) => child.kind === 'field')).toBe(true)
  })
})

describe('the replace-header-footer-blocks command — segment edits', () => {
  it('applies and undoes a mixed-row segment edit as one step, through the normal command path', () => {
    const document = documentWith([chapterTitleWithPageNumberParagraph()], [paragraph('Page 1')])
    const command = buildHeaderFooterSegmentEdit(document, 'header', 'rId4', 0, 0, 'Renamed\t')!
    const applied = applyCommand(document, command)

    expect(listHeaderFooterParts(applied.document)[0].rows[0]).toMatchObject({
      segments: [{ kind: 'atom' }, { kind: 'text', segmentIndex: 0, text: 'Renamed\t' }, { kind: 'atom' }],
    })

    const undone = applyCommand(applied.document, applied.inverse)
    expect(undone.document.headers.get('rId4')!.blocks[0]).toBe(document.headers.get('rId4')!.blocks[0])
  })
})
