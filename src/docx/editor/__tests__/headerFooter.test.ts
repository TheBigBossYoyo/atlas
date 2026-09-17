/** D29 — reading and rewriting header/footer text. */
import { describe, expect, it } from 'vitest'

import { applyCommand } from '../commands'
import { blocksToText, listHeaderFooterParts, setHeaderFooterText } from '../headerFooter'
import type { Block, Document, Header, Footer, Paragraph } from '../../model'

function paragraph(text: string, bold = false): Paragraph {
  return {
    kind: 'paragraph',
    props: { jc: 'center' },
    children: [{ kind: 'run', props: bold ? { b: true } : {}, children: [{ kind: 'text', value: text }] }],
  } as unknown as Paragraph
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
  it('lists each referenced part once, with its current text', () => {
    const document = documentWith([paragraph('Quarterly report')], [paragraph('Page 1')])
    expect(listHeaderFooterParts(document)).toEqual([
      { kind: 'header', id: 'rId4', type: 'default', text: 'Quarterly report', hasRichContent: false },
      { kind: 'footer', id: 'rId5', type: 'default', text: 'Page 1', hasRichContent: false },
    ])
  })

  it('flags a part that holds more than plain text', () => {
    const table = { kind: 'table', rows: [] } as unknown as Block
    const document = documentWith([table], [paragraph('Page 1')])
    expect(listHeaderFooterParts(document)[0].hasRichContent).toBe(true)
  })
})

describe('setHeaderFooterText', () => {
  it('rewrites the text while keeping paragraph and run formatting', () => {
    const document = documentWith([paragraph('Old title', true)], [paragraph('Page 1')])
    const next = setHeaderFooterText(document, 'header', 'rId4', 'New title\nSecond line')

    const blocks = next.headers.get('rId4')!.blocks as ReadonlyArray<Paragraph>
    expect(blocksToText(blocks)).toBe('New title\nSecond line')
    expect(blocks[0].props).toEqual({ jc: 'center' })
    expect(blocks[0].children[0]).toMatchObject({ kind: 'run', props: { b: true } })
    // The second line inherits the formatting of the paragraph before it.
    expect(blocks[1].children[0]).toMatchObject({ kind: 'run', props: { b: true } })
    // The footer is untouched, and so is the original document.
    expect(next.footers).toBe(document.footers)
    expect(blocksToText(document.headers.get('rId4')!.blocks)).toBe('Old title')
  })

  it('returns the same document for an unknown part or unchanged text', () => {
    const document = documentWith([paragraph('Title')], [paragraph('Page 1')])
    expect(setHeaderFooterText(document, 'header', 'nope', 'x')).toBe(document)
    expect(setHeaderFooterText(document, 'header', 'rId4', 'Title')).toBe(document)
  })
})

describe('the set-header-footer-text command', () => {
  it('applies and undoes through the normal command path', () => {
    const document = documentWith([paragraph('Before')], [paragraph('Page 1')])
    const applied = applyCommand(document, { kind: 'set-header-footer-text', target: 'header', id: 'rId4', text: 'After' })
    expect(blocksToText(applied.document.headers.get('rId4')!.blocks)).toBe('After')

    const undone = applyCommand(applied.document, applied.inverse)
    expect(blocksToText(undone.document.headers.get('rId4')!.blocks)).toBe('Before')
  })
})
