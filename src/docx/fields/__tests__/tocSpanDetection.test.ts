/**
 * DEFER-5 — telling the truth about a table of contents Atlas cannot regenerate.
 *
 * `updateTableOfContents` reported `updated: false` for three different
 * situations and the UI said "No table of contents found to update" for all of
 * them. On a real Word document that is simply untrue: the document HAS a table
 * of contents. Word writes it as a field whose `fldChar begin` sits in one
 * paragraph and whose `end` sits in a later one — OOXML requires that, since a
 * `w:r` cannot contain a `w:p` — and Atlas's parser only groups a complex field
 * within a single paragraph, so the field never becomes a `Field` node for
 * `findTocField` to see.
 *
 * Regenerating such a field is a much larger change (it spans paragraphs, so
 * "update" means replacing a range of blocks). Until then, the useful thing is
 * to say which of the three cases happened, so a user is not left doubting
 * their file.
 *
 * The document builder mirrors `toc.test.ts`'s own, so both tests agree about
 * what a minimal document looks like.
 */
import { describe, expect, it } from 'vitest'

import { updateTableOfContents } from '../toc'
import type { Document, Paragraph, Section, Style } from '../../model'

function makeStyles(styles: ReadonlyArray<Style>): ReadonlyMap<string, Style> {
  return new Map(styles.map((style) => [style.id, style]))
}

const HEADING_STYLES = makeStyles([{ id: 'Heading1', type: 'paragraph', paragraph: { outlineLvl: 0 } }])

function textParagraph(text: string, styleId?: string): Paragraph {
  return {
    kind: 'paragraph',
    ...(styleId ? { props: { pStyle: styleId } } : {}),
    children: [{ kind: 'run', children: [{ kind: 'text', value: text }] }],
  }
}

function makeDocument(blocks: Section['blocks']): Document {
  return {
    kind: 'document',
    sections: [{ kind: 'section', props: {}, blocks }],
    styles: HEADING_STYLES,
    numbering: new Map(),
    comments: new Map(),
    footnotes: new Map(),
    endnotes: new Map(),
    headers: new Map(),
    footers: new Map(),
  }
}

describe('why a table of contents was not regenerated (DEFER-5)', () => {
  it('reports "none" for a document with no table of contents at all', () => {
    const result = updateTableOfContents(makeDocument([textParagraph('Just a heading', 'Heading1')]))
    expect(result.updated).toBe(false)
    expect(result.reason).toBe('none')
  })

  it('reports "unsupported-span" when the TOC instruction survives as run text', () => {
    // What a Word-written TOC leaves behind once the parser fails to group it:
    // the instruction text sitting in an ordinary run.
    const result = updateTableOfContents(
      makeDocument([
        textParagraph('TOC \\o "1-3" \\h \\z \\u'),
        textParagraph('Introduction\t1'),
        textParagraph('Chapter one\t4'),
      ]),
    )
    expect(result.updated).toBe(false)
    expect(result.reason).toBe('unsupported-span')
  })

  it('does not mistake a heading that merely says "TOC" for a field instruction', () => {
    // The instruction is matched as `TOC` followed by a switch, so prose is
    // safe. Without that, a document with a heading called "TOC" would get a
    // message about a table of contents it does not have.
    const result = updateTableOfContents(
      makeDocument([
        textParagraph('TOC', 'Heading1'),
        textParagraph('This chapter explains the TOC and how to use it.'),
      ]),
    )
    expect(result.reason).toBe('none')
  })

  it('leaves the document untouched in every case', () => {
    // The message is the only thing that changed; nothing here may edit the
    // document, least of all a TOC it admits it cannot regenerate.
    const document = makeDocument([textParagraph('TOC \\o "1-3"')])
    const result = updateTableOfContents(document)
    expect(result.document).toBe(document)
  })
})
