import { describe, expect, it } from 'vitest'

import type { Document, Field, Paragraph, Section, Style } from '../../model'
import { collectTocEntries, parseTocOptions, renderTocFieldResult, updateTableOfContents } from '../toc'

function makeStyles(styles: ReadonlyArray<Style>): ReadonlyMap<string, Style> {
  return new Map(styles.map((style) => [style.id, style]))
}

const HEADING_STYLES = makeStyles([
  { id: 'Heading1', type: 'paragraph', paragraph: { outlineLvl: 0 } },
  { id: 'Heading2', type: 'paragraph', paragraph: { outlineLvl: 1 } },
  { id: 'Heading3', type: 'paragraph', paragraph: { outlineLvl: 2 } },
])

function headingParagraph(text: string, styleId: string): Paragraph {
  return {
    kind: 'paragraph',
    props: { pStyle: styleId },
    children: [{ kind: 'run', children: [{ kind: 'text', value: text }] }],
  }
}

function bodyParagraph(text: string): Paragraph {
  return { kind: 'paragraph', children: [{ kind: 'run', children: [{ kind: 'text', value: text }] }] }
}

function makeDocument(sections: ReadonlyArray<Section>, styles = HEADING_STYLES): Document {
  return {
    kind: 'document',
    sections,
    styles,
    numbering: new Map(),
    comments: new Map(),
    footnotes: new Map(),
    endnotes: new Map(),
    headers: new Map(),
    footers: new Map(),
  }
}

describe('parseTocOptions', () => {
  it('defaults to levels 1-3, no hyperlinks, when no switches are present', () => {
    expect(parseTocOptions('TOC')).toEqual({ minLevel: 1, maxLevel: 3, hyperlink: false })
  })

  it('parses an explicit \\o "min-max" level range', () => {
    expect(parseTocOptions('TOC \\o "2-4"')).toMatchObject({ minLevel: 2, maxLevel: 4 })
  })

  it('parses a single-number \\o as 1..N', () => {
    expect(parseTocOptions('TOC \\o "5"')).toMatchObject({ minLevel: 1, maxLevel: 5 })
  })

  it('sets hyperlink true when \\h is present', () => {
    expect(parseTocOptions('TOC \\o "1-3" \\h \\z \\u').hyperlink).toBe(true)
  })
})

describe('collectTocEntries', () => {
  it('collects headings within the level range, in document order, with their text', () => {
    const document = makeDocument([
      {
        kind: 'section',
        props: {},
        blocks: [
          headingParagraph('Introduction', 'Heading1'),
          bodyParagraph('Some body text.'),
          headingParagraph('Background', 'Heading2'),
          headingParagraph('Conclusion', 'Heading1'),
        ],
      },
    ])

    const entries = collectTocEntries(document, { minLevel: 1, maxLevel: 3, hyperlink: false })

    expect(entries).toEqual([
      { level: 1, text: 'Introduction' },
      { level: 2, text: 'Background' },
      { level: 1, text: 'Conclusion' },
    ])
  })

  it('excludes headings outside the requested level range', () => {
    const document = makeDocument([
      {
        kind: 'section',
        props: {},
        blocks: [
          headingParagraph('Chapter', 'Heading1'),
          headingParagraph('Section', 'Heading2'),
          headingParagraph('Subsection', 'Heading3'),
        ],
      },
    ])

    const entries = collectTocEntries(document, { minLevel: 1, maxLevel: 2, hyperlink: false })

    expect(entries.map((entry) => entry.text)).toEqual(['Chapter', 'Section'])
  })

  it('treats a paragraph with a direct outlineLvl (no heading style) as a heading too', () => {
    const document = makeDocument([
      {
        kind: 'section',
        props: {},
        blocks: [
          { kind: 'paragraph', props: { outlineLvl: 0 }, children: [{ kind: 'run', children: [{ kind: 'text', value: 'Direct heading' }] }] },
        ],
      },
    ])

    expect(collectTocEntries(document, { minLevel: 1, maxLevel: 3, hyperlink: false })).toEqual([
      { level: 1, text: 'Direct heading' },
    ])
  })

  it('skips a heading paragraph with no visible text', () => {
    const document = makeDocument([
      { kind: 'section', props: {}, blocks: [{ kind: 'paragraph', props: { pStyle: 'Heading1' }, children: [] }] },
    ])
    expect(collectTocEntries(document, { minLevel: 1, maxLevel: 3, hyperlink: false })).toEqual([])
  })

  it('includes the heading\'s own leading bookmark name when present', () => {
    const paragraph: Paragraph = {
      kind: 'paragraph',
      props: { pStyle: 'Heading1' },
      children: [
        { kind: 'bookmark', id: '0', boundary: 'start', name: '_Toc1' },
        { kind: 'run', children: [{ kind: 'text', value: 'Bookmarked heading' }] },
        { kind: 'bookmark', id: '0', boundary: 'end' },
      ],
    }
    const document = makeDocument([{ kind: 'section', props: {}, blocks: [paragraph] }])

    expect(collectTocEntries(document, { minLevel: 1, maxLevel: 3, hyperlink: false })).toEqual([
      { level: 1, text: 'Bookmarked heading', bookmarkName: '_Toc1' },
    ])
  })

  it('attaches a page number via the supplied pageOf resolver', () => {
    const document = makeDocument([
      { kind: 'section', props: {}, blocks: [headingParagraph('Intro', 'Heading1')] },
    ])

    const entries = collectTocEntries(
      document,
      { minLevel: 1, maxLevel: 3, hyperlink: false },
      (sectionIndex, blockIndex) => (sectionIndex === 0 && blockIndex === 0 ? 2 : undefined),
    )

    expect(entries).toEqual([{ level: 1, text: 'Intro', pageNumber: 2 }])
  })
})

describe('renderTocFieldResult', () => {
  it('renders one line per entry with a tab before the page number, joined by line breaks', () => {
    const rendered = renderTocFieldResult(
      [
        { level: 1, text: 'Introduction', pageNumber: 1 },
        { level: 1, text: 'Conclusion', pageNumber: 5 },
      ],
      { minLevel: 1, maxLevel: 3, hyperlink: false },
    )

    expect(rendered).toEqual([
      { kind: 'run', children: [{ kind: 'text', value: 'Introduction' }, { kind: 'tab' }, { kind: 'text', value: '1' }] },
      { kind: 'run', children: [{ kind: 'break', breakType: 'line' }] },
      { kind: 'run', children: [{ kind: 'text', value: 'Conclusion' }, { kind: 'tab' }, { kind: 'text', value: '5' }] },
    ])
  })

  it('indents a deeper-level entry with leading tabs', () => {
    const rendered = renderTocFieldResult(
      [{ level: 2, text: 'Sub-topic', pageNumber: 3 }],
      { minLevel: 1, maxLevel: 3, hyperlink: false },
    )
    expect(rendered[0]).toMatchObject({
      children: [{ kind: 'text', value: '\tSub-topic' }, { kind: 'tab' }, { kind: 'text', value: '3' }],
    })
  })

  it('leaves the page-number text empty when no page is known', () => {
    const rendered = renderTocFieldResult([{ level: 1, text: 'Intro' }], { minLevel: 1, maxLevel: 3, hyperlink: false })
    expect(rendered[0]).toMatchObject({ children: [{ kind: 'text', value: 'Intro' }, { kind: 'tab' }, { kind: 'text', value: '' }] })
  })

  it('wraps an entry in a hyperlink to its bookmark when hyperlink is requested and a bookmark exists', () => {
    const rendered = renderTocFieldResult(
      [{ level: 1, text: 'Intro', bookmarkName: '_Toc1', pageNumber: 1 }],
      { minLevel: 1, maxLevel: 3, hyperlink: true },
    )
    expect(rendered[0]).toMatchObject({ kind: 'hyperlink', anchor: '_Toc1' })
  })

  it('renders plain text (no hyperlink) when hyperlink is requested but the heading has no bookmark', () => {
    const rendered = renderTocFieldResult([{ level: 1, text: 'Intro', pageNumber: 1 }], { minLevel: 1, maxLevel: 3, hyperlink: true })
    expect(rendered[0].kind).toBe('run')
  })
})

describe('updateTableOfContents', () => {
  function makeTocField(instruction = 'TOC \\o "1-3" \\h \\z \\u'): Field {
    return { kind: 'field', fieldType: 'TOC', instruction, result: [], raw: '<w:r>...</w:r>' }
  }

  it('returns updated: false and the document unchanged when there is no TOC field', () => {
    const document = makeDocument([{ kind: 'section', props: {}, blocks: [headingParagraph('Intro', 'Heading1')] }])
    const result = updateTableOfContents(document)
    expect(result.updated).toBe(false)
    expect(result.document).toBe(document)
  })

  it('returns updated: false and leaves a locked (w:fldLock) TOC field untouched, matching Word\'s own "Update Table of Contents"', () => {
    const lockedField: Field = { ...makeTocField(), locked: true }
    const document = makeDocument([
      {
        kind: 'section',
        props: {},
        blocks: [
          { kind: 'paragraph', children: [lockedField] },
          headingParagraph('Introduction', 'Heading1'),
        ],
      },
    ])

    const result = updateTableOfContents(document)

    expect(result.updated).toBe(false)
    expect(result.document).toBe(document)
  })

  it('replaces the TOC field\'s result with freshly generated entries and clears raw', () => {
    const document = makeDocument([
      {
        kind: 'section',
        props: {},
        blocks: [
          { kind: 'paragraph', children: [makeTocField()] },
          headingParagraph('Introduction', 'Heading1'),
          headingParagraph('Background', 'Heading2'),
        ],
      },
    ])

    const result = updateTableOfContents(document)
    expect(result.updated).toBe(true)

    const tocField = result.document.sections[0].blocks[0]
    expect(tocField.kind === 'paragraph' && tocField.children[0]).toMatchObject({
      kind: 'field',
      fieldType: 'TOC',
      raw: undefined,
    })
    const field = tocField.kind === 'paragraph' ? tocField.children[0] : undefined
    expect(field?.kind === 'field' ? field.result.length : 0).toBeGreaterThan(0)
  })

  it('passes a pageOf resolver through to the generated entries', () => {
    const document = makeDocument([
      {
        kind: 'section',
        props: {},
        blocks: [{ kind: 'paragraph', children: [makeTocField()] }, headingParagraph('Intro', 'Heading1')],
      },
    ])

    const result = updateTableOfContents(document, (sectionIndex, blockIndex) =>
      sectionIndex === 0 && blockIndex === 1 ? 7 : undefined,
    )

    const block = result.document.sections[0].blocks[0]
    const field = block.kind === 'paragraph' ? block.children[0] : undefined
    const runChildren = field?.kind === 'field' ? field.result[0] : undefined
    expect(runChildren?.kind === 'run' && runChildren.children.some((c) => c.kind === 'text' && c.value === '7')).toBe(true)
  })
})
