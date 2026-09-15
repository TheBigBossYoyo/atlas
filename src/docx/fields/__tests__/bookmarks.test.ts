import { describe, expect, it } from 'vitest'

import type { Document, Paragraph, Section } from '../../model'
import { collectBookmarkMaps } from '../bookmarks'
import type { TocPageResolver } from '../toc'

function makeDocument(sections: ReadonlyArray<Section>, overrides: Partial<Document> = {}): Document {
  return {
    kind: 'document',
    sections,
    styles: new Map(),
    numbering: new Map(),
    comments: new Map(),
    footnotes: new Map(),
    endnotes: new Map(),
    headers: new Map(),
    footers: new Map(),
    ...overrides,
  }
}

function text(value: string): Paragraph['children'][number] {
  return { kind: 'run', children: [{ kind: 'text', value }] }
}

function section(blocks: Section['blocks']): Section {
  return { kind: 'section', props: {}, blocks }
}

describe('collectBookmarkMaps', () => {
  it('collects a bookmark\'s plain text between its start and end within one paragraph', () => {
    const document = makeDocument([
      section([
        {
          kind: 'paragraph',
          children: [
            text('See '),
            { kind: 'bookmark', id: '1', boundary: 'start', name: '_Ref1' },
            text('Section One'),
            { kind: 'bookmark', id: '1', boundary: 'end' },
            text(' for details.'),
          ],
        },
      ]),
    ])

    const { bookmarkText } = collectBookmarkMaps(document)
    expect(bookmarkText.get('_Ref1')).toBe('Section One')
  })

  it('accumulates text across a bookmark spanning multiple paragraphs', () => {
    const document = makeDocument([
      section([
        {
          kind: 'paragraph',
          children: [
            { kind: 'bookmark', id: '1', boundary: 'start', name: 'Spanning' },
            text('First'),
          ],
        },
        {
          kind: 'paragraph',
          children: [text('Second'), { kind: 'bookmark', id: '1', boundary: 'end' }, text('Third')],
        },
      ]),
    ])

    const { bookmarkText } = collectBookmarkMaps(document)
    expect(bookmarkText.get('Spanning')).toBe('FirstSecond')
  })

  it('collects text from a bookmark nested inside a hyperlink', () => {
    const document = makeDocument([
      section([
        {
          kind: 'paragraph',
          children: [
            {
              kind: 'hyperlink',
              relationshipId: 'rId1',
              children: [
                { kind: 'bookmark', id: '1', boundary: 'start', name: 'InLink' },
                { kind: 'run', children: [{ kind: 'text', value: 'Linked text' }] },
                { kind: 'bookmark', id: '1', boundary: 'end' },
              ],
            },
          ],
        },
      ]),
    ])

    const { bookmarkText } = collectBookmarkMaps(document)
    expect(bookmarkText.get('InLink')).toBe('Linked text')
  })

  it('includes a nested field\'s cached result as the bookmark\'s text', () => {
    const document = makeDocument([
      section([
        {
          kind: 'paragraph',
          children: [
            { kind: 'bookmark', id: '1', boundary: 'start', name: 'WithField' },
            {
              kind: 'field',
              fieldType: 'unknown',
              instruction: 'MERGEFIELD X',
              result: [{ kind: 'run', children: [{ kind: 'text', value: 'merged' }] }],
            },
            { kind: 'bookmark', id: '1', boundary: 'end' },
          ],
        },
      ]),
    ])

    const { bookmarkText } = collectBookmarkMaps(document)
    expect(bookmarkText.get('WithField')).toBe('merged')
  })

  it('resolves a body bookmark\'s PAGEREF page from the supplied page resolver, keyed on the start boundary\'s position', () => {
    const document = makeDocument([
      section([
        { kind: 'paragraph', children: [text('Intro')] },
        {
          kind: 'paragraph',
          children: [
            { kind: 'bookmark', id: '1', boundary: 'start', name: '_Ref1' },
            text('Target'),
            { kind: 'bookmark', id: '1', boundary: 'end' },
          ],
        },
      ]),
    ])

    const pageOf: TocPageResolver = (sectionIndex, blockIndex) =>
      sectionIndex === 0 && blockIndex === 1 ? 3 : undefined

    const { bookmarkText, bookmarkPage } = collectBookmarkMaps(document, pageOf)
    expect(bookmarkText.get('_Ref1')).toBe('Target')
    expect(bookmarkPage.get('_Ref1')).toBe(3)
  })

  it('gives a bookmark inside a table cell a text entry but no page entry', () => {
    const document = makeDocument([
      section([
        {
          kind: 'table',
          tblGrid: [],
          rows: [
            {
              kind: 'table-row',
              cells: [
                {
                  kind: 'table-cell',
                  blocks: [
                    {
                      kind: 'paragraph',
                      children: [
                        { kind: 'bookmark', id: '1', boundary: 'start', name: 'InCell' },
                        text('Cell text'),
                        { kind: 'bookmark', id: '1', boundary: 'end' },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ]),
    ])

    const pageOf: TocPageResolver = () => 7
    const { bookmarkText, bookmarkPage } = collectBookmarkMaps(document, pageOf)

    expect(bookmarkText.get('InCell')).toBe('Cell text')
    expect(bookmarkPage.has('InCell')).toBe(false)
  })

  it('collects text from a bookmark inside a header, with no page entry', () => {
    const document = makeDocument([section([])], {
      headers: new Map([
        [
          'header1',
          {
            kind: 'header',
            id: 'header1',
            blocks: [
              {
                kind: 'paragraph',
                children: [
                  { kind: 'bookmark', id: '1', boundary: 'start', name: 'InHeader' },
                  text('Header text'),
                  { kind: 'bookmark', id: '1', boundary: 'end' },
                ],
              },
            ],
          },
        ],
      ]),
    })

    const { bookmarkText, bookmarkPage } = collectBookmarkMaps(document, () => 5)
    expect(bookmarkText.get('InHeader')).toBe('Header text')
    expect(bookmarkPage.has('InHeader')).toBe(false)
  })

  it('omits a bookmark that never closes (no matching end)', () => {
    const document = makeDocument([
      section([
        {
          kind: 'paragraph',
          children: [{ kind: 'bookmark', id: '1', boundary: 'start', name: 'Unclosed' }, text('Text')],
        },
      ]),
    ])

    const { bookmarkText } = collectBookmarkMaps(document)
    expect(bookmarkText.has('Unclosed')).toBe(false)
  })

  it('keeps two independent bookmarks in the same paragraph separate', () => {
    const document = makeDocument([
      section([
        {
          kind: 'paragraph',
          children: [
            { kind: 'bookmark', id: '1', boundary: 'start', name: 'First' },
            text('one'),
            { kind: 'bookmark', id: '1', boundary: 'end' },
            text(' between '),
            { kind: 'bookmark', id: '2', boundary: 'start', name: 'Second' },
            text('two'),
            { kind: 'bookmark', id: '2', boundary: 'end' },
          ],
        },
      ]),
    ])

    const { bookmarkText } = collectBookmarkMaps(document)
    expect(bookmarkText.get('First')).toBe('one')
    expect(bookmarkText.get('Second')).toBe('two')
  })

  it('handles an inner bookmark nested fully within an outer bookmark, feeding text to both', () => {
    const document = makeDocument([
      section([
        {
          kind: 'paragraph',
          children: [
            { kind: 'bookmark', id: '1', boundary: 'start', name: 'Outer' },
            text('before '),
            { kind: 'bookmark', id: '2', boundary: 'start', name: 'Inner' },
            text('middle'),
            { kind: 'bookmark', id: '2', boundary: 'end' },
            text(' after'),
            { kind: 'bookmark', id: '1', boundary: 'end' },
          ],
        },
      ]),
    ])

    const { bookmarkText } = collectBookmarkMaps(document)
    expect(bookmarkText.get('Inner')).toBe('middle')
    expect(bookmarkText.get('Outer')).toBe('before middle after')
  })
})
