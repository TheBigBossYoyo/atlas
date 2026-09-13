/**
 * Tests for src/docx/editor/comments.ts (Wave E.2)
 */

import { describe, it, expect } from 'vitest'

import type {
  Block,
  Comment as CommentNode,
  Document as DocxDocument,
  Paragraph,
  ParagraphChild,
  Section,
  Style,
  NumberingDef,
  Footnote,
  Endnote,
  Header,
  Footer,
  UnknownNode,
} from '../../model/document'
import { extractCommentText, findCommentAnchors } from '../comments'
import { parseComments } from '../../parser/comments'

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeUnknownXmlNode(rawJson: string): UnknownNode {
  return { kind: 'unknown', xml: rawJson }
}

function makeCommentWithRawXml(id: string, rawJson: string): CommentNode {
  // The structured Comment.body is Paragraph[], but extractCommentText also
  // accepts a legacy `blocks` field containing UnknownNode payloads (Wave A.5
  // raw-XML storage).  Cast through unknown to attach that legacy field for
  // the raw-XML extraction tests.
  return {
    kind: 'comment',
    id,
    body: [],
    blocks: [makeUnknownXmlNode(rawJson)],
  } as unknown as CommentNode
}

function makeParagraph(children: ReadonlyArray<ParagraphChild>): Paragraph {
  return { kind: 'paragraph', children }
}

function makeSection(blocks: ReadonlyArray<Block>): Section {
  return {
    kind: 'section',
    props: {},
    blocks,
  }
}

function makeDocument(
  sections: ReadonlyArray<Section>,
  comments: ReadonlyMap<string, CommentNode> = new Map(),
): DocxDocument {
  return {
    kind: 'document',
    sections,
    styles: new Map<string, Style>(),
    numbering: new Map<string, NumberingDef>(),
    comments,
    footnotes: new Map<string, Footnote>(),
    endnotes: new Map<string, Endnote>(),
    headers: new Map<string, Header>(),
    footers: new Map<string, Footer>(),
  }
}

// ---------------------------------------------------------------------------
// extractCommentText
// ---------------------------------------------------------------------------

describe('extractCommentText', () => {
  it('extracts text from a single-paragraph w:t in raw JSON payload', () => {
    const raw = JSON.stringify({
      '@_w:id': '1',
      '@_w:author': 'Alice',
      'w:p': { 'w:r': { 'w:t': 'This needs revision.' } },
    })
    const comment = makeCommentWithRawXml('1', raw)

    expect(extractCommentText(comment)).toBe('This needs revision.')
  })

  it('joins text from multiple paragraphs with a single space', () => {
    const raw = JSON.stringify({
      '@_w:id': '2',
      'w:p': [
        { 'w:r': { 'w:t': 'First line.' } },
        { 'w:r': { 'w:t': 'Second line.' } },
      ],
    })
    const comment = makeCommentWithRawXml('2', raw)

    expect(extractCommentText(comment)).toBe('First line. Second line.')
  })

  it('skips w:delText and w:instrText', () => {
    const raw = JSON.stringify({
      '@_w:id': '3',
      'w:p': {
        'w:r': [
          { 'w:t': 'Keep this.' },
          { 'w:delText': 'Drop this.' },
          { 'w:instrText': 'PAGE' },
        ],
      },
    })
    const comment = makeCommentWithRawXml('3', raw)

    expect(extractCommentText(comment)).toContain('Keep this.')
    expect(extractCommentText(comment)).not.toContain('Drop this.')
    expect(extractCommentText(comment)).not.toContain('PAGE')
  })

  it('returns empty string when payload is malformed JSON', () => {
    const comment = makeCommentWithRawXml('4', '<<< invalid')
    expect(extractCommentText(comment)).toBe('')
  })

  it('returns empty string when comment has no body', () => {
    const comment: CommentNode = { kind: 'comment', id: '5', body: [] }
    expect(extractCommentText(comment)).toBe('')
  })

  it('handles structured Block paragraphs as well as UnknownNode payloads', () => {
    const comment: CommentNode = {
      kind: 'comment',
      id: '6',
      body: [
        makeParagraph([
          {
            kind: 'run',
            children: [{ kind: 'text', value: 'Structured comment.' }],
          },
        ]),
      ],
    }

    expect(extractCommentText(comment)).toBe('Structured comment.')
  })

  it('extracts text from a real parsed comment containing a hyperlink', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="11" w:author="Alice">
    <w:p>
      <w:hyperlink r:id="rId1">
        <w:r><w:t>See the style guide</w:t></w:r>
      </w:hyperlink>
    </w:p>
  </w:comment>
</w:comments>`

    const parsed = parseComments(xml)
    const comment = parsed.get('11')
    expect(comment).toBeDefined()
    expect(comment?.body[0]?.kind).toBe('paragraph')
    expect((comment?.body[0] as Paragraph).children[0]?.kind).toBe('hyperlink')
    expect(extractCommentText(comment as CommentNode)).toBe('See the style guide')
  })

  it('extracts correctly ordered, correctly spaced text from a real comment with a hyperlink sandwiched between runs', () => {
    // Companion to the parser-level regression test in
    // src/docx/parser/__tests__/comments.test.ts: this exercises the same
    // "run, hyperlink, run" shape end-to-end through extractCommentText to
    // confirm the fix produces the right *displayed* text, not just the
    // right AST shape.
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="12" w:author="Erin"><w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:hyperlink r:id="rId1"><w:r><w:t>the guide</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> for details.</w:t></w:r></w:p></w:comment></w:comments>`

    const parsed = parseComments(xml)
    const comment = parsed.get('12')
    expect(comment).toBeDefined()
    expect(extractCommentText(comment as CommentNode)).toBe('See the guide for details.')
  })
})

// ---------------------------------------------------------------------------
// findCommentAnchors
// ---------------------------------------------------------------------------

describe('findCommentAnchors', () => {
  it('returns empty map when document has no anchors', () => {
    const document = makeDocument([
      makeSection([
        makeParagraph([
          { kind: 'run', children: [{ kind: 'text', value: 'Hello' }] },
        ]),
      ]),
    ])

    expect(findCommentAnchors(document).size).toBe(0)
  })

  it('finds a comment-range start anchor at the correct paragraph index', () => {
    const document = makeDocument([
      makeSection([
        makeParagraph([
          { kind: 'run', children: [{ kind: 'text', value: 'first' }] },
        ]),
        makeParagraph([
          { kind: 'comment-range', id: '7', boundary: 'start' },
          { kind: 'run', children: [{ kind: 'text', value: 'middle' }] },
          { kind: 'comment-range', id: '7', boundary: 'end' },
        ]),
        makeParagraph([
          { kind: 'run', children: [{ kind: 'text', value: 'last' }] },
        ]),
      ]),
    ])

    const anchors = findCommentAnchors(document)
    expect(anchors.get('7')).toBe(1)
  })

  it('finds a nested comment-reference inside a run', () => {
    const document = makeDocument([
      makeSection([
        makeParagraph([
          {
            kind: 'run',
            children: [
              { kind: 'text', value: 'before' },
              { kind: 'comment-reference', id: '8' },
            ],
          },
        ]),
      ]),
    ])

    expect(findCommentAnchors(document).get('8')).toBe(0)
  })

  it('prefers a comment-range start over a later reference for the same id', () => {
    const document = makeDocument([
      makeSection([
        makeParagraph([
          { kind: 'run', children: [{ kind: 'text', value: 'p0' }] },
        ]),
        makeParagraph([
          { kind: 'comment-range', id: '9', boundary: 'start' },
          { kind: 'run', children: [{ kind: 'text', value: 'p1' }] },
        ]),
        makeParagraph([
          {
            kind: 'run',
            children: [{ kind: 'comment-reference', id: '9' }],
          },
        ]),
      ]),
    ])

    expect(findCommentAnchors(document).get('9')).toBe(1)
  })

  it('descends into table cells and reports the nested paragraph index', () => {
    const document = makeDocument([
      makeSection([
        makeParagraph([
          { kind: 'run', children: [{ kind: 'text', value: 'pre-table' }] },
        ]),
        {
          kind: 'table',
          rows: [
            {
              kind: 'table-row',
              cells: [
                {
                  kind: 'table-cell',
                  blocks: [
                    makeParagraph([
                      { kind: 'comment-range', id: '10', boundary: 'start' },
                      { kind: 'run', children: [{ kind: 'text', value: 'cell' }] },
                    ]),
                  ],
                },
              ],
            },
          ],
        },
      ]),
    ])

    expect(findCommentAnchors(document).get('10')).toBe(1)
  })
})
