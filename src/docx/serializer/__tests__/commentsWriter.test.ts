import { describe, expect, it } from 'vitest'

import type { Comment, Paragraph } from '../../model'
import { parseComments } from '../../parser'
import { writeCommentsXml } from '../commentsWriter'

function makeParagraph(text: string): Paragraph {
  return {
    kind: 'paragraph',
    children: [
      {
        kind: 'run',
        children: [{ kind: 'text', value: text }],
      },
    ],
  }
}

function makeComment(
  id: string,
  texts: ReadonlyArray<string>,
  metadata: Pick<Comment, 'author' | 'date' | 'initials' | 'parentId'> = {},
): Comment {
  return {
    kind: 'comment',
    id,
    ...metadata,
    body: texts.map(makeParagraph),
  }
}

describe('writeCommentsXml', () => {
  it('writes an empty comments part', () => {
    const xml = writeCommentsXml([])

    expect(xml.startsWith('<?xml')).toBe(true)
    expect(xml).toContain('<w:comments')
    expect(xml).not.toContain('<w:comment ')
  })

  it('writes a single comment with metadata', () => {
    const xml = writeCommentsXml([
      makeComment('1', ['This needs revision.'], {
        author: 'Alice Smith',
        date: '2024-01-15T10:30:00Z',
        initials: 'AS',
      }),
    ])

    expect(xml).toContain('<w:comment w:id="1" w:author="Alice Smith" w:date="2024-01-15T10:30:00Z" w:initials="AS">')
    expect(xml).toContain('This needs revision.')
  })

  it('writes multiple comments', () => {
    const xml = writeCommentsXml([
      makeComment('1', ['Alpha']),
      makeComment('2', ['Beta']),
    ])

    expect(xml).toContain('<w:comment w:id="1">')
    expect(xml).toContain('<w:comment w:id="2">')
    expect(xml).toContain('Alpha')
    expect(xml).toContain('Beta')
  })

  it('writes comments with multiple paragraphs', () => {
    const xml = writeCommentsXml([
      makeComment('3', ['Line one.', 'Line two.']),
    ])

    expect(xml).toContain('<w:comment w:id="3">')
    expect(xml).toContain('Line one.')
    expect(xml).toContain('Line two.')
  })

  it('writes threaded replies with w15:parentId', () => {
    const xml = writeCommentsXml([makeComment('4', ['Reply'], { parentId: '1' })])

    expect(xml).toContain('xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"')
    expect(xml).toContain('<w:comment w:id="4" w15:parentId="1">')
  })

  it('declares the full standard namespace set even without a threaded reply (D19 / DXS-08)', () => {
    const xml = writeCommentsXml([makeComment('1', ['Plain comment, no reply'])])

    // Previously only xmlns:w (+ a conditional xmlns:w15) was declared, so
    // a table/drawing/hyperlink inside a comment body emitted an
    // undeclared namespace prefix.
    for (const prefix of ['w', 'r', 'wp', 'a', 'pic', 'v', 'mc', 'w14', 'w15']) {
      expect(xml).toContain(`xmlns:${prefix}=`)
    }
  })

  it(
    'substitutes real XML back in for a nested unrecognized node instead of leaking an '
      + 'unrestored atlas-raw-unknown placeholder',
    () => {
      const comment: Comment = {
        kind: 'comment',
        id: '1',
        body: [
          {
            kind: 'paragraph',
            children: [
              {
                kind: 'run',
                children: [
                  { kind: 'unknown', xml: '<w:proofErr w:type="spellStart"/>' },
                  { kind: 'text', value: 'Helo' },
                ],
              },
            ],
          },
        ],
      }

      const xml = writeCommentsXml([comment])

      expect(xml).not.toContain('atlas-raw-unknown')
      expect(xml).toContain('<w:proofErr w:type="spellStart"/>')
    },
  )

  // DOCX-2 — round-trip fidelity audit follow-up: parseComments (via
  // partBody.ts's parseBlocksFromXmlFragment) now records any w:sdt/
  // mc:AlternateContent wrapper region for a comment's body, and
  // writeCommentsXml now retrieves it (through
  // getWrapperRegionsForFragmentBlocks) instead of discarding it — see
  // partBody.ts's WeakMap doc comment for why this needs no new field on
  // the Comment model type itself. Goes through the real parser, not a
  // hand-built model object, because the passthrough is keyed on the exact
  // `body` array reference the parser produced.
  it('round-trips an unedited w:sdt content control inside a comment byte-identical (DOCX-2)', () => {
    const commentsXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
      + '<w:comment w:id="1" w:author="Reviewer"><w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr>'
      + '<w:sdtContent><w:p><w:r><w:t xml:space="preserve">Approved</w:t></w:r></w:p></w:sdtContent></w:sdt>'
      + '</w:comment></w:comments>'

    const comments = Array.from(parseComments(commentsXml).values())
    const xml = writeCommentsXml(comments)

    expect(xml).toContain(
      '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr>'
        + '<w:sdtContent><w:p><w:r><w:t xml:space="preserve">Approved</w:t></w:r></w:p></w:sdtContent></w:sdt>',
    )
  })
})
