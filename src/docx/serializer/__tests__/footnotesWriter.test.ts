import { describe, expect, it } from 'vitest'

import type { Endnote, Footnote, Paragraph } from '../../model'
import { writeEndnotesXml, writeFootnotesXml } from '../footnotesWriter'

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

function makeFootnote(
  id: string,
  text: string,
  noteType?: Footnote['noteType'],
): Footnote {
  return {
    kind: 'footnote',
    id,
    ...(noteType !== undefined ? { noteType } : {}),
    blocks: text.length === 0 ? [] : [makeParagraph(text)],
  }
}

function makeEndnote(id: string, text: string, noteType?: Endnote['noteType']): Endnote {
  return {
    kind: 'endnote',
    id,
    ...(noteType !== undefined ? { noteType } : {}),
    blocks: text.length === 0 ? [] : [makeParagraph(text)],
  }
}

describe('writeFootnotesXml', () => {
  it('writes an empty footnotes part', () => {
    const xml = writeFootnotesXml([])

    expect(xml.startsWith('<?xml')).toBe(true)
    expect(xml).toContain('<w:footnotes')
    expect(xml).not.toContain('<w:footnote ')
  })

  it('writes separator, continuation separator, and a user footnote', () => {
    const xml = writeFootnotesXml([
      makeFootnote('-1', 'Separator', 'separator'),
      makeFootnote('0', 'Continuation', 'continuationSeparator'),
      makeFootnote('1', 'First footnote text.'),
    ])

    expect(xml).toContain('<w:footnote w:id="-1" w:type="separator">')
    expect(xml).toContain('<w:footnote w:id="0" w:type="continuationSeparator">')
    expect(xml).toContain('<w:footnote w:id="1">')
    expect(xml).toContain('First footnote text.')
  })

  it('writes multiple user footnotes', () => {
    const xml = writeFootnotesXml([
      makeFootnote('1', 'First'),
      makeFootnote('2', 'Second'),
    ])

    expect(xml).toContain('<w:footnote w:id="1">')
    expect(xml).toContain('<w:footnote w:id="2">')
    expect(xml).toContain('First')
    expect(xml).toContain('Second')
  })

  it('writes endnotes with the same overall shape', () => {
    const xml = writeEndnotesXml([
      makeEndnote('-1', 'Separator', 'separator'),
      makeEndnote('3', 'Endnote text'),
    ])

    expect(xml.startsWith('<?xml')).toBe(true)
    expect(xml).toContain('<w:endnotes')
    expect(xml).toContain('<w:endnote w:id="-1" w:type="separator">')
    expect(xml).toContain('<w:endnote w:id="3">')
    expect(xml).toContain('Endnote text')
  })
})
