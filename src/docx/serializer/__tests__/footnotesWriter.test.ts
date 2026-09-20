import { describe, expect, it } from 'vitest'

import type { Endnote, Footnote, Paragraph } from '../../model'
import { parseEndnotes, parseFootnotes } from '../../parser'
import { writeEndnotesXml, writeFootnotesXml } from '../footnotesWriter'

function makeParagraphWithUnknownRunChild(rawXml: string, text: string): Paragraph {
  return {
    kind: 'paragraph',
    children: [
      {
        kind: 'run',
        children: [{ kind: 'unknown', xml: rawXml }, { kind: 'text', value: text }],
      },
    ],
  }
}

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

  it(
    'substitutes real XML back in for a nested unrecognized run child instead of leaking an '
      + 'unrestored atlas-raw-unknown placeholder, across multiple footnotes sharing one part',
    () => {
      const footnoteA: Footnote = {
        kind: 'footnote',
        id: '1',
        blocks: [makeParagraphWithUnknownRunChild('<w:proofErr w:type="spellStart"/>', 'Fisrt')],
      }
      const footnoteB: Footnote = {
        kind: 'footnote',
        id: '2',
        blocks: [makeParagraphWithUnknownRunChild('<w:proofErr w:type="spellEnd"/>', 'Second')],
      }

      const xml = writeFootnotesXml([footnoteA, footnoteB])

      expect(xml).not.toContain('atlas-raw-unknown')
      expect(xml).toContain('<w:proofErr w:type="spellStart"/>')
      expect(xml).toContain('<w:proofErr w:type="spellEnd"/>')
    },
  )

  // DOCX-2 — round-trip fidelity audit follow-up: parseFootnotes/parseEndnotes
  // (via partBody.ts's parseBlocksFromXmlFragment) now record any w:sdt/
  // mc:AlternateContent wrapper region for a note's body, and
  // writeFootnotesXml/writeEndnotesXml now retrieve it (through
  // getWrapperRegionsForFragmentBlocks) instead of discarding it — see
  // partBody.ts's WeakMap doc comment for why this needs no new field on
  // the Footnote/Endnote model types themselves. These go through the real
  // parser, not a hand-built model object, because the passthrough is keyed
  // on the exact `blocks` array reference the parser produced.
  it('round-trips an unedited w:sdt content control inside a footnote byte-identical (DOCX-2)', () => {
    const footnotesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
      + '<w:footnote w:id="1"><w:sdt><w:sdtPr><w:id w:val="42"/><w:alias w:val="My Control"/></w:sdtPr>'
      + '<w:sdtContent><w:p><w:r><w:t xml:space="preserve">Hi</w:t></w:r></w:p></w:sdtContent></w:sdt>'
      + '</w:footnote></w:footnotes>'

    const footnotes = Array.from(parseFootnotes(footnotesXml).values())
    const xml = writeFootnotesXml(footnotes)

    expect(xml).toContain(
      '<w:sdt><w:sdtPr><w:id w:val="42"/><w:alias w:val="My Control"/></w:sdtPr>'
        + '<w:sdtContent><w:p><w:r><w:t xml:space="preserve">Hi</w:t></w:r></w:p></w:sdtContent></w:sdt>',
    )
  })

  it('round-trips an unedited mc:AlternateContent shape fallback inside an endnote byte-identical (DOCX-2)', () => {
    const endnotesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      + 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
      + '<w:endnote w:id="1"><mc:AlternateContent><mc:Choice Requires="wps">'
      + '<w:p><w:r><w:t xml:space="preserve">Shape text</w:t></w:r></w:p></mc:Choice>'
      + '<mc:Fallback><w:p><w:r><w:t xml:space="preserve">Fallback text</w:t></w:r></w:p></mc:Fallback>'
      + '</mc:AlternateContent></w:endnote></w:endnotes>'

    const endnotes = Array.from(parseEndnotes(endnotesXml).values())
    const xml = writeEndnotesXml(endnotes)

    expect(xml).toContain(
      '<mc:AlternateContent><mc:Choice Requires="wps">'
        + '<w:p><w:r><w:t xml:space="preserve">Shape text</w:t></w:r></w:p></mc:Choice>'
        + '<mc:Fallback><w:p><w:r><w:t xml:space="preserve">Fallback text</w:t></w:r></w:p></mc:Fallback>'
        + '</mc:AlternateContent>',
    )
  })
})
