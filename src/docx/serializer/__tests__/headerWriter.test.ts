import { describe, expect, it } from 'vitest'

import { parseHeader } from '../../parser/headers'
import type { Header, Paragraph } from '../../model'
import { writeHeaderXml } from '../headerWriter'

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

function makeHeader(blocks: ReadonlyArray<Paragraph>): Header {
  return {
    kind: 'header',
    id: 'rId3',
    blocks,
  }
}

describe('writeHeaderXml', () => {
  it('writes an empty header part', () => {
    const xml = writeHeaderXml(makeHeader([]))

    expect(xml.startsWith('<?xml')).toBe(true)
    expect(xml).toContain('<w:hdr')
    expect(xml).not.toContain('<w:p')
  })

  it('writes a header paragraph with text', () => {
    const xml = writeHeaderXml(makeHeader([makeParagraph('Page Header')]))

    expect(xml).toContain('<w:hdr')
    expect(xml).toContain('<w:p>')
    expect(xml).toContain('Page Header')
  })

  it('writes multiple header paragraphs', () => {
    const xml = writeHeaderXml(makeHeader([makeParagraph('First'), makeParagraph('Second')]))

    expect(xml).toContain('First')
    expect(xml).toContain('Second')
  })

  it('declares the full standard namespace set, not just xmlns:w (D19 / DXS-08)', () => {
    const xml = writeHeaderXml(makeHeader([]))

    // Previously only xmlns:w (+ a hardcoded xmlns:r) was declared, so a
    // drawing/hyperlink/shape inside a header emitted an undeclared
    // namespace prefix — an XML well-formedness violation.
    for (const prefix of ['w', 'r', 'wp', 'a', 'pic', 'v', 'mc', 'w14']) {
      expect(xml).toContain(`xmlns:${prefix}=`)
    }
  })

  it(
    'substitutes real XML back in for a nested unrecognized node (e.g. w:proofErr) instead of '
      + 'leaking an unrestored atlas-raw-unknown placeholder',
    () => {
      // w:proofErr is not modeled (it becomes an UnknownNode run child) but
      // is ubiquitous in real Word-authored documents (inserted around
      // nearly every word the spell-checker flags) — a header/footer/
      // footnote/endnote/comment containing one previously round-tripped
      // to a literal, un-substituted `<atlas-raw-unknown data-id="..."/>`
      // element instead of the original raw XML, because buildParagraph's
      // placeholder/restoration mechanism only ran for the main document
      // body, never for these standalone parts.
      const sourceXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        + '<w:p><w:r><w:proofErr w:type="spellStart"/><w:t>Helo</w:t><w:proofErr w:type="spellEnd"/></w:r></w:p>'
        + '</w:hdr>'

      const header = parseHeader(sourceXml, 'rId1')
      const written = writeHeaderXml(header)

      expect(written).not.toContain('atlas-raw-unknown')
      expect(written).toContain('<w:proofErr w:type="spellStart"/>')
      expect(written).toContain('<w:proofErr w:type="spellEnd"/>')
    },
  )
})
