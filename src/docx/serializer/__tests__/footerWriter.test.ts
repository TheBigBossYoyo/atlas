import { describe, expect, it } from 'vitest'

import { parseFooter } from '../../parser/footers'
import type { Footer, Paragraph } from '../../model'
import { writeFooterXml } from '../footerWriter'

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

function makeFooter(blocks: ReadonlyArray<Paragraph>): Footer {
  return {
    kind: 'footer',
    id: 'rId4',
    blocks,
  }
}

describe('writeFooterXml', () => {
  it('writes an empty footer part', () => {
    const xml = writeFooterXml(makeFooter([]))

    expect(xml.startsWith('<?xml')).toBe(true)
    expect(xml).toContain('<w:ftr')
    expect(xml).not.toContain('<w:p')
  })

  it('writes a footer paragraph with text', () => {
    const xml = writeFooterXml(makeFooter([makeParagraph('Page Footer')]))

    expect(xml).toContain('<w:ftr')
    expect(xml).toContain('<w:p>')
    expect(xml).toContain('Page Footer')
  })

  it('writes multiple footer paragraphs', () => {
    const xml = writeFooterXml(makeFooter([makeParagraph('Left'), makeParagraph('Right')]))

    expect(xml).toContain('Left')
    expect(xml).toContain('Right')
  })

  it(
    'substitutes real XML back in for a nested unrecognized node (e.g. w:proofErr) instead of '
      + 'leaking an unrestored atlas-raw-unknown placeholder',
    () => {
      const sourceXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        + '<w:p><w:r><w:proofErr w:type="spellStart"/><w:t>Helo</w:t><w:proofErr w:type="spellEnd"/></w:r></w:p>'
        + '</w:ftr>'

      const footer = parseFooter(sourceXml, 'rId2')
      const written = writeFooterXml(footer)

      expect(written).not.toContain('atlas-raw-unknown')
      expect(written).toContain('<w:proofErr w:type="spellStart"/>')
      expect(written).toContain('<w:proofErr w:type="spellEnd"/>')
    },
  )
})
