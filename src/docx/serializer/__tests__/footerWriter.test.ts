import { describe, expect, it } from 'vitest'

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
})
