import { describe, expect, it } from 'vitest'

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
})
