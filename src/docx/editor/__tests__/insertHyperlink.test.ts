import { describe, expect, it } from 'vitest'

import type { DocxBundle } from '../../index'
import type { Document, Paragraph, Section } from '../../model/document'
import { applyCommand } from '../commands'
import type { Range } from '../commandTypes'
import { insertHyperlinkIntoBundle } from '../insertHyperlink'

function emptyMaps() {
  return {
    styles: new Map(),
    numbering: new Map(),
    comments: new Map(),
    footnotes: new Map(),
    endnotes: new Map(),
    headers: new Map(),
    footers: new Map(),
  }
}

function makeBundle(text = 'Visit our site today'): DocxBundle {
  const paragraph: Paragraph = {
    kind: 'paragraph',
    children: [{ kind: 'run', children: [{ kind: 'text', value: text }] }],
  }
  const section: Section = { kind: 'section', props: {}, blocks: [paragraph] }
  const document: Document = { kind: 'document', sections: [section], ...emptyMaps() }
  return {
    document,
    relationships: [
      { id: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', target: 'styles.xml' },
    ],
    rawArchive: new Map<string, Uint8Array>(),
  }
}

describe('insertHyperlinkIntoBundle', () => {
  it('allocates a fresh External relationship id', () => {
    const bundle = makeBundle()
    const range: Range = { anchor: { paragraphPath: [0, 0], runIndex: 0, charOffset: 6 }, focus: { paragraphPath: [0, 0], runIndex: 0, charOffset: 9 } }

    const result = insertHyperlinkIntoBundle(bundle, range, 'https://example.com')

    expect(result.bundle.relationships).toHaveLength(2)
    const added = result.bundle.relationships![1]
    expect(added.id).toBe('rId2')
    expect(added.type).toContain('/hyperlink')
    expect(added.target).toBe('https://example.com')
    expect(added.targetMode).toBe('External')
  })

  it('wraps the selected text in a Hyperlink referencing the new relationship', () => {
    const bundle = makeBundle()
    // "Visit our site today" — offset 6 = 'o' of "our", offset 9 = right
    // after "our".
    const range: Range = { anchor: { paragraphPath: [0, 0], runIndex: 0, charOffset: 6 }, focus: { paragraphPath: [0, 0], runIndex: 0, charOffset: 9 } }

    const result = insertHyperlinkIntoBundle(bundle, range, 'https://example.com')

    const paragraph = result.document.sections[0].blocks[0] as Paragraph
    expect(paragraph.children[1]).toMatchObject({ kind: 'hyperlink', relationshipId: 'rId2' })
  })

  it('does not mutate the input bundle', () => {
    const bundle = makeBundle()
    const originalRelsLen = bundle.relationships!.length
    const range: Range = { anchor: { paragraphPath: [0, 0], runIndex: 0, charOffset: 0 }, focus: { paragraphPath: [0, 0], runIndex: 0, charOffset: 0 } }

    insertHyperlinkIntoBundle(bundle, range, 'https://example.com')

    expect(bundle.relationships!.length).toBe(originalRelsLen)
  })

  it('returns an inverse that removes the hyperlink and restores the original paragraph exactly', () => {
    const bundle = makeBundle()
    const range: Range = { anchor: { paragraphPath: [0, 0], runIndex: 0, charOffset: 6 }, focus: { paragraphPath: [0, 0], runIndex: 0, charOffset: 9 } }

    const result = insertHyperlinkIntoBundle(bundle, range, 'https://example.com')
    const reverted = applyCommand(result.document, result.inverse)

    expect(reverted.document).toEqual(bundle.document)
  })

  it('avoids colliding relationship ids when the document already defines rId2', () => {
    const bundle: DocxBundle = {
      ...makeBundle(),
      relationships: [
        { id: 'rId1', type: 'x', target: 'a' },
        { id: 'rId2', type: 'y', target: 'b' },
      ],
    }
    const range: Range = { anchor: { paragraphPath: [0, 0], runIndex: 0, charOffset: 0 }, focus: { paragraphPath: [0, 0], runIndex: 0, charOffset: 0 } }

    const result = insertHyperlinkIntoBundle(bundle, range, 'https://example.com')

    expect(result.bundle.relationships![2].id).toBe('rId3')
  })
})
