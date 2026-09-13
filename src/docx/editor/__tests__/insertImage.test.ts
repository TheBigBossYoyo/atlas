import { describe, expect, it } from 'vitest'

import type { DocxBundle } from '../../index'
import type { Document, Paragraph, Section } from '../../model/document'
import type { Position } from '../commandTypes'
import { insertImageIntoBundle, type InsertImageInput } from '../insertImage'

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

function makeBundle(): DocxBundle {
  const paragraph: Paragraph = {
    kind: 'paragraph',
    children: [{ kind: 'run', children: [{ kind: 'text', value: 'Hello' }] }],
  }
  const section: Section = { kind: 'section', props: {}, blocks: [paragraph] }
  const document: Document = {
    kind: 'document',
    sections: [section],
    ...emptyMaps(),
  }
  return {
    document,
    relationships: [
      {
        id: 'rId1',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
        target: 'styles.xml',
      },
    ],
    rawArchive: new Map<string, Uint8Array>(),
  }
}

const PNG_BYTES = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])

const IMAGE: InsertImageInput = {
  bytes: PNG_BYTES,
  mime: 'image/png',
  suggestedName: 'shot.png',
  widthPt: 200,
  heightPt: 100,
  description: 'a screenshot',
}

const POSITION: Position = { paragraphPath: [0, 0], runIndex: 0, charOffset: 0 }

describe('insertImageIntoBundle', () => {
  it('appends a new run with a Drawing child to the target paragraph', () => {
    const bundle = makeBundle()
    const result = insertImageIntoBundle(bundle, POSITION, IMAGE)

    const section = result.document.sections[0]
    expect(section).toBeDefined()
    const paragraph = section!.blocks[0] as Paragraph
    expect(paragraph.children.length).toBe(2)

    const newRun = paragraph.children[1]
    expect(newRun?.kind).toBe('run')
    if (newRun?.kind !== 'run') return
    expect(newRun.children.length).toBe(1)
    const drawing = newRun.children[0]
    expect(drawing?.kind).toBe('drawing')
    if (drawing?.kind !== 'drawing') return
    expect(drawing.layout).toBe('inline')
    expect(drawing.relationshipId).toBe('rId2')
    expect(drawing.extent).toEqual({ cx: 200 * 12700, cy: 100 * 12700 })
    expect(drawing.description).toBe('a screenshot')
  })

  it('allocates a fresh relationship id and registers the image type', () => {
    const bundle = makeBundle()
    const result = insertImageIntoBundle(bundle, POSITION, IMAGE)

    expect(result.bundle.relationships).toBeDefined()
    const rels = result.bundle.relationships!
    expect(rels.length).toBe(2)
    const added = rels[1]
    expect(added?.id).toBe('rId2')
    expect(added?.target).toBe('media/image1.png')
    expect(added?.type).toContain('/image')
  })

  it('writes the image bytes into rawArchive at word/media/image1.png', () => {
    const bundle = makeBundle()
    const result = insertImageIntoBundle(bundle, POSITION, IMAGE)

    expect(result.bundle.rawArchive).toBeDefined()
    const archive = result.bundle.rawArchive!
    const stored = archive.get('word/media/image1.png')
    expect(stored).toBeDefined()
    expect(stored!.length).toBe(PNG_BYTES.length)
  })

  it('avoids colliding media filenames', () => {
    const bundle = makeBundle()
    const archive = new Map<string, Uint8Array>(bundle.rawArchive)
    archive.set('word/media/image1.png', new Uint8Array([1]))
    archive.set('word/media/image2.png', new Uint8Array([2]))
    const seeded: DocxBundle = { ...bundle, rawArchive: archive }
    const result = insertImageIntoBundle(seeded, POSITION, IMAGE)
    expect(result.bundle.rawArchive!.has('word/media/image3.png')).toBe(true)
  })

  it('positions the cursor immediately after the inserted run', () => {
    const bundle = makeBundle()
    const result = insertImageIntoBundle(bundle, POSITION, IMAGE)
    expect(result.range.anchor).toEqual({
      paragraphPath: [0, 0],
      runIndex: 2,
      charOffset: 0,
    })
    expect(result.range.focus).toEqual(result.range.anchor)
  })

  it('does not mutate the input bundle, document, or relationships', () => {
    const bundle = makeBundle()
    const originalRelsLen = bundle.relationships!.length
    const originalArchiveSize = bundle.rawArchive!.size
    const originalParagraphChildren = (bundle.document.sections[0]!.blocks[0] as Paragraph).children
      .length

    insertImageIntoBundle(bundle, POSITION, IMAGE)

    expect(bundle.relationships!.length).toBe(originalRelsLen)
    expect(bundle.rawArchive!.size).toBe(originalArchiveSize)
    expect(
      (bundle.document.sections[0]!.blocks[0] as Paragraph).children.length,
    ).toBe(originalParagraphChildren)
  })

  it('throws when the paragraph path points to an out-of-range section', () => {
    const bundle = makeBundle()
    const badPos: Position = { paragraphPath: [5, 0], runIndex: 0, charOffset: 0 }
    expect(() => insertImageIntoBundle(bundle, badPos, IMAGE)).toThrow()
  })

  it('handles JPEG mime', () => {
    const bundle = makeBundle()
    const jpeg: InsertImageInput = { ...IMAGE, mime: 'image/jpeg', suggestedName: 'p.jpg' }
    const result = insertImageIntoBundle(bundle, POSITION, jpeg)
    expect(result.bundle.rawArchive!.has('word/media/image1.jpeg')).toBe(true)
  })
})
