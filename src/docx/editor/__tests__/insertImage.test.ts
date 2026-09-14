import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import type { DocxBundle } from '../../index'
import type { Document, Paragraph, Run, Section } from '../../model/document'
import { applyCommand } from '../commands'
import type { Position } from '../commandTypes'
import { decodeImageNaturalSizePt, insertImageIntoBundle, type InsertImageInput } from '../insertImage'

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

function makeBundle(text = 'Hello'): DocxBundle {
  const paragraph: Paragraph = {
    kind: 'paragraph',
    children: [{ kind: 'run', children: [{ kind: 'text', value: text }] }],
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

describe('insertImageIntoBundle', () => {
  it('inserts a run with a Drawing child at the caret position', () => {
    const bundle = makeBundle()
    const position: Position = { paragraphPath: [0, 0], runIndex: 0, charOffset: 0 }
    const result = insertImageIntoBundle(bundle, position, IMAGE)

    const section = result.document.sections[0]
    const paragraph = section!.blocks[0] as Paragraph
    expect(paragraph.children.length).toBe(2)

    const newRun = paragraph.children[0]
    expect(newRun?.kind).toBe('run')
    if (newRun?.kind !== 'run') return
    const drawing = newRun.children[0]
    expect(drawing?.kind).toBe('drawing')
    if (drawing?.kind !== 'drawing') return
    expect(drawing.layout).toBe('inline')
    expect(drawing.relationshipId).toBe('rId2')
    expect(drawing.extent).toEqual({ cx: 200 * 12700, cy: 100 * 12700 })
    expect(drawing.description).toBe('a screenshot')
  })

  it('DXE-18: inserts at the actual cursor, splitting the run so surrounding text is preserved', () => {
    const bundle = makeBundle('HelloWorld')
    const position: Position = { paragraphPath: [0, 0], runIndex: 0, charOffset: 5 }
    const result = insertImageIntoBundle(bundle, position, IMAGE)

    const paragraph = result.document.sections[0]!.blocks[0] as Paragraph
    expect(paragraph.children.length).toBe(3)
    const before = paragraph.children[0] as Run
    const after = paragraph.children[2] as Run
    expect(before.children).toEqual([{ kind: 'text', value: 'Hello' }])
    expect(after.children).toEqual([{ kind: 'text', value: 'World' }])
    expect(paragraph.children[1]).toMatchObject({ kind: 'run' })
  })

  it('allocates a fresh relationship id and registers the image type', () => {
    const bundle = makeBundle()
    const position: Position = { paragraphPath: [0, 0], runIndex: 0, charOffset: 0 }
    const result = insertImageIntoBundle(bundle, position, IMAGE)

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
    const position: Position = { paragraphPath: [0, 0], runIndex: 0, charOffset: 0 }
    const result = insertImageIntoBundle(bundle, position, IMAGE)

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
    const position: Position = { paragraphPath: [0, 0], runIndex: 0, charOffset: 0 }
    const result = insertImageIntoBundle(seeded, position, IMAGE)
    expect(result.bundle.rawArchive!.has('word/media/image3.png')).toBe(true)
  })

  it('positions the cursor immediately after the inserted image', () => {
    const bundle = makeBundle()
    const position: Position = { paragraphPath: [0, 0], runIndex: 0, charOffset: 0 }
    const result = insertImageIntoBundle(bundle, position, IMAGE)
    expect(result.range.anchor).toEqual({
      paragraphPath: [0, 0],
      runIndex: 1,
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

    const position: Position = { paragraphPath: [0, 0], runIndex: 0, charOffset: 0 }
    insertImageIntoBundle(bundle, position, IMAGE)

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
    const position: Position = { paragraphPath: [0, 0], runIndex: 0, charOffset: 0 }
    const result = insertImageIntoBundle(bundle, position, jpeg)
    expect(result.bundle.rawArchive!.has('word/media/image1.jpeg')).toBe(true)
  })

  it('registers a Default content-type entry for the image extension (DXS-07)', () => {
    const bundle = makeBundle()
    expect(bundle.contentTypes).toBeUndefined()

    const position: Position = { paragraphPath: [0, 0], runIndex: 0, charOffset: 0 }
    const result = insertImageIntoBundle(bundle, position, IMAGE)

    expect(result.bundle.contentTypes).toBeDefined()
    const defaults = result.bundle.contentTypes!.defaults
    expect(defaults).toContainEqual({ extension: 'png', contentType: 'image/png' })
  })

  it('does not duplicate the content-type entry when one already exists', () => {
    const bundle: DocxBundle = {
      ...makeBundle(),
      contentTypes: {
        defaults: [{ extension: 'png', contentType: 'image/png' }],
        overrides: [],
      },
    }

    const position: Position = { paragraphPath: [0, 0], runIndex: 0, charOffset: 0 }
    const result = insertImageIntoBundle(bundle, position, IMAGE)

    const defaults = result.bundle.contentTypes!.defaults
    expect(defaults.filter((entry) => entry.extension === 'png')).toHaveLength(1)
  })

  // ---------------------------------------------------------------------------
  // DXE-18 — routed through Command/History so it is undoable
  // ---------------------------------------------------------------------------

  it('returns an inverse command that removes the image and restores the original paragraph exactly', () => {
    const bundle = makeBundle('HelloWorld')
    const position: Position = { paragraphPath: [0, 0], runIndex: 0, charOffset: 5 }
    const result = insertImageIntoBundle(bundle, position, IMAGE)

    const reverted = applyCommand(result.document, result.inverse)
    expect(reverted.document).toEqual(bundle.document)
  })
})

describe('decodeImageNaturalSizePt', () => {
  const originalCreateImageBitmap = globalThis.createImageBitmap

  afterEach(() => {
    if (originalCreateImageBitmap === undefined) {
      Reflect.deleteProperty(globalThis, 'createImageBitmap')
    } else {
      globalThis.createImageBitmap = originalCreateImageBitmap
    }
    vi.restoreAllMocks()
  })

  beforeEach(() => {
    Reflect.deleteProperty(globalThis, 'createImageBitmap')
  })

  it('falls back to a fixed 200x150pt box when createImageBitmap is unavailable', async () => {
    const size = await decodeImageNaturalSizePt(PNG_BYTES, 'image/png')
    expect(size).toEqual({ widthPt: 200, heightPt: 150 })
  })

  it('converts decoded pixel dimensions to points at 96dpi', async () => {
    globalThis.createImageBitmap = vi.fn().mockResolvedValue({ width: 96, height: 48, close: vi.fn() })

    const size = await decodeImageNaturalSizePt(PNG_BYTES, 'image/png')

    expect(size.widthPt).toBeCloseTo(72)
    expect(size.heightPt).toBeCloseTo(36)
  })

  it('scales down a wider-than-page image, preserving its aspect ratio', async () => {
    // 1200x600px at 96dpi is 900x450pt — wider than the 468pt content-width
    // cap, and 2:1 aspect ratio (portrait/landscape must be preserved, not
    // forced into a 4:3 box).
    globalThis.createImageBitmap = vi.fn().mockResolvedValue({ width: 1200, height: 600, close: vi.fn() })

    const size = await decodeImageNaturalSizePt(PNG_BYTES, 'image/png')

    expect(size.widthPt).toBeCloseTo(468)
    expect(size.heightPt).toBeCloseTo(234)
    expect(size.widthPt / size.heightPt).toBeCloseTo(2)
  })

  it('preserves a portrait aspect ratio instead of forcing 4:3', async () => {
    globalThis.createImageBitmap = vi.fn().mockResolvedValue({ width: 400, height: 800, close: vi.fn() })

    const size = await decodeImageNaturalSizePt(PNG_BYTES, 'image/png')

    expect(size.widthPt / size.heightPt).toBeCloseTo(0.5)
  })

  it('falls back to the fixed box when decoding throws', async () => {
    globalThis.createImageBitmap = vi.fn().mockRejectedValue(new Error('unsupported image'))

    const size = await decodeImageNaturalSizePt(PNG_BYTES, 'image/png')

    expect(size).toEqual({ widthPt: 200, heightPt: 150 })
  })
})
