/** USR-16 — PPTX edits rewrite only what they touch and reparse to the expected slides. */
import { beforeEach, describe, expect, it } from 'vitest'

import { parsePptxSlides } from '../../parser'
import { canEditNotes, deleteShape, insertTextBox, setShapeBox, setShapeText, setSlideNotes } from '../pptxEdits'
import { loadPptxPackage, packageArchive, readPart, writePptxPackage, type PptxPackage } from '../pptxPackage'
import { addSlide, deleteSlide, duplicateSlide, listSlides, moveSlide } from '../pptxSlideOps'
import { buildEditableDeck } from './editableDeck'

const SLIDE_1 = 'ppt/slides/slide1.xml'
const SLIDE_2 = 'ppt/slides/slide2.xml'

async function slides(pkg: PptxPackage) {
  return parsePptxSlides(packageArchive(pkg), { cancelled: false })
}

let pkg: PptxPackage
beforeEach(async () => {
  pkg = await loadPptxPackage(await buildEditableDeck())
})

describe('shape edits', () => {
  it('replaces text while keeping paragraph and run formatting', async () => {
    const next = setShapeText(pkg, SLIDE_1, '2', 'Annual review\nSecond line')
    const xml = readPart(next, SLIDE_1)!
    expect(xml).toContain('<a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="4000" b="1"/><a:t>Annual review</a:t>')
    expect(xml).toContain('<a:t>Second line</a:t>')
    const [first] = await slides(next)
    const title = first.shapes.find((s) => s.sourceId === '2')
    expect(title).toMatchObject({ kind: 'text', text: 'Annual review\nSecond line' })
    // Untouched parts are the very same strings.
    expect(next.parts.get('ppt/slideLayouts/slideLayout1.xml')).toBe(pkg.parts.get('ppt/slideLayouts/slideLayout1.xml'))
  })

  it('returns the same package for an unknown shape', () => {
    expect(setShapeText(pkg, SLIDE_1, '999', 'x')).toBe(pkg)
  })

  it('writes a box for an inherited placeholder and keeps rotation on an existing xfrm', async () => {
    let next = setShapeBox(pkg, SLIDE_1, '2', { x: 10, y: 20, w: 300, h: 40 })
    next = setShapeBox(next, SLIDE_1, '5', { x: 50, y: 60, w: 200, h: 30 })
    const [first] = await slides(next)
    expect(first.shapes.find((s) => s.sourceId === '2')?.transform).toMatchObject({ x: 10, y: 20, w: 300, h: 40 })
    expect(first.shapes.find((s) => s.sourceId === '5')?.transform).toMatchObject({ x: 50, y: 60, w: 200, h: 30, rotationDeg: 10 })
  })

  it('inserts and deletes a text box', async () => {
    const inserted = insertTextBox(pkg, SLIDE_2, { x: 100, y: 100, w: 400, h: 50 }, 'New <text> & more')
    expect(inserted.sourceId).toBe('3')
    let [, second] = await slides(inserted.pkg)
    expect(second.shapes.find((s) => s.sourceId === '3')).toMatchObject({ kind: 'text', text: 'New <text> & more', movable: true })

    const removed = deleteShape(inserted.pkg, SLIDE_2, '3')
    ;[, second] = await slides(removed)
    expect(second.shapes.some((s) => s.sourceId === '3')).toBe(false)
  })

  it('edits existing notes and creates a notes page from the notes master', async () => {
    expect(canEditNotes(pkg, SLIDE_2)).toBe(true)
    let next = setSlideNotes(pkg, SLIDE_1, 'Thank the team')
    next = setSlideNotes(next, SLIDE_2, 'Go through the agenda\nKeep it short')
    const [first, second] = await slides(next)
    expect(first.notes).toBe('Thank the team')
    expect(second.notes).toBe('Go through the agenda\nKeep it short')
    expect(readPart(next, '[Content_Types].xml')).toContain('/ppt/notesSlides/notesSlide2.xml')
  })
})

describe('slide operations', () => {
  it('adds a "Title and Content" slide after a title slide, with empty editable placeholders', async () => {
    const { pkg: next, index } = addSlide(pkg, 0)
    expect(index).toBe(1)
    const deck = await slides(next)
    expect(deck.map((s) => s.partPath)).toEqual([SLIDE_1, 'ppt/slides/slide3.xml', SLIDE_2])
    expect(deck[1].shapes.map((s) => (s.kind === 'text' ? s.placeholderPrompt : s.kind))).toEqual([
      'Click to add title',
      'Click to add text',
    ])
    expect(readPart(next, 'ppt/slides/_rels/slide3.xml.rels')).toContain('../slideLayouts/slideLayout2.xml')
    const presentation = readPart(next, 'ppt/presentation.xml')!
    expect(presentation.match(/id="258"/g)).toHaveLength(2) // slide list + section list
  })

  it('duplicates a slide without its notes page', async () => {
    const next = duplicateSlide(pkg, 0)
    const deck = await slides(next)
    expect(deck).toHaveLength(3)
    expect(deck[1].title).toBe('Quarterly review')
    expect(deck[1].notes).toBeUndefined()
    expect(deck[0].notes).toBe('Welcome everyone')
  })

  it('deletes a slide with its notes and refuses to delete the last one', async () => {
    const next = deleteSlide(pkg, 0)
    expect(listSlides(next).map((s) => s.path)).toEqual([SLIDE_2])
    expect(next.parts.has('ppt/notesSlides/notesSlide1.xml')).toBe(false)
    expect(readPart(next, '[Content_Types].xml')).not.toContain('slide1.xml')
    expect(readPart(next, 'ppt/presentation.xml')).not.toContain('id="256"')
    expect(deleteSlide(next, 0)).toBe(next)
  })

  it('moves slides and round-trips through a saved file', async () => {
    const moved = moveSlide(pkg, 1, 0)
    const reopened = await loadPptxPackage((await writePptxPackage(moved)).buffer as ArrayBuffer)
    const deck = await slides(reopened)
    expect(deck.map((s) => s.title)).toEqual(['Agenda', 'Quarterly review'])
    expect(reopened.parts.get('ppt/media/image1.png')).toEqual(pkg.parts.get('ppt/media/image1.png'))
  })
})
