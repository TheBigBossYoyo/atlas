/** USR-16 (ODP) — edits rewrite only `content.xml`, and the deck re-parses to what was asked for. */
import JSZip from 'jszip'
import { beforeEach, describe, expect, it } from 'vitest'

import { loadOfficePackage, packageArchive, readPart, writeOfficePackage, type OfficePackage } from '../../../../../office/officePackage'
import { parseOdpSlides } from '../../parser'
import { buildOdpFixtureZip } from '../../__tests__/odpFixture'
import {
  CONTENT_PART,
  addSlide,
  deleteShape,
  deleteSlide,
  duplicateSlide,
  insertTextBox,
  moveSlide,
  setShapeBox,
  setShapeText,
  setSlideNotes,
} from '../odpEdits'

async function slidesOf(pkg: OfficePackage) {
  return parseOdpSlides(packageArchive(pkg), { cancelled: false })
}

let pkg: OfficePackage
beforeEach(async () => {
  pkg = await loadOfficePackage(await buildOdpFixtureZip().generateAsync({ type: 'arraybuffer' }))
})

describe('shape edits', () => {
  it('replaces a shape text, leaving every other part of the package alone', async () => {
    const [before] = await slidesOf(pkg)
    const target = before.shapes.find((shape) => shape.kind === 'text' && shape.text.length > 0)!
    const next = setShapeText(pkg, 0, target.sourceId!, 'Rewritten title\nSecond line')

    const [after] = await slidesOf(next)
    const edited = after.shapes.find((shape) => shape.sourceId === target.sourceId)
    expect(edited).toMatchObject({ kind: 'text', text: 'Rewritten title\nSecond line' })
    expect(next.parts.get('styles.xml')).toBe(pkg.parts.get('styles.xml'))
    expect(readPart(next, CONTENT_PART)).not.toBe(readPart(pkg, CONTENT_PART))
  })

  it('moves and resizes a shape in centimetres, which read back as the same pixels', async () => {
    const [before] = await slidesOf(pkg)
    const target = before.shapes.find((shape) => shape.movable)!
    const next = setShapeBox(pkg, 0, target.sourceId!, { x: 100, y: 50, w: 300, h: 120 })

    const [after] = await slidesOf(next)
    const moved = after.shapes.find((shape) => shape.sourceId === target.sourceId)!
    expect(moved.transform.x).toBeCloseTo(100, 0)
    expect(moved.transform.y).toBeCloseTo(50, 0)
    expect(moved.transform.w).toBeCloseTo(300, 0)
    expect(moved.transform.h).toBeCloseTo(120, 0)
  })

  it('inserts a text box and deletes a shape', async () => {
    const inserted = insertTextBox(pkg, 0, { x: 10, y: 10, w: 200, h: 50 }, 'Added here')
    expect(inserted.sourceId).not.toBeNull()

    const [withBox] = await slidesOf(inserted.pkg)
    const added = withBox.shapes.find((shape) => shape.sourceId === inserted.sourceId)
    expect(added).toMatchObject({ kind: 'text', text: 'Added here', movable: true })

    const [afterDelete] = await slidesOf(deleteShape(inserted.pkg, 0, inserted.sourceId!))
    expect(afterDelete.shapes.some((shape) => shape.kind === 'text' && shape.text === 'Added here')).toBe(false)
  })

  it('leaves the package untouched when the shape or slide does not exist', () => {
    expect(setShapeText(pkg, 9, '0', 'nope')).toBe(pkg)
    expect(setShapeText(pkg, 0, '999', 'nope')).toBe(pkg)
    expect(deleteShape(pkg, 0, 'not-a-number')).toBe(pkg)
  })
})

describe('notes and slides', () => {
  it('writes speaker notes, creating the notes page when there is none', async () => {
    const next = setSlideNotes(pkg, 0, 'Remember to smile\nAnd breathe')
    const [slide] = await slidesOf(next)
    expect(slide.notes).toContain('Remember to smile')
    expect(slide.notes).toContain('And breathe')
  })

  it('adds, duplicates, deletes and reorders slides', async () => {
    const before = await slidesOf(pkg)

    const added = addSlide(pkg, 0)
    expect(added.index).toBe(1)
    const afterAdd = await slidesOf(added.pkg)
    expect(afterAdd).toHaveLength(before.length + 1)
    expect(afterAdd[1].shapes).toHaveLength(0)

    const duplicated = await slidesOf(duplicateSlide(pkg, 0))
    expect(duplicated).toHaveLength(before.length + 1)
    expect(duplicated[1].title).toBe(before[0].title)

    if (before.length > 1) {
      const reordered = await slidesOf(moveSlide(pkg, 1, 0))
      expect(reordered[0].title).toBe(before[1].title)
    }

    let shrunk = pkg
    for (let index = before.length - 1; index > 0; index--) shrunk = deleteSlide(shrunk, index)
    expect(await slidesOf(shrunk)).toHaveLength(1)
    // The last slide cannot be deleted.
    expect(deleteSlide(shrunk, 0)).toBe(shrunk)
  })

  it('round-trips through a saved file, keeping mimetype first and uncompressed', async () => {
    const edited = setSlideNotes(setShapeText(pkg, 0, '0', 'Saved title'), 0, 'Saved notes')
    const bytes = await writeOfficePackage(edited)
    const zip = await JSZip.loadAsync(bytes)

    expect(Object.keys(zip.files)[0]).toBe('mimetype')
    expect(await zip.file('mimetype')!.async('string')).toContain('presentation')

    const reopened = await loadOfficePackage(bytes.buffer as ArrayBuffer)
    const [slide] = await slidesOf(reopened)
    expect(slide.notes).toContain('Saved notes')
  })
})
