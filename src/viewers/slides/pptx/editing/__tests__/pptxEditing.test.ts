/** USR-16 — PPTX edits rewrite only what they touch and reparse to the expected slides. */
import JSZip from 'jszip'
import { beforeEach, describe, expect, it } from 'vitest'

import { parsePptxSlides } from '../../parser'
import { canEditNotes, deleteShape, insertTextBox, setShapeBox, setShapeText, setSlideNotes } from '../pptxEdits'
import {
  loadOfficePackage,
  packageArchive,
  readPart,
  withParts,
  writeOfficePackage,
  type OfficePackage,
} from '../../../../../office/officePackage'
import { addSlide, deleteSlide, duplicateSlide, listSlides, moveSlide } from '../pptxSlideOps'
import { buildEditableDeck } from './editableDeck'

const SLIDE_1 = 'ppt/slides/slide1.xml'
const SLIDE_2 = 'ppt/slides/slide2.xml'
const NS = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`

/**
 * SHELL-5 — a package with NO `ppt/slideLayouts/`, `ppt/slideMasters/` or
 * `ppt/theme/` parts at all (the exact shape `tests/e2e/fixtures/generate.mjs`
 * used to produce for every pptx fixture, and still does for
 * `sample-noid.pptx`). Real PowerPoint always writes these; a lossy
 * converter or hand-built file may not.
 */
async function buildLayoutlessDeck(): Promise<ArrayBuffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`)
  zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation ${NS}>
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>`)
  zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`)
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
  <p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>Only slide</a:t></a:r></a:p></p:txBody></p:sp>
</p:spTree></p:cSld></p:sld>`)
  return zip.generateAsync({ type: 'arraybuffer' })
}

async function slides(pkg: OfficePackage) {
  return parsePptxSlides(packageArchive(pkg), { cancelled: false })
}

let pkg: OfficePackage
beforeEach(async () => {
  pkg = await loadOfficePackage(await buildEditableDeck())
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

  // USR-16 regression — "Insert Text Box" seeds the new shape with empty
  // text (so typing immediately replaces it instead of landing after a
  // leftover filler word), which used to make the shape parse away to
  // nothing: it has no text AND no fill/border (`<a:noFill/>`, rect
  // geometry), so it fell through `buildShapeElement`'s
  // `text.length > 0 || placeholderPrompt !== undefined` check into the
  // `hasVisibleFill || border || ...` branch below it and came back `null`
  // — an invisible shape nothing could select, open for editing, or type
  // into. A manually inserted text box (`p:cNvSpPr txBox="1"`) must stay
  // addressable even with no text yet, same as an empty placeholder does.
  it('keeps a freshly inserted, still-empty text box addressable (not filtered out as invisible)', async () => {
    const inserted = insertTextBox(pkg, SLIDE_2, { x: 100, y: 100, w: 400, h: 50 }, '')
    expect(inserted.sourceId).toBe('3')
    const [, second] = await slides(inserted.pkg)
    const shape = second.shapes.find((s) => s.sourceId === '3')
    expect(shape).toBeDefined()
    expect(shape).toMatchObject({ kind: 'text', text: '', movable: true })
  })

  // SHELL-3 — a shape with no cNvPr id (real PowerPoint always writes one;
  // hand-built/converted files may not) used to be completely unaddressable:
  // parser.ts gave it `sourceId: undefined`, so SlideEditCanvas's hitTest
  // silently ignored every click on it, AND even a caller that somehow had a
  // sourceId for it couldn't have found it in the XML via id-matching. It
  // must now parse to a truthy, positional sourceId, and an edit using that
  // sourceId must land on the right (and only the right) shape's XML.
  it('SHELL-3 — a shape with no cNvPr id at all is addressable, and an edit lands on it alone', async () => {
    const noIdSlide = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>Has an id</a:t></a:r></a:p></p:txBody></p:sp>` +
      `<p:sp><p:txBody><a:bodyPr/><a:p><a:r><a:t>No cNvPr at all</a:t></a:r></a:p></p:txBody></p:sp>` +
      `</p:spTree></p:cSld></p:sld>`
    const withoutId = withParts(pkg, { [SLIDE_1]: noIdSlide })

    const [before] = await slides(withoutId)
    const idLess = before.shapes.find((s) => s.kind === 'text' && s.text === 'No cNvPr at all')
    expect(idLess?.sourceId).toBe('@1') // shape 0 has a real id ("2"); this one falls back to its position.

    const next = setShapeText(withoutId, SLIDE_1, idLess!.sourceId!, 'Edited via fallback address')
    const xml = readPart(next, SLIDE_1)!
    expect(xml).toContain('<a:t>Edited via fallback address</a:t>')
    expect(xml).toContain('<a:t>Has an id</a:t>') // the other shape is untouched.

    const [after] = await slides(next)
    expect(after.shapes.find((s) => s.kind === 'text' && s.sourceId === '2')).toMatchObject({ text: 'Has an id' })
    expect(after.shapes.find((s) => s.kind === 'text' && s.sourceId === '@1')).toMatchObject({
      text: 'Edited via fallback address',
    })
  })

  // SHELL-3 — same bug class: two shapes sharing one cNvPr id ("7" below) are
  // just as unaddressable by plain id-matching as no id at all — findShape's
  // old implementation always returned the FIRST cNvPr in the document that
  // matched, so an edit meant for the second shape landed on the first one
  // instead. Each must now resolve to its OWN distinct XML node.
  it('SHELL-3 — shapes sharing a duplicate cNvPr id are each addressed and edited independently', async () => {
    const dupSlide = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="7" name="First"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>First shape</a:t></a:r></a:p></p:txBody></p:sp>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="7" name="Second"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>Second shape</a:t></a:r></a:p></p:txBody></p:sp>` +
      `</p:spTree></p:cSld></p:sld>`
    const withDupIds = withParts(pkg, { [SLIDE_1]: dupSlide })

    const [before] = await slides(withDupIds)
    expect(before.shapes.map((s) => s.sourceId)).toEqual(['@0', '@1'])

    // Editing the SECOND shape must not touch the first.
    const next = setShapeText(withDupIds, SLIDE_1, '@1', 'Edited second only')
    const xml = readPart(next, SLIDE_1)!
    expect(xml).toContain('<a:t>First shape</a:t>')
    expect(xml).toContain('<a:t>Edited second only</a:t>')
    expect(xml).not.toContain('<a:t>Second shape</a:t>')

    const [after] = await slides(next)
    expect(after.shapes.find((s) => s.kind === 'text' && s.sourceId === '@0')).toMatchObject({ text: 'First shape' })
    expect(after.shapes.find((s) => s.kind === 'text' && s.sourceId === '@1')).toMatchObject({
      text: 'Edited second only',
    })
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
    const reopened = await loadOfficePackage((await writeOfficePackage(moved)).buffer as ArrayBuffer)
    const deck = await slides(reopened)
    expect(deck.map((s) => s.title)).toEqual(['Agenda', 'Quarterly review'])
    expect(reopened.parts.get('ppt/media/image1.png')).toEqual(pkg.parts.get('ppt/media/image1.png'))
  })

  // SHELL-5 — "New Slide" on a .pptx with no slideLayout at all (non-conforming,
  // but a real file — see buildLayoutlessDeck) used to be a silent no-op:
  // layoutForNewSlide found an empty `layouts` list and returned null, and
  // addSlide's `if (!layout) return { pkg, index: afterIndex }` handed back
  // the SAME package and the SAME index, with no error anywhere. It must now
  // synthesize a minimal layout (once) and actually add the slide.
  it('SHELL-5 — adds a slide to a deck with no slideLayout at all, by synthesizing a minimal one', async () => {
    const layoutless = await loadOfficePackage(await buildLayoutlessDeck())
    expect([...layoutless.parts.keys()].some((path) => path.startsWith('ppt/slideLayouts/'))).toBe(false)

    const { pkg: next, index } = addSlide(layoutless, 0)
    expect(index).toBe(1) // not still 0 — a real slide was actually added.

    const deck = await slides(next)
    expect(deck).toHaveLength(2)
    expect(deck[0].shapes.some((s) => s.kind === 'text' && s.text === 'Only slide')).toBe(true)

    // The synthesized master/layout/theme are real, registered parts, not just
    // XML floating in the zip.
    expect([...next.parts.keys()].some((path) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(path))).toBe(true)
    expect([...next.parts.keys()].some((path) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(path))).toBe(true)
    expect([...next.parts.keys()].some((path) => /^ppt\/theme\/theme\d+\.xml$/.test(path))).toBe(true)
    expect(readPart(next, 'ppt/presentation.xml')).toContain('<p:sldMasterIdLst>')

    // Adding a SECOND slide reuses the same synthesized layout instead of
    // creating another one.
    const { pkg: twice } = addSlide(next, 1)
    const layoutParts = [...twice.parts.keys()].filter((path) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(path))
    expect(layoutParts).toHaveLength(1)
    expect((await slides(twice))).toHaveLength(3)
  })
})
