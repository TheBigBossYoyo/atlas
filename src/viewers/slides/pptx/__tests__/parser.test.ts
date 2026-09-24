import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import type { SlideImage, SlideShapeOnly, SlideTable, SlideTextBox, SlideUnsupported } from '../../../shared/SlideDeck.types'
import type { CancelSignal, ZipArchive } from '../../shared/xmlUtils'
import { parsePptxSlides } from '../parser'
import { buildPptxFixtureZip } from './pptxFixture'

async function parseFixture() {
  const zip = buildPptxFixtureZip()
  const signal: CancelSignal = { cancelled: false }
  return parsePptxSlides(zip as unknown as ZipArchive, signal)
}

/** A minimal, real-world-shaped slide: a text shape with no spPr/xfrm and no placeholder at all. */
function buildMinimalPptxZip(): JSZip {
  const zip = new JSZip()
  zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>`)
  zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`)
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:sp><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Atlas PPTX fixture</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>
</p:sld>`)
  return zip
}

/**
 * SHELL-3 — a slide whose two shapes share the SAME `cNvPr id` (a lossy
 * converter or hand-built file can produce this just as easily as one with
 * no id at all — plain id-matching can no longer tell them apart either
 * way, so both fall back to positional addressing).
 */
function buildDuplicateIdPptxZip(): JSZip {
  const zip = new JSZip()
  zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>`)
  zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`)
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:sp><p:nvSpPr><p:cNvPr id="5" name="First"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>First shape</a:t></a:r></a:p></p:txBody></p:sp>
    <p:sp><p:nvSpPr><p:cNvPr id="5" name="Second"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Second shape</a:t></a:r></a:p></p:txBody></p:sp>
  </p:spTree></p:cSld>
</p:sld>`)
  return zip
}

describe('parsePptxSlides', () => {
  it('resolves the deck slide size from p:sldSz', async () => {
    const slides = await parseFixture()
    expect(slides[0]?.width).toBe(1280)
    expect(slides[0]?.height).toBe(720)
  })

  it('returns one SlideData per <p:sldId>, in order', async () => {
    const slides = await parseFixture()
    expect(slides).toHaveLength(3)
    expect(slides.map(slide => slide.index)).toEqual([0, 1, 2])
  })

  it('S1 — inherits a placeholder position down slide -> layout -> master when the slide omits its own xfrm', async () => {
    const [slide1] = await parseFixture()
    const title = slide1?.shapes.find(shape => shape.kind === 'text' && shape.placeholderType === 'title') as
      | SlideTextBox
      | undefined

    expect(title?.transform).toEqual({ x: 48, y: 30, w: 864, h: 120 })
  })

  it('S1 — a shape with no xfrm and no placeholder to inherit from still renders via generic stacking', async () => {
    const zip = buildMinimalPptxZip()
    const signal: CancelSignal = { cancelled: false }
    const [slide1] = await parsePptxSlides(zip as unknown as ZipArchive, signal)

    // Regression: previously such a shape was silently dropped rather than
    // falling back to a generic stacked position.
    const shape = slide1?.shapes.find(candidate => candidate.kind === 'text') as SlideTextBox | undefined
    expect(shape?.text).toBe('Atlas PPTX fixture')
    expect(shape?.transform).toEqual({ x: 24, y: 24, w: expect.any(Number), h: 24 })
  })

  // SHELL-3 — a shape with no <p:nvSpPr><p:cNvPr id="…"/></p:nvSpPr> at all (real
  // PowerPoint always writes one; hand-built or lossily-converted files may not)
  // used to get `sourceId: undefined`, which made it silently unselectable and
  // uneditable (SlideEditCanvas's hitTest requires a truthy sourceId). It must
  // now fall back to a truthy, structural-position address.
  it('SHELL-3 — a shape with no cNvPr id at all still gets an addressable (truthy) sourceId', async () => {
    const zip = buildMinimalPptxZip()
    const signal: CancelSignal = { cancelled: false }
    const [slide1] = await parsePptxSlides(zip as unknown as ZipArchive, signal)

    const shape = slide1?.shapes.find(candidate => candidate.kind === 'text') as SlideTextBox | undefined
    expect(shape?.sourceId).toBe('@0')
  })

  // SHELL-3 — same bug class: two shapes sharing one cNvPr id are just as
  // unaddressable by plain id-matching as no id at all (editing either one
  // would always hit whichever shares the id first in the XML). Both must
  // fall back to distinct, truthy, positional addresses.
  it('SHELL-3 — shapes sharing a duplicate cNvPr id fall back to distinct positional sourceIds', async () => {
    const zip = buildDuplicateIdPptxZip()
    const signal: CancelSignal = { cancelled: false }
    const [slide1] = await parsePptxSlides(zip as unknown as ZipArchive, signal)

    const [first, second] = (slide1?.shapes ?? []) as SlideTextBox[]
    expect(first?.text).toBe('First shape')
    expect(second?.text).toBe('Second shape')
    expect(first?.sourceId).toBe('@0')
    expect(second?.sourceId).toBe('@1')
    expect(first?.sourceId).not.toBe(second?.sourceId)
  })

  it('S3 — resolves run formatting via the master txStyles fallback and theme accent colors', async () => {
    const [slide1] = await parseFixture()
    const title = slide1?.shapes.find(shape => shape.kind === 'text' && shape.placeholderType === 'title') as
      | SlideTextBox
      | undefined
    const runs = title?.paragraphs[0]?.runs ?? []
    const firstRun = runs.find(run => run.text === 'Quarterly Report')
    const secondRun = runs.find(run => run.text === 'Q3 2024')

    // Own run rPr wins for color; master titleStyle's bold+size fall through for both runs.
    expect(firstRun).toMatchObject({ color: '#4472C4', bold: true })
    expect(firstRun?.fontSizePx).toBeCloseTo(44 * (96 / 72), 3)
    expect(secondRun).toMatchObject({ color: '#000000', bold: true })
  })

  it('S4 — resolves buChar bullets at their own paragraph level', async () => {
    const [slide1] = await parseFixture()
    const body = slide1?.shapes.find(shape => shape.kind === 'text' && shape.placeholderType === 'body') as
      | SlideTextBox
      | undefined

    expect(body?.paragraphs[0]?.bullet).toEqual({ level: 0, char: '-', numbered: undefined })
    expect(body?.paragraphs[1]?.bullet).toEqual({ level: 1, char: '-', numbered: undefined })
    // Master bodyStyle's per-level size fallback (lvl1pPr vs lvl2pPr).
    expect(body?.paragraphs[0]?.runs[0]?.fontSizePx).toBeCloseTo(24 * (96 / 72), 3)
    expect(body?.paragraphs[1]?.runs[0]?.fontSizePx).toBeCloseTo(20 * (96 / 72), 3)
  })

  it('S5 — a soft <a:br> line break becomes a real line, not "Quarterly ReportQ3 2024"', async () => {
    const [slide1] = await parseFixture()
    const title = slide1?.shapes.find(shape => shape.kind === 'text' && shape.placeholderType === 'title') as
      | SlideTextBox
      | undefined

    expect(title?.text).toBe('Quarterly Report\nQ3 2024')
    expect(title?.paragraphs[0]?.runs.map(run => run.text)).toEqual(['Quarterly Report', '\n', 'Q3 2024'])
  })

  it('S6 — renders a:tbl as a real row/cell grid, positioned at the graphicFrame box', async () => {
    const [slide1] = await parseFixture()
    const table = slide1?.shapes.find(shape => shape.kind === 'table') as SlideTable | undefined

    expect(table?.transform).toEqual({ x: 48, y: 450, w: 400, h: 100, rotationDeg: undefined, flipH: false, flipV: false })
    expect(table?.rows).toEqual([
      [{ text: 'Region', runs: [expect.objectContaining({ text: 'Region' })] }, { text: 'Growth', runs: [expect.objectContaining({ text: 'Growth' })] }],
      [{ text: 'EMEA', runs: [expect.objectContaining({ text: 'EMEA' })] }, { text: '18%', runs: [expect.objectContaining({ text: '18%' })] }],
    ])
  })

  it('S7 — composes a grouped shape\'s transform through the group\'s off/ext -> chOff/chExt scale', async () => {
    const [slide1] = await parseFixture()
    const grouped = slide1?.shapes.find(
      shape => shape.kind === 'shape' && shape.fill?.kind === 'solid' && shape.fill.color === '#00FF00',
    ) as SlideShapeOnly | undefined

    // group off=(100,600) ext=(200,200) chOff=(0,0) chExt=(400,400); child off=(0,0) ext=(200,200)
    // -> scale 0.5 -> absolute (100, 600, 100, 100).
    expect(grouped?.transform).toEqual({ x: 100, y: 600, w: 100, h: 100, rotationDeg: undefined, flipH: false, flipV: false })
  })

  it('S8 — resolves shape fill/geometry and the master\'s background', async () => {
    const [slide1] = await parseFixture()
    const banner = slide1?.shapes.find(
      shape => shape.kind === 'shape' && shape.fill?.kind === 'solid' && shape.fill.color === '#FF0000',
    ) as SlideShapeOnly | undefined

    expect(banner?.geometry).toBe('roundRect')
    expect(slide1?.background).toEqual({ kind: 'solid', color: '#FFFFFF' })
  })

  it('S8 — a shape\'s own solid fill is not shadowed by its <a:ln>\'s <a:noFill> (no border)', async () => {
    const [slide1] = await parseFixture()
    const banner = slide1?.shapes.find(shape => shape.kind === 'shape' && shape.geometry === 'roundRect') as
      | SlideShapeOnly
      | undefined

    // Regression: resolveFill previously searched the whole spPr subtree for
    // noFill/solidFill/gradFill, so the Banner's <a:ln><a:noFill/></a:ln>
    // (its border, not its fill) was found first and the shape's own
    // <a:solidFill> was never reached — the banner resolved as transparent
    // (and, having no text/border/non-rect geometry, was dropped entirely).
    expect(banner?.fill).toEqual({ kind: 'solid', color: '#FF0000' })
    expect(banner?.border).toBeUndefined()
  })

  it('S1 — a placeholder\'s own partial xfrm (rotation only) still merges in the layout\'s position/size', async () => {
    const [slide1] = await parseFixture()
    const rotated = slide1?.shapes.find(shape => shape.kind === 'text' && shape.text === 'Rotated placeholder') as
      | SlideTextBox
      | undefined

    // Regression: resolvePlaceholderPosition merged boxes via `{ ...layoutBox,
    // ...ownBox }`; since readXfrmBox always returns all four keys (even as
    // `undefined`), a shape whose own <a:xfrm> sets only `rot` (no off/ext)
    // had its spread wipe out the fully-specified layout box instead of
    // falling back to it, collapsing the shape to zero size.
    expect(rotated?.transform).toEqual({ x: 50, y: 50, w: 100, h: 100, rotationDeg: 90, flipH: false, flipV: false })
  })

  it('S9 — isolates one slide\'s parse failure instead of discarding the whole deck', async () => {
    const slides = await parseFixture()

    expect(slides[2]?.error).toBeTruthy()
    expect(slides[2]?.shapes).toEqual([])
    // The other, unrelated slides are unaffected.
    expect(slides[0]?.error).toBeUndefined()
    expect(slides[1]?.error).toBeUndefined()
  })

  it('S10/S19 — resolves image crop fractions, rotation, and alt text', async () => {
    const [slide1] = await parseFixture()
    const image = slide1?.shapes.find(shape => shape.kind === 'image') as SlideImage | undefined

    expect(image?.transform).toEqual({ x: 700, y: 100, w: 150, h: 150, rotationDeg: 45, flipH: false, flipV: false })
    expect(image?.crop).toEqual({ top: 0.2, right: 0.05, bottom: 0.15, left: 0.1 })
    expect(image?.alt).toBe('A quarterly chart')
    expect(image?.src.startsWith('data:image/png;base64,')).toBe(true)
  })

  it('S11 — a chart graphicFrame becomes a labeled "not supported" placeholder', async () => {
    const [slide1] = await parseFixture()
    const chart = slide1?.shapes.find(shape => shape.kind === 'unsupported') as SlideUnsupported | undefined

    expect(chart?.label).toBe('Chart not supported')
  })

  it('S12 — extracts the notesSlide relationship\'s body text, including a soft <a:br> line break', async () => {
    const [slide1] = await parseFixture()
    expect(slide1?.notes).toBe('Remember to mention EMEA growth drivers.\nFollow up with APAC next.')
  })

  it('S14 — a slide with show="0" is flagged hidden', async () => {
    const slides = await parseFixture()
    expect(slides[0]?.hidden).toBe(false)
    expect(slides[1]?.hidden).toBe(true)
  })

  it('S17 — extracts the title placeholder\'s text for nav labels', async () => {
    const [slide1] = await parseFixture()
    expect(slide1?.title).toBe('Quarterly Report Q3 2024')
  })

  it('S20 — carries normAutofit\'s fontScale through to the text box', async () => {
    const [slide1] = await parseFixture()
    const title = slide1?.shapes.find(shape => shape.kind === 'text' && shape.placeholderType === 'title') as
      | SlideTextBox
      | undefined

    expect(title?.fontScale).toBeCloseTo(0.925, 5)
  })

  it('S21 — resolves spcBef paragraph spacing in pixels', async () => {
    const [slide1] = await parseFixture()
    const body = slide1?.shapes.find(shape => shape.kind === 'text' && shape.placeholderType === 'body') as
      | SlideTextBox
      | undefined

    expect(body?.paragraphs[0]?.spaceBeforePx).toBeCloseTo(6 * (96 / 72), 3)
  })
})
