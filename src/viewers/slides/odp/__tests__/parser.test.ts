import { describe, expect, it } from 'vitest'

import type { SlideImage, SlideShapeOnly, SlideTable, SlideTextBox, SlideUnsupported } from '../../../shared/SlideDeck.types'
import type { CancelSignal, ZipArchive } from '../../shared/xmlUtils'
import { parseOdpSlides } from '../parser'
import { buildOdpFixtureZip } from './odpFixture'

async function parseFixture() {
  const zip = buildOdpFixtureZip()
  const signal: CancelSignal = { cancelled: false }
  return parseOdpSlides(zip as unknown as ZipArchive, signal)
}

function findByTransform<T extends { transform: { x: number; y: number } }>(
  shapes: ReadonlyArray<T>,
  x: number,
  y: number,
): T | undefined {
  return shapes.find(shape => shape.transform.x === x && shape.transform.y === y)
}

describe('parseOdpSlides', () => {
  it('S2 — resolves slide size from styles.xml\'s master page -> page layout', async () => {
    const slides = await parseFixture()
    expect(slides[0]?.width).toBe(1280)
    expect(slides[0]?.height).toBe(720)
  })

  it('returns one SlideData per draw:page, in order', async () => {
    const slides = await parseFixture()
    expect(slides).toHaveLength(2)
    expect(slides.map(slide => slide.index)).toEqual([0, 1])
  })

  it('S3/S5 — resolves text:span run formatting and a text:line-break soft break', async () => {
    const [slide1] = await parseFixture()
    const title = slide1?.shapes.find(shape => shape.kind === 'text' && shape.transform.x === 40 && shape.transform.y === 20) as
      | SlideTextBox
      | undefined

    expect(title?.text).toBe('Quarterly Report\nQ3 2024')
    expect(title?.paragraphs[0]?.runs).toEqual([
      expect.objectContaining({ text: 'Quarterly Report', bold: true, color: '#4472C4' }),
      { text: '\n' },
      { text: 'Q3 2024' },
    ])
  })

  it('S4 — resolves list-level bullets across nested text:list/text:list-item', async () => {
    const [slide1] = await parseFixture()
    const outline = slide1?.shapes.find(shape => shape.kind === 'text' && shape.transform.x === 40 && shape.transform.y === 140) as
      | SlideTextBox
      | undefined

    expect(outline?.paragraphs).toHaveLength(2)
    expect(outline?.paragraphs[0]).toMatchObject({ level: 0, bullet: { level: 0, char: '•' } })
    expect(outline?.paragraphs[1]).toMatchObject({ level: 1, bullet: { level: 1, char: '◦' } })
    // The P1 paragraph style's own fo:font-size applies to its un-spanned run text.
    expect(outline?.paragraphs[0]?.runs[0]?.fontSizePx).toBe(24)
  })

  it('S6 — renders table:table as a real row/cell grid', async () => {
    const [slide1] = await parseFixture()
    const table = slide1?.shapes.find(shape => shape.kind === 'table') as SlideTable | undefined

    expect(table?.transform).toEqual({ x: 40, y: 360, w: 400, h: 100 })
    expect(table?.rows.map(row => row.map(cell => cell.text))).toEqual([
      ['Region', 'Growth'],
      ['EMEA', '18%'],
    ])
  })

  it('S7 — composes a draw:g/draw:transform translate offset onto its children', async () => {
    const [slide1] = await parseFixture()
    const rect = findByTransform(
      slide1?.shapes.filter((shape): shape is SlideShapeOnly => shape.kind === 'shape') ?? [],
      100,
      500,
    )

    expect(rect?.transform).toEqual({ x: 100, y: 500, w: 100, h: 60 })
  })

  it('S8 — resolves a bare shape\'s fill/border and the master page\'s background', async () => {
    const [slide1] = await parseFixture()
    const rect = findByTransform(
      slide1?.shapes.filter((shape): shape is SlideShapeOnly => shape.kind === 'shape') ?? [],
      100,
      500,
    )

    expect(rect?.fill).toEqual({ kind: 'solid', color: '#FF0000' })
    expect(rect?.border).toEqual({ color: '#000000', widthPx: 2 })
    expect(slide1?.background).toEqual({ kind: 'solid', color: '#FFFFFF' })
  })

  it('S11 — a draw:object frame becomes a labeled "not supported" placeholder', async () => {
    const [slide1] = await parseFixture()
    const object = slide1?.shapes.find(shape => shape.kind === 'unsupported') as SlideUnsupported | undefined

    expect(object?.label).toBe('Embedded object not supported')
  })

  it('S12 — extracts the presentation:notes body text', async () => {
    const [slide1] = await parseFixture()
    expect(slide1?.notes).toBe('Remember to mention EMEA growth drivers.')
  })

  it('S14 — presentation:visibility="hidden" flags a slide hidden', async () => {
    const slides = await parseFixture()
    expect(slides[0]?.hidden).toBe(false)
    expect(slides[1]?.hidden).toBe(true)
  })

  it('S17 — extracts the presentation:class="title" frame\'s text for nav labels', async () => {
    const [slide1] = await parseFixture()
    expect(slide1?.title).toBe('Quarterly Report\nQ3 2024')
  })

  it('S19 — reads svg:title as the image\'s alt text', async () => {
    const [slide1] = await parseFixture()
    const image = slide1?.shapes.find(shape => shape.kind === 'image') as SlideImage | undefined

    expect(image?.alt).toBe('A quarterly chart')
    expect(image?.transform).toEqual({ x: 500, y: 360, w: 150, h: 150 })
    expect(image?.src.startsWith('data:image/png;base64,')).toBe(true)
  })
})
