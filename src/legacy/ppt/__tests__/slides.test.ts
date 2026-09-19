/**
 * End-to-end test for extractLegacyPptSlides: builds a minimal but complete
 * in-memory .ppt (see `../../__tests__/fixtures.ts`) and verifies the whole
 * records/text/slides pipeline produces the right `SlideData[]`.
 */
import { describe, expect, it } from 'vitest'

import { MAX_LEGACY_SLIDES, extractLegacyPptSlides } from '../slides'
import { LegacyFormatError } from '../../errors'
import { buildCfbBytes, buildMinimalPptBytes } from '../../__tests__/fixtures'

describe('extractLegacyPptSlides', () => {
  it('extracts every slide, in document order, with title and body text', () => {
    const slides = extractLegacyPptSlides(
      buildMinimalPptBytes([
        { title: 'Slide One Title', body: 'Body line1\rBody line2' },
        { title: 'Slide Two Title' },
      ]),
    )

    expect(slides).toHaveLength(2)
    expect(slides[0].index).toBe(0)
    expect(slides[0].title).toBe('Slide One Title')
    expect(slides[1].index).toBe(1)
    expect(slides[1].title).toBe('Slide Two Title')
  })

  it('splits a text group on \\r into separate paragraphs', () => {
    const slides = extractLegacyPptSlides(
      buildMinimalPptBytes([{ title: 'Slide One Title', body: 'Body line1\rBody line2' }]),
    )
    const bodyShape = slides[0].shapes.find((shape) => shape.kind === 'text' && shape.text.includes('Body line1'))

    expect(bodyShape?.kind).toBe('text')
    if (bodyShape?.kind === 'text') {
      expect(bodyShape.paragraphs.map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['Body line1', 'Body line2'])
    }
  })

  it('gives every slide a fixed, non-zero width/height (no trustworthy geometry exists to read)', () => {
    const slides = extractLegacyPptSlides(buildMinimalPptBytes([{ title: 'Only Slide' }]))
    for (const slide of slides) {
      expect(slide.width).toBeGreaterThan(0)
      expect(slide.height).toBeGreaterThan(0)
    }
  })

  it('produces no slides for a presentation with no Slide containers, without throwing', () => {
    expect(extractLegacyPptSlides(buildMinimalPptBytes([]))).toEqual([])
  })

  it('throws LegacyFormatError when the "PowerPoint Document" stream is missing', () => {
    const bytes = buildCfbBytes([['SomethingElse', new Uint8Array([1])]])
    expect(() => extractLegacyPptSlides(bytes)).toThrow(LegacyFormatError)
  })

  it('throws LegacyFormatError for a file that is not a CFB container at all', () => {
    expect(() => extractLegacyPptSlides(new Uint8Array([9, 9, 9]))).toThrow(LegacyFormatError)
  })

  it('rejects a crafted deck with more slide containers than the cap instead of building them all', () => {
    // Empty RT_Slide containers (recVer 0xF, type 1006, length 0): 8 bytes each.
    const record = [0x0f, 0x00, 0xee, 0x03, 0, 0, 0, 0]
    const stream = new Uint8Array((MAX_LEGACY_SLIDES + 1) * record.length)
    for (let i = 0; i < MAX_LEGACY_SLIDES + 1; i++) stream.set(record, i * record.length)

    expect(() => extractLegacyPptSlides(buildCfbBytes([['PowerPoint Document', stream]]))).toThrow(LegacyFormatError)
  })
})
