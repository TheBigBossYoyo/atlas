import { describe, expect, it } from 'vitest'

import {
  MAX_ZOOM,
  MIN_ZOOM,
  clampZoom,
  computeFitScale,
  resolveScaleForPage,
  zoomIn,
  zoomOut,
} from '../geometry'

describe('clampZoom', () => {
  it('passes values already within range through unchanged', () => {
    expect(clampZoom(1)).toBe(1)
    expect(clampZoom(1.5)).toBe(1.5)
  })

  it('clamps below MIN_ZOOM up to MIN_ZOOM', () => {
    expect(clampZoom(0.01)).toBe(MIN_ZOOM)
  })

  it('clamps above MAX_ZOOM down to MAX_ZOOM', () => {
    expect(clampZoom(10)).toBe(MAX_ZOOM)
  })

  it('falls back to the default zoom for non-finite or non-positive input', () => {
    expect(clampZoom(0)).toBe(1)
    expect(clampZoom(-1)).toBe(1)
    expect(clampZoom(NaN)).toBe(1)
    expect(clampZoom(Infinity)).toBe(1)
  })
})

describe('computeFitScale', () => {
  it('fit-width scales purely from the container width, ignoring height', () => {
    // A 200x400 page fit-width into a 400-wide container should scale 2x,
    // regardless of how tall the container is.
    expect(computeFitScale(200, 400, 400, 100, 'fit-width')).toBe(2)
    expect(computeFitScale(200, 400, 400, 5000, 'fit-width')).toBe(2)
  })

  it('fit-page scales by whichever dimension is more constraining', () => {
    // 200x400 page into an 800x400 container: width scale=4, height scale=1
    // -> fit-page must pick the smaller (1) so the whole page is visible.
    expect(computeFitScale(200, 400, 800, 400, 'fit-page')).toBe(1)
  })

  it('computes per-page scale independently for mixed page sizes (PDF-09)', () => {
    // A portrait and a landscape page in the same fit-width container get
    // different scales, each fitted to its own width.
    const portraitScale = computeFitScale(300, 600, 900, 1000, 'fit-width')
    const landscapeScale = computeFitScale(600, 300, 900, 1000, 'fit-width')
    expect(portraitScale).toBe(3)
    expect(landscapeScale).toBe(1.5)
  })

  it('returns the default zoom for degenerate (zero/negative) inputs', () => {
    expect(computeFitScale(0, 400, 400, 400, 'fit-width')).toBe(1)
    expect(computeFitScale(200, 400, 0, 400, 'fit-width')).toBe(1)
  })

  it('clamps the computed fit scale to the documented zoom range', () => {
    expect(computeFitScale(10, 10, 10000, 10000, 'fit-width')).toBe(MAX_ZOOM)
    expect(computeFitScale(10000, 10000, 10, 10, 'fit-page')).toBe(MIN_ZOOM)
  })
})

describe('resolveScaleForPage', () => {
  it('dispatches fit-width/fit-page to computeFitScale', () => {
    expect(resolveScaleForPage('fit-width', 200, 400, 400, 100)).toBe(2)
  })

  it('treats a numeric zoom mode as a literal (clamped) scale, ignoring page/container size', () => {
    expect(resolveScaleForPage(1.5, 200, 400, 1, 1)).toBe(1.5)
    expect(resolveScaleForPage(10, 200, 400, 1, 1)).toBe(MAX_ZOOM)
  })
})

describe('zoomIn / zoomOut', () => {
  it('steps by the fixed increment and clamps at the bounds', () => {
    expect(zoomIn(1)).toBe(1.25)
    expect(zoomOut(1)).toBe(0.75)
    expect(zoomIn(MAX_ZOOM)).toBe(MAX_ZOOM)
    expect(zoomOut(MIN_ZOOM)).toBe(MIN_ZOOM)
  })
})
