import { describe, expect, it } from 'vitest'

import {
  clampPageNumber,
  computeActivePageWindow,
  includePendingJumpTarget,
  pickMostVisiblePage,
  updateVisibilityRatios,
} from '../virtualization'

describe('computeActivePageWindow', () => {
  it('expands each intersecting page by the overscan amount', () => {
    const active = computeActivePageWindow([10], 20, 2)
    expect([...active].sort((a, b) => a - b)).toEqual([8, 9, 10, 11, 12])
  })

  it('clamps the window to [1, pageCount] at the document edges', () => {
    const start = computeActivePageWindow([1], 20, 3)
    expect([...start].sort((a, b) => a - b)).toEqual([1, 2, 3, 4])

    const end = computeActivePageWindow([20], 20, 3)
    expect([...end].sort((a, b) => a - b)).toEqual([17, 18, 19, 20])
  })

  it('unions overlapping windows from multiple intersecting pages', () => {
    const active = computeActivePageWindow([5, 6], 20, 1)
    expect([...active].sort((a, b) => a - b)).toEqual([4, 5, 6, 7])
  })

  it('returns an empty set for a zero-page document', () => {
    expect(computeActivePageWindow([1], 0, 2).size).toBe(0)
  })

  it('renders only a bounded window regardless of document length (PDF-01)', () => {
    // A 500-page document with only page 250 intersecting should never
    // activate anywhere near all 500 pages.
    const active = computeActivePageWindow([250], 500, 2)
    expect(active.size).toBe(5)
  })
})

describe('includePendingJumpTarget', () => {
  it('adds an out-of-window jump target to the active set', () => {
    const active = computeActivePageWindow([1], 100, 2)
    const withTarget = includePendingJumpTarget(active, 80, 100)
    expect(withTarget.has(80)).toBe(true)
    // Original pages are preserved.
    expect(withTarget.has(1)).toBe(true)
  })

  it('is a no-op when the target is already active', () => {
    const active = computeActivePageWindow([10], 100, 2)
    const result = includePendingJumpTarget(active, 10, 100)
    expect(result).toBe(active)
  })

  it('ignores an out-of-range or null pending page', () => {
    const active = computeActivePageWindow([1], 100, 2)
    expect(includePendingJumpTarget(active, null, 100)).toBe(active)
    expect(includePendingJumpTarget(active, 0, 100)).toBe(active)
    expect(includePendingJumpTarget(active, 101, 100)).toBe(active)
  })
})

describe('updateVisibilityRatios', () => {
  it('records ratios for intersecting entries', () => {
    const result = updateVisibilityRatios(new Map(), [
      { pageNumber: 1, ratio: 0.4, isIntersecting: true },
      { pageNumber: 2, ratio: 0.9, isIntersecting: true },
    ])
    expect(result.get(1)).toBe(0.4)
    expect(result.get(2)).toBe(0.9)
  })

  it('removes a page once it stops intersecting instead of resetting everything', () => {
    const previous = new Map([
      [1, 0.9],
      [2, 0.3],
    ])
    const result = updateVisibilityRatios(previous, [
      { pageNumber: 1, ratio: 0, isIntersecting: false },
    ])
    expect(result.has(1)).toBe(false)
    // Page 2 wasn't in this batch at all — its last-known ratio survives
    // (PDF-08: the old code reset this to 0 every callback).
    expect(result.get(2)).toBe(0.3)
  })

  it('does not mutate the previous map', () => {
    const previous = new Map([[1, 0.5]])
    updateVisibilityRatios(previous, [{ pageNumber: 1, ratio: 0.9, isIntersecting: true }])
    expect(previous.get(1)).toBe(0.5)
  })
})

describe('pickMostVisiblePage', () => {
  it('picks the page with the highest ratio', () => {
    const ratios = new Map([
      [1, 0.2],
      [2, 0.8],
      [3, 0.5],
    ])
    expect(pickMostVisiblePage(ratios, 1)).toBe(2)
  })

  it('falls back when nothing is currently visible', () => {
    expect(pickMostVisiblePage(new Map(), 7)).toBe(7)
  })
})

describe('clampPageNumber', () => {
  it('clamps into [1, pageCount]', () => {
    expect(clampPageNumber(-5, 10)).toBe(1)
    expect(clampPageNumber(0, 10)).toBe(1)
    expect(clampPageNumber(50, 10)).toBe(10)
    expect(clampPageNumber(5, 10)).toBe(5)
  })

  it('truncates fractional input', () => {
    expect(clampPageNumber(4.9, 10)).toBe(4)
  })

  it('returns 1 for a non-finite input or an empty document', () => {
    expect(clampPageNumber(NaN, 10)).toBe(1)
    expect(clampPageNumber(5, 0)).toBe(1)
  })
})
