/** USR-16 — selection and resizing must follow a rotated shape, not its unrotated box. */
import { describe, expect, it } from 'vitest'

import { isPointInShape, resizeBox, rotateAround } from '../slideGeometry'
import type { SlideTransform } from '../SlideDeck.types'

const flat: SlideTransform = { x: 100, y: 100, w: 200, h: 100 }
const turned: SlideTransform = { ...flat, rotationDeg: 90 }

const round = (box: SlideTransform) => ({
  x: Math.round(box.x),
  y: Math.round(box.y),
  w: Math.round(box.w),
  h: Math.round(box.h),
})

describe('isPointInShape', () => {
  it('hit-tests an unrotated shape by its box', () => {
    expect(isPointInShape({ x: 150, y: 150 }, flat)).toBe(true)
    expect(isPointInShape({ x: 150, y: 260 }, flat)).toBe(false)
  })

  it('follows the shape when it is rotated', () => {
    // Rotated 90° about its centre (200,150), the shape covers x 150..250, y 50..250.
    expect(isPointInShape({ x: 200, y: 240 }, turned)).toBe(true) // under the shape once turned
    expect(isPointInShape({ x: 200, y: 240 }, flat)).toBe(false)
    expect(isPointInShape({ x: 290, y: 150 }, turned)).toBe(false) // inside the UNROTATED box only
    expect(isPointInShape({ x: 290, y: 150 }, flat)).toBe(true)
  })
})

describe('resizeBox', () => {
  it('moves a shape by the raw delta, rotation or not', () => {
    expect(round(resizeBox(flat, 'move', { x: 10, y: -5 }))).toMatchObject({ x: 110, y: 95, w: 200, h: 100 })
    expect(round(resizeBox(turned, 'move', { x: 10, y: -5 }))).toMatchObject({ x: 110, y: 95 })
  })

  it('grows an unrotated shape from the dragged edge', () => {
    expect(round(resizeBox(flat, 'e', { x: 40, y: 0 }))).toEqual({ x: 100, y: 100, w: 240, h: 100 })
    expect(round(resizeBox(flat, 'nw', { x: 20, y: 20 }))).toEqual({ x: 120, y: 120, w: 180, h: 80 })
  })

  it('never shrinks below the minimum size', () => {
    expect(resizeBox(flat, 'e', { x: -1000, y: 0 }).w).toBe(8)
  })

  it('resizes a rotated shape along ITS axes, keeping the opposite edge put', () => {
    // Dragging the east handle of a 90°-rotated shape: the pointer moves DOWN
    // the slide, and the shape's west edge (drawn at the top) must not move.
    const grown = resizeBox(turned, 'e', { x: 0, y: 40 })
    expect(round(grown)).toEqual({ x: 80, y: 120, w: 240, h: 100 })

    // The west edge (drawn at the top once rotated) must not have moved.
    const cornerBefore = rotateAround({ x: turned.x, y: turned.y }, { x: 200, y: 150 }, 90)
    const cornerAfter = rotateAround({ x: grown.x, y: grown.y }, { x: grown.x + grown.w / 2, y: grown.y + grown.h / 2 }, 90)
    expect(cornerAfter.x).toBeCloseTo(cornerBefore.x, 6)
    expect(cornerAfter.y).toBeCloseTo(cornerBefore.y, 6)
  })
})
