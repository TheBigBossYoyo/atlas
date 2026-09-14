import { describe, expect, it } from 'vitest'

import { applyRotationToDimensions, normalizeRotation, rotateClockwise, rotateCounterClockwise } from '../rotation'

describe('normalizeRotation', () => {
  it('passes through the four valid rotations', () => {
    expect(normalizeRotation(0)).toBe(0)
    expect(normalizeRotation(90)).toBe(90)
    expect(normalizeRotation(180)).toBe(180)
    expect(normalizeRotation(270)).toBe(270)
  })

  it('wraps values >= 360 and negative values into range', () => {
    expect(normalizeRotation(360)).toBe(0)
    expect(normalizeRotation(450)).toBe(90)
    expect(normalizeRotation(-90)).toBe(270)
    expect(normalizeRotation(-360)).toBe(0)
  })

  it('falls back to 0 for a value not a multiple of 90', () => {
    expect(normalizeRotation(45)).toBe(0)
  })
})

describe('rotateClockwise / rotateCounterClockwise', () => {
  it('steps forward by 90 and wraps past 270', () => {
    expect(rotateClockwise(0)).toBe(90)
    expect(rotateClockwise(270)).toBe(0)
  })

  it('steps backward by 90 and wraps past 0', () => {
    expect(rotateCounterClockwise(90)).toBe(0)
    expect(rotateCounterClockwise(0)).toBe(270)
  })
})

describe('applyRotationToDimensions', () => {
  it('passes dimensions through unchanged at 0deg and 180deg', () => {
    expect(applyRotationToDimensions(300, 500, 0)).toEqual({ width: 300, height: 500 })
    expect(applyRotationToDimensions(300, 500, 180)).toEqual({ width: 300, height: 500 })
  })

  it('swaps width/height at 90deg and 270deg, matching pdf.js viewport swap (PDF-09/PDF-12 interaction)', () => {
    // A fit-width/fit-page scale must be computed against these swapped
    // dimensions once a page is rotated sideways, or the rendered page
    // (whose actual viewport pdf.js already swaps) won't match the fitted
    // container size.
    expect(applyRotationToDimensions(300, 500, 90)).toEqual({ width: 500, height: 300 })
    expect(applyRotationToDimensions(300, 500, 270)).toEqual({ width: 500, height: 300 })
  })
})
