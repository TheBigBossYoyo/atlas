import { describe, expect, it } from 'vitest'

import { normalizeRotation, rotateClockwise, rotateCounterClockwise } from '../rotation'

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
