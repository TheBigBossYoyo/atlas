/**
 * T7 (DAT-17) — approximate page-count estimation for RTF/ODT.
 */
import { describe, expect, it } from 'vitest'

import { DEFAULT_PAGE_HEIGHT_PX, estimatePageCount, parseLengthToPx } from '../pageEstimate'

describe('parseLengthToPx', () => {
  it('converts centimeters to pixels at 96dpi', () => {
    expect(parseLengthToPx('2.54cm')).toBeCloseTo(96, 1)
  })

  it('converts inches to pixels at 96dpi', () => {
    expect(parseLengthToPx('1in')).toBe(96)
  })

  it('converts millimeters to pixels at 96dpi', () => {
    expect(parseLengthToPx('25.4mm')).toBeCloseTo(96, 1)
  })

  it('converts points to pixels at 96dpi', () => {
    expect(parseLengthToPx('72pt')).toBeCloseTo(96, 1)
  })

  it('passes pixels through unchanged', () => {
    expect(parseLengthToPx('150px')).toBe(150)
  })

  it('returns undefined for an absent value', () => {
    expect(parseLengthToPx(undefined)).toBeUndefined()
  })

  it('returns undefined for an unrecognized unit or malformed string', () => {
    expect(parseLengthToPx('wide')).toBeUndefined()
    expect(parseLengthToPx('10em')).toBeUndefined()
  })
})

describe('estimatePageCount', () => {
  it('returns 1 for a height at or under one page', () => {
    expect(estimatePageCount(0)).toBe(1)
    expect(estimatePageCount(500, 1000)).toBe(1)
    expect(estimatePageCount(1000, 1000)).toBe(1)
  })

  it('rounds to the nearest whole page for a multi-page document', () => {
    expect(estimatePageCount(2000, 1000)).toBe(2)
    expect(estimatePageCount(2600, 1000)).toBe(3)
    expect(estimatePageCount(2400, 1000)).toBe(2)
  })

  it('falls back to at least 1 page for a non-positive scroll height or page height', () => {
    expect(estimatePageCount(-5, 1000)).toBe(1)
    expect(estimatePageCount(1000, 0)).toBe(1)
    expect(estimatePageCount(1000, -1)).toBe(1)
  })

  it('uses DEFAULT_PAGE_HEIGHT_PX (US Letter, 96dpi) when no page height is given', () => {
    expect(estimatePageCount(DEFAULT_PAGE_HEIGHT_PX * 2)).toBe(2)
  })
})
