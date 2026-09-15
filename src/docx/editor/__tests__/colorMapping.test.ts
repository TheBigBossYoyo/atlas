import { describe, expect, it } from 'vitest'

import { parseCssColor, toHighlightColor, toNearestHighlightColor } from '../colorMapping'

describe('toNearestHighlightColor (USR-08)', () => {
  it('keeps exact highlight swatches', () => {
    expect(toNearestHighlightColor('#FFFF00')).toBe('yellow')
  })

  it('maps arbitrary picker colors to the closest highlight instead of null', () => {
    expect(toNearestHighlightColor('#123456')).toBe('darkBlue')
    expect(toNearestHighlightColor('#fbbc04')).toBe('yellow')
    expect(toNearestHighlightColor('#b0800f')).toBe('darkYellow')
    expect(toNearestHighlightColor('not-a-color')).toBeNull()
  })
})

describe('toHighlightColor', () => {
  it('maps a known swatch hex to its highlight name', () => {
    expect(toHighlightColor('#FFFF00')).toBe('yellow')
    expect(toHighlightColor('#ff0000')).toBe('red')
  })

  it('is case-insensitive and trims whitespace', () => {
    expect(toHighlightColor('  #FfFf00  ')).toBe('yellow')
  })

  it('returns null for a hex outside the fixed highlight palette', () => {
    expect(toHighlightColor('#123456')).toBeNull()
  })
})

describe('parseCssColor', () => {
  it('normalizes a 6-digit hex to lowercase', () => {
    expect(parseCssColor('#FF00AA')).toBe('#ff00aa')
  })

  it('expands a 3-digit hex shorthand', () => {
    expect(parseCssColor('#f0a')).toBe('#ff00aa')
  })

  it('converts rgb() to hex', () => {
    expect(parseCssColor('rgb(255, 0, 170)')).toBe('#ff00aa')
  })

  it('converts rgba() to hex, ignoring the alpha channel', () => {
    expect(parseCssColor('rgba(255, 0, 170, 0.5)')).toBe('#ff00aa')
  })

  it('clamps and rounds out-of-range/fractional rgb components', () => {
    expect(parseCssColor('rgb(300, -10, 127.6)')).toBe('#ff0080')
  })

  it('returns null for a named color, hsl(), or garbage input', () => {
    expect(parseCssColor('red')).toBeNull()
    expect(parseCssColor('hsl(0, 100%, 50%)')).toBeNull()
    expect(parseCssColor('not-a-color')).toBeNull()
  })
})
