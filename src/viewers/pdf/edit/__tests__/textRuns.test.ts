import { describe, expect, it } from 'vitest'

import { guessFontStyle, layoutTextBox, textBoxCorners } from '../textLayout'
import { mergeLineRects, sampleRunColors } from '../textRuns'

describe('sampleRunColors', () => {
  function block(width: number, height: number, paint: (x: number, y: number) => [number, number, number]) {
    const data = new Uint8ClampedArray(width * height * 4)
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) {
        const [r, g, b] = paint(x, y)
        data.set([r, g, b, 255], (y * width + x) * 4)
      }
    }
    return { data, width, height }
  }

  it('finds the paper colour on the border and the ink in the middle', () => {
    const colors = sampleRunColors(block(20, 10, (x, y) => (x > 5 && x < 12 && y > 2 && y < 7 ? [200, 0, 0] : [240, 240, 220])))
    expect(colors.background.r).toBeCloseTo(240 / 255, 2)
    expect(colors.background.b).toBeCloseTo(220 / 255, 2)
    expect(colors.ink).toEqual({ r: 200 / 255, g: 0, b: 0 })
  })

  it('treats an unpainted (transparent) canvas as white paper with black ink', () => {
    expect(sampleRunColors({ data: new Uint8ClampedArray(4 * 4 * 4), width: 4, height: 4 })).toEqual({
      background: { r: 1, g: 1, b: 1 },
      ink: { r: 0, g: 0, b: 0 },
    })
  })

  it('picks white ink on a dark background with no glyph pixels', () => {
    expect(sampleRunColors(block(4, 4, () => [10, 10, 40])).ink).toEqual({ r: 1, g: 1, b: 1 })
  })
})

describe('mergeLineRects', () => {
  it('joins the pieces of one line and keeps separate lines apart', () => {
    const merged = mergeLineRects([
      { left: 10, top: 100, right: 50, bottom: 112 },
      { left: 52, top: 101, right: 90, bottom: 112 },
      { left: 10, top: 120, right: 60, bottom: 132 },
      { left: 0, top: 0, right: 0.5, bottom: 0.5 },
    ])
    expect(merged).toEqual([
      { left: 10, top: 100, right: 90, bottom: 112 },
      { left: 10, top: 120, right: 60, bottom: 132 },
    ])
  })
})

describe('text box layout', () => {
  it('sizes the box from its widest line and its line count', () => {
    const layout = layoutTextBox('ab\nabcd', 10, (line) => line.length * 5)
    expect(layout.lines).toEqual(['ab', 'abcd'])
    expect(layout.width).toBeCloseTo(20 + 4 + 10 * 0.3)
    expect(layout.baselines[1] - layout.baselines[0]).toBeCloseTo(12)
  })

  it('turns the box corners about its top-left anchor', () => {
    const [topLeft, topRight, , bottomLeft] = textBoxCorners({ x: 10, y: 10 }, 90, 20, 5)
    expect(topLeft).toEqual({ x: 10, y: 10 })
    expect(topRight.x).toBeCloseTo(10)
    expect(topRight.y).toBeCloseTo(30)
    expect(bottomLeft.x).toBeCloseTo(15)
    expect(bottomLeft.y).toBeCloseTo(10)
  })
})

describe('guessFontStyle', () => {
  it('maps PostScript names and generic families onto the standard fonts', () => {
    expect(guessFontStyle('serif', 'TimesNewRomanPS-BoldItalicMT')).toEqual({ family: 'times', bold: true, italic: true })
    expect(guessFontStyle('sans-serif', 'MicrosoftSansSerif')).toEqual({ family: 'helvetica', bold: false, italic: false })
    expect(guessFontStyle('monospace', undefined).family).toBe('courier')
    expect(guessFontStyle(undefined, 'Calibri-Bold').bold).toBe(true)
  })
})
