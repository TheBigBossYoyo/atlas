/**
 * Tests for src/docx/fonts/canvasMetrics.ts
 *
 * jsdom has no real canvas backend, so `HTMLCanvasElement.prototype.getContext`
 * is stubbed with a deterministic fake `CanvasRenderingContext2D` whose
 * `measureText` derives a width purely from the currently-set `font` string
 * and the text length — enough to exercise canvasMetrics.ts's own logic
 * (segmenting, caching, cache invalidation) without needing real glyph
 * rasterization.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import {
  __resetCanvasMetricsCachesForTests,
  measureFragmentPt,
  measureLineMetricsPt,
} from '../canvasMetrics'

interface StubTextMetrics {
  readonly width: number
  readonly fontBoundingBoxAscent: number
  readonly fontBoundingBoxDescent: number
}

interface StubContext {
  readonly calls: string[]
  readonly ctx: CanvasRenderingContext2D
}

function parseSizePx(font: string): number {
  const match = /(\d+(?:\.\d+)?)px/.exec(font)
  return match ? Number(match[1]) : 100
}

function createStubContext(): StubContext {
  const calls: string[] = []
  let currentFont = ''

  const ctx = {
    get font(): string {
      return currentFont
    },
    set font(value: string) {
      currentFont = value
    },
    measureText(text: string): StubTextMetrics {
      calls.push(`${currentFont}::${text}`)
      const sizePx = parseSizePx(currentFont)
      return {
        width: text.length * (sizePx / 10),
        fontBoundingBoxAscent: sizePx * 0.8,
        fontBoundingBoxDescent: sizePx * 0.2,
      }
    },
  }

  return { calls, ctx: ctx as unknown as CanvasRenderingContext2D }
}

function installStubCanvas(): StubContext {
  const stub = createStubContext()
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(stub.ctx)
  return stub
}

describe('canvasMetrics', () => {
  const originalDevicePixelRatio = window.devicePixelRatio

  beforeEach(() => {
    __resetCanvasMetricsCachesForTests()
    window.devicePixelRatio = 1
  })

  afterEach(() => {
    vi.restoreAllMocks()
    __resetCanvasMetricsCachesForTests()
    window.devicePixelRatio = originalDevicePixelRatio
  })

  describe('small-caps synthetic sizing (DXP-17)', () => {
    it('measures an originally-lowercase fragment narrower than the same letters already capitalized', () => {
      installStubCanvas()
      const widthAlreadyCaps = measureFragmentPt('AB', 'Calibri', 'regular', 12, 0, 'none')

      __resetCanvasMetricsCachesForTests()
      installStubCanvas()
      const widthSmallCaps = measureFragmentPt('ab', 'Calibri', 'regular', 12, 0, 'smallCaps')

      expect(widthAlreadyCaps).not.toBeNull()
      expect(widthSmallCaps).not.toBeNull()
      // Both ultimately measure "AB" glyphs, but small-caps synthesizes the
      // lowercase-derived pair at a reduced size — narrower than full caps.
      expect(widthSmallCaps as number).toBeLessThan(widthAlreadyCaps as number)
      expect(widthSmallCaps as number).toBeCloseTo((widthAlreadyCaps as number) * 0.8, 5)
    })

    it('measures a mixed-case fragment as separate full-size and reduced-size segments', () => {
      const stub = installStubCanvas()
      measureFragmentPt('Ab', 'Calibri', 'regular', 12, 0, 'smallCaps')

      // "A" is already a capital (full size), "b" is lowercase-derived
      // (reduced size) — two distinct measureText calls, not one.
      expect(stub.calls).toEqual([
        `normal normal 100px Calibri::A`,
        `normal normal 80px Calibri::B`,
      ])
    })

    it('leaves digits and punctuation at full size in small-caps text', () => {
      const stub = installStubCanvas()
      measureFragmentPt('a1', 'Calibri', 'regular', 12, 0, 'smallCaps')

      expect(stub.calls).toEqual([
        `normal normal 80px Calibri::A`,
        `normal normal 100px Calibri::1`,
      ])
    })
  })

  describe('devicePixelRatio-aware cache invalidation (DXP-21)', () => {
    it('serves a cached width for repeated identical measurements under a stable devicePixelRatio', () => {
      const stub = installStubCanvas()

      measureFragmentPt('Hello', 'Calibri', 'regular', 12)
      measureFragmentPt('Hello', 'Calibri', 'regular', 12)

      expect(stub.calls).toHaveLength(1)
    })

    it('invalidates cached measurements when devicePixelRatio changes', () => {
      const stub = installStubCanvas()

      const first = measureFragmentPt('Hello', 'Calibri', 'regular', 12)
      expect(stub.calls).toHaveLength(1)

      window.devicePixelRatio = 2
      const second = measureFragmentPt('Hello', 'Calibri', 'regular', 12)

      expect(stub.calls).toHaveLength(2)
      // The stub's measured width formula is DPR-independent, so the two
      // values still agree — this test is about the cache actually being
      // re-queried, not about the returned number changing.
      expect(second).toEqual(first)
    })

    it('does not invalidate the cache across calls when devicePixelRatio is unchanged', () => {
      const stub = installStubCanvas()

      measureLineMetricsPt('Calibri', 'regular', 12)
      measureLineMetricsPt('Calibri', 'regular', 12)

      expect(stub.calls).toHaveLength(1)
    })
  })
})
