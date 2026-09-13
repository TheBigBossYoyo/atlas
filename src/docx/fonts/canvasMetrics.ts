import type { FontMetrics } from './metrics'
import type { FontVariant } from './families'

/**
 * Wrap a TTF-derived FontMetrics so that per-glyph advance widths come from
 * what the browser will ACTUALLY paint (via Canvas2D `measureText`) instead
 * of the bundled TTF tables.
 *
 * Why: Even when we register a Word font name to a metric-compatible substitute
 * (e.g. Calibri → Carlito) via @font-face, the browser's text engine applies
 * additional shaping rules (kerning, ligatures, OS fallback for glyphs missing
 * from the substitute) whose advance widths can differ from the raw TTF hmtx.
 * Measuring on canvas lets us account for the browser's real rendering pipeline.
 *
 * Vertical metrics (ascender, descender, lineGap, xHeight, capHeight, unitsPerEm)
 * are preserved from the TTF — those drive line box heights and don't observably
 * differ between TTF metrics and canvas measurements.
 *
 * Performance: a single shared Canvas2D context is reused across all
 * measurements, and per-call results are memoized by composite key. Large
 * DOCX files (1000+ paragraphs, 200+ tables) would otherwise allocate a
 * fresh canvas per measureFragmentPt / measureLineMetricsPt invocation and
 * recompute identical fragment widths thousands of times, blocking the
 * renderer thread for tens of seconds (observed on the LaBoetie 3.0.6
 * "nothing loaded" regression).
 */

const REFERENCE_SIZE_PX = 100
const MAX_FRAGMENT_CACHE_ENTRIES = 20_000

type CanvasContext = CanvasRenderingContext2D

let sharedContext: CanvasContext | null | undefined
const fragmentWidthCache = new Map<string, number>()
const lineMetricsCache = new Map<string, CanvasLineMetricsPt | null>()

function getCanvasContext(): CanvasContext | null {
  if (sharedContext !== undefined) {
    return sharedContext
  }
  if (typeof document === 'undefined') {
    sharedContext = null
    return sharedContext
  }
  const canvas = document.createElement('canvas')
  sharedContext = canvas.getContext('2d')
  return sharedContext
}

function buildFontShorthand(family: string, variant: FontVariant): string {
  const weight = variant === 'bold' || variant === 'boldItalic' ? 'bold' : 'normal'
  const style = variant === 'italic' || variant === 'boldItalic' ? 'italic' : 'normal'
  const quoted = /\s/.test(family) ? `"${family}"` : family
  return `${style} ${weight} ${REFERENCE_SIZE_PX}px ${quoted}`
}

function cacheFragmentWidth(key: string, value: number): number {
  if (fragmentWidthCache.size >= MAX_FRAGMENT_CACHE_ENTRIES) {
    fragmentWidthCache.clear()
  }
  fragmentWidthCache.set(key, value)
  return value
}

/**
 * Test-only hook: clear all measurement caches. Production code never calls
 * this; we expose it so unit tests can isolate runs.
 */
export function __resetCanvasMetricsCachesForTests(): void {
  fragmentWidthCache.clear()
  lineMetricsCache.clear()
  sharedContext = undefined
}

/**
 * Wrap `base` so its `advanceWidth(codepoint)` returns a canvas-measured
 * width re-scaled into TTF font units. The downstream `measureRun` formula
 * `(advance / unitsPerEm) * sizePt` then yields the correct point width.
 *
 * Falls back transparently to the TTF metric when no DOM/canvas is available
 * (Node test environment) so unit tests keep working.
 */
export function wrapWithCanvasAdvance(
  base: FontMetrics,
  family: string,
  variant: FontVariant,
): FontMetrics {
  const ctx = getCanvasContext()
  if (ctx === null) {
    return base
  }
  const fontShorthand = buildFontShorthand(family, variant)

  const cache = new Map<number, number>()
  // Pre-compute a representative fallback in canvas units so a missing-glyph
  // codepoint can still return a sensible non-zero advance (matches the
  // fallback policy in metrics.ts createMetrics).
  ctx.font = fontShorthand
  const fallbackProbe = ctx.measureText('o').width
  const fallbackCanvasPx = Number.isFinite(fallbackProbe) && fallbackProbe > 0
    ? fallbackProbe
    : REFERENCE_SIZE_PX / 2
  const canvasPxToFontUnits = (px: number) => (px / REFERENCE_SIZE_PX) * base.unitsPerEm

  return {
    unitsPerEm: base.unitsPerEm,
    ascender: base.ascender,
    descender: base.descender,
    lineGap: base.lineGap,
    xHeight: base.xHeight,
    capHeight: base.capHeight,
    advanceWidth: (codepoint: number) => {
      const cached = cache.get(codepoint)
      if (cached !== undefined) {
        return cached
      }
      const char = String.fromCodePoint(codepoint)
      ctx.font = fontShorthand
      const measured = ctx.measureText(char).width
      const px = Number.isFinite(measured) && measured > 0 ? measured : fallbackCanvasPx
      const fontUnits = canvasPxToFontUnits(px)
      cache.set(codepoint, fontUnits)
      return fontUnits
    },
    hasGlyph: base.hasGlyph,
  }
}

/**
 * Browser-derived line metrics in points.
 *
 * Returns ascent/descent/height as the browser will actually render them,
 * derived from Canvas2D `fontBoundingBoxAscent` + `fontBoundingBoxDescent`
 * at REFERENCE_SIZE_PX then scaled to the requested point size.
 *
 * `null` when no DOM/canvas is available (Node test env) so callers can
 * fall back to TTF-derived metrics.
 */
export interface CanvasLineMetricsPt {
  readonly ascentPt: number
  readonly descentPt: number
  readonly heightPt: number
}

export function measureLineMetricsPt(
  family: string,
  variant: FontVariant,
  sizePt: number,
): CanvasLineMetricsPt | null {
  const ctx = getCanvasContext()
  if (ctx === null) {
    return null
  }

  const cacheKey = `${family}::${variant}::${sizePt}`
  const cached = lineMetricsCache.get(cacheKey)
  if (cached !== undefined) {
    return cached
  }

  ctx.font = buildFontShorthand(family, variant)
  // Sampling an alphabetic + descender + accented string yields the font's
  // own bounding box (Chromium returns family-level values regardless of
  // the sampled glyphs, but a non-empty sample is required).
  const sample = ctx.measureText('Hgëpqj')
  const ascentRefPx = sample.fontBoundingBoxAscent
  const descentRefPx = sample.fontBoundingBoxDescent

  if (
    typeof ascentRefPx !== 'number'
    || typeof descentRefPx !== 'number'
    || !Number.isFinite(ascentRefPx)
    || !Number.isFinite(descentRefPx)
  ) {
    lineMetricsCache.set(cacheKey, null)
    return null
  }

  const scale = sizePt / REFERENCE_SIZE_PX
  const ascentPt = ascentRefPx * scale
  const descentPt = descentRefPx * scale
  const result: CanvasLineMetricsPt = {
    ascentPt,
    descentPt,
    heightPt: ascentPt + descentPt,
  }
  lineMetricsCache.set(cacheKey, result)
  return result
}

/**
 * Measure a whole text fragment's painted width in points.
 *
 * Whole-string `measureText` preserves the browser's kerning, ligatures, and
 * OS fallback-shaping decisions — unlike summing per-codepoint advances which
 * loses inter-glyph adjustments and causes right-edge clipping.
 *
 * Manually applies `text-transform` (uppercase/lowercase/smallCaps) before
 * measurement because CSS transforms only affect rendering, not measurement
 * of the raw input string.
 *
 * `letterSpacingPt` is added once per codepoint after the additional
 * inter-glyph spacing CSS will inject at render time.
 *
 * `null` when no DOM/canvas is available so callers can fall back to
 * per-codepoint TTF measurement, OR when the browser returned a non-finite
 * width (defensive guard so downstream layout doesn't poison with NaN).
 */
export type CanvasTextTransform = 'none' | 'uppercase' | 'lowercase' | 'smallCaps'

export function measureFragmentPt(
  text: string,
  family: string,
  variant: FontVariant,
  sizePt: number,
  letterSpacingPt: number = 0,
  transform: CanvasTextTransform = 'none',
): number | null {
  const ctx = getCanvasContext()
  if (ctx === null) {
    return null
  }

  const cacheKey = `${family}::${variant}::${sizePt}::${letterSpacingPt}::${transform}::${text}`
  const cached = fragmentWidthCache.get(cacheKey)
  if (cached !== undefined) {
    return cached
  }

  ctx.font = buildFontShorthand(family, variant)

  const transformed = applyTransform(text, transform)
  const widthRefPx = ctx.measureText(transformed).width

  if (!Number.isFinite(widthRefPx)) {
    return null
  }

  const widthPt = (widthRefPx / REFERENCE_SIZE_PX) * sizePt

  if (letterSpacingPt === 0) {
    return cacheFragmentWidth(cacheKey, widthPt)
  }

  // letter-spacing inserts inter-glyph spacing per *grapheme*; approximate
  // with codepoint count which matches CSS behaviour for Latin / accented
  // text. Cluster-script callers already segment to graphemes upstream.
  const codepointCount = Array.from(transformed).length

  return cacheFragmentWidth(cacheKey, widthPt + letterSpacingPt * codepointCount)
}

function applyTransform(text: string, transform: CanvasTextTransform): string {
  if (transform === 'uppercase' || transform === 'smallCaps') {
    return text.toUpperCase()
  }
  if (transform === 'lowercase') {
    return text.toLowerCase()
  }
  return text
}
