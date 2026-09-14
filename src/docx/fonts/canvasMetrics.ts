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

/**
 * Approximate size ratio browsers use when synthesizing `font-variant:
 * small-caps` for a font with no true small-caps OpenType feature (our
 * bundled substitutes have none) — the originally-lowercase letters are
 * upper-cased and painted at a reduced size relative to genuine capitals.
 * Real browsers vary (roughly 0.7-0.85 depending on engine); 0.8 is a
 * reasonable mid-range approximation. This is documented as an
 * approximation, not an exact match (DXP-17) — closing the gap completely
 * would require reading the live DOM's actual rendered small-caps metrics,
 * which is out of scope for this measurement-only module.
 */
const SMALL_CAPS_SYNTHETIC_SCALE = 0.8

type CanvasContext = CanvasRenderingContext2D

let sharedContext: CanvasContext | null | undefined
const fragmentWidthCache = new Map<string, number>()
const lineMetricsCache = new Map<string, CanvasLineMetricsPt | null>()

/**
 * Tracks the `devicePixelRatio` in effect when the caches above were last
 * populated (DXP-21). `measureText` itself is specified to return widths in
 * CSS-pixel/user-space units independent of a canvas's backing-store
 * resolution, so this deliberately does NOT rescale measured widths by
 * `devicePixelRatio` — doing so would be a guess dressed up as a fix, and
 * could silently corrupt every measurement the paginator depends on.
 *
 * What genuinely is a DPI-dependent hazard: Windows laptops routinely run
 * at 125%/150%/200% OS display scaling, and Electron windows are commonly
 * dragged between monitors with different scale factors while the app is
 * open. Some rasterizer backends apply DPI-aware hinting that can shift
 * sub-pixel glyph advances at the effective device resolution. Caching
 * measurements keyed only by family/variant/size (as before) means a
 * measurement taken under one display's DPI keeps being served, unchanged,
 * after the window moves to a display with a different DPI. Invalidating
 * the caches whenever `devicePixelRatio` changes ensures every measurement
 * reflects the DPI Atlas is actually painting under right now.
 */
let lastKnownDevicePixelRatio: number | undefined

function getDevicePixelRatio(): number {
  return typeof window !== 'undefined'
    && typeof window.devicePixelRatio === 'number'
    && Number.isFinite(window.devicePixelRatio)
    && window.devicePixelRatio > 0
    ? window.devicePixelRatio
    : 1
}

function invalidateCachesIfDevicePixelRatioChanged(): void {
  const currentDpr = getDevicePixelRatio()
  if (lastKnownDevicePixelRatio === undefined) {
    lastKnownDevicePixelRatio = currentDpr
    return
  }
  if (currentDpr !== lastKnownDevicePixelRatio) {
    fragmentWidthCache.clear()
    lineMetricsCache.clear()
    lastKnownDevicePixelRatio = currentDpr
  }
}

function getCanvasContext(): CanvasContext | null {
  invalidateCachesIfDevicePixelRatioChanged()

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

function buildFontShorthand(family: string, variant: FontVariant, sizeScale: number = 1): string {
  const weight = variant === 'bold' || variant === 'boldItalic' ? 'bold' : 'normal'
  const style = variant === 'italic' || variant === 'boldItalic' ? 'italic' : 'normal'
  const quoted = /\s/.test(family) ? `"${family}"` : family
  const sizePx = REFERENCE_SIZE_PX * sizeScale
  return `${style} ${weight} ${sizePx}px ${quoted}`
}

/**
 * Whether `ch` is a cased letter that renders lowercase in normal text (and
 * therefore gets synthesized down to a reduced-size capital under
 * `font-variant: small-caps`). Digits, punctuation, spaces, and already-
 * uppercase letters are unaffected by the small-caps size reduction.
 */
function isLowerCaseLetter(ch: string): boolean {
  return ch !== ch.toUpperCase()
}

/**
 * Measure small-caps text the way the browser actually paints it: split
 * into contiguous originally-lowercase / originally-not-lowercase segments,
 * upper-case each, and measure the lowercase-derived segments at the
 * reduced synthetic small-caps size while the rest stays full-size (DXP-17).
 * A single flat `ctx.measureText(text.toUpperCase())` call — the previous
 * behaviour — always measures at full size, overstating the width of any
 * text that mixes cases.
 */
function measureSmallCapsWidthPx(
  ctx: CanvasContext,
  text: string,
  fullSizeFontShorthand: string,
  reducedSizeFontShorthand: string,
): number {
  let width = 0
  let index = 0

  while (index < text.length) {
    const segmentIsLowerDerived = isLowerCaseLetter(text[index])
    let end = index + 1
    while (end < text.length && isLowerCaseLetter(text[end]) === segmentIsLowerDerived) {
      end += 1
    }

    ctx.font = segmentIsLowerDerived ? reducedSizeFontShorthand : fullSizeFontShorthand
    width += ctx.measureText(text.slice(index, end).toUpperCase()).width

    index = end
  }

  return width
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
  lastKnownDevicePixelRatio = undefined
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

  const fullSizeFontShorthand = buildFontShorthand(family, variant)
  let widthRefPx: number

  if (transform === 'smallCaps') {
    const reducedSizeFontShorthand = buildFontShorthand(family, variant, SMALL_CAPS_SYNTHETIC_SCALE)
    widthRefPx = measureSmallCapsWidthPx(ctx, text, fullSizeFontShorthand, reducedSizeFontShorthand)
  } else {
    ctx.font = fullSizeFontShorthand
    const transformed = applyTransform(text, transform)
    widthRefPx = ctx.measureText(transformed).width
  }

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
  // Counted against the original `text`, not a transformed copy — none of
  // uppercase/lowercase/smallCaps change codepoint count.
  const codepointCount = Array.from(text).length

  return cacheFragmentWidth(cacheKey, widthPt + letterSpacingPt * codepointCount)
}

/**
 * Applies `uppercase`/`lowercase`/`none`. `smallCaps` is handled separately
 * by {@link measureSmallCapsWidthPx} — it needs per-segment sizing, not a
 * single whole-string case transform (DXP-17).
 */
function applyTransform(text: string, transform: 'none' | 'uppercase' | 'lowercase'): string {
  if (transform === 'uppercase') {
    return text.toUpperCase()
  }
  if (transform === 'lowercase') {
    return text.toLowerCase()
  }
  return text
}
