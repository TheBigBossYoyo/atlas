import type { TtfTables } from './ttf'

export type FontMetrics = Readonly<{
  unitsPerEm: number
  ascender: number
  descender: number
  lineGap: number
  xHeight: number
  capHeight: number
  advanceWidth: (codepoint: number) => number
  hasGlyph: (codepoint: number) => boolean
}>

export type SerializedFontMetrics = Readonly<{
  unitsPerEm: number
  ascender: number
  descender: number
  lineGap: number
  xHeight: number
  capHeight: number
  widths: Readonly<Record<number, number>>
}>

const WIDTHS_BY_METRICS = new WeakMap<FontMetrics, ReadonlyMap<number, number>>()

export function buildMetrics(tables: TtfTables): FontMetrics {
  const widths = new Map<number, number>()

  for (const [codepoint, glyphId] of tables.cmap) {
    const advanceWidth = tables.hmtx.advanceWidths[glyphId]
    if (typeof advanceWidth === 'number') {
      widths.set(codepoint, advanceWidth)
    }
  }

  return createMetrics({
    unitsPerEm: tables.head.unitsPerEm,
    ascender: tables.os2.sTypoAscender,
    descender: tables.os2.sTypoDescender,
    lineGap: tables.hhea.lineGap,
    xHeight: tables.os2.sxHeight ?? 0,
    capHeight: tables.os2.sCapHeight ?? 0,
    widths,
  })
}

export function serializeMetrics(metrics: FontMetrics): SerializedFontMetrics {
  const widths = WIDTHS_BY_METRICS.get(metrics)
  if (!widths) {
    throw new Error('Unknown FontMetrics instance')
  }

  const serializedWidths: Record<number, number> = {}
  for (const [codepoint, width] of widths) {
    serializedWidths[codepoint] = width
  }

  return {
    unitsPerEm: metrics.unitsPerEm,
    ascender: metrics.ascender,
    descender: metrics.descender,
    lineGap: metrics.lineGap,
    xHeight: metrics.xHeight,
    capHeight: metrics.capHeight,
    widths: serializedWidths,
  }
}

export function deserializeMetrics(serialized: SerializedFontMetrics): FontMetrics {
  const widths = new Map<number, number>()
  for (const [codepoint, width] of Object.entries(serialized.widths)) {
    widths.set(Number(codepoint), width)
  }

  return createMetrics({
    unitsPerEm: serialized.unitsPerEm,
    ascender: serialized.ascender,
    descender: serialized.descender,
    lineGap: serialized.lineGap,
    xHeight: serialized.xHeight,
    capHeight: serialized.capHeight,
    widths,
  })
}

export function measureRun(text: string, metrics: FontMetrics, sizePt: number): number {
  let total = 0
  for (const glyph of text) {
    const codepoint = glyph.codePointAt(0)
    if (typeof codepoint === 'number') {
      total += (metrics.advanceWidth(codepoint) / metrics.unitsPerEm) * sizePt
    }
  }
  return total
}

function createMetrics(config: {
  unitsPerEm: number
  ascender: number
  descender: number
  lineGap: number
  xHeight: number
  capHeight: number
  widths: ReadonlyMap<number, number>
}): FontMetrics {
  // Pre-compute a sensible fallback width for codepoints absent from the cmap.
  // Returning 0 (the previous behaviour) collapses unknown glyphs (e.g. accented
  // French letters, smart quotes) to zero width, producing the "words overlap"
  // / "INTRODUCTION clipped" symptoms when the browser actually renders a
  // visible glyph via OS fallback. We prefer 'o' (0x6F) as a representative
  // lowercase advance, fall back to 'M' (0x4D) for caps-only fonts, then to
  // half the em as a last resort.
  const fallbackWidth =
    config.widths.get(0x6f) ??
    config.widths.get(0x4d) ??
    config.widths.get(0x20) ??
    config.unitsPerEm / 2

  const metrics: FontMetrics = {
    unitsPerEm: config.unitsPerEm,
    ascender: config.ascender,
    descender: config.descender,
    lineGap: config.lineGap,
    xHeight: config.xHeight,
    capHeight: config.capHeight,
    advanceWidth: (codepoint) => config.widths.get(codepoint) ?? fallbackWidth,
    hasGlyph: (codepoint) => config.widths.has(codepoint),
  }

  WIDTHS_BY_METRICS.set(metrics, config.widths)
  return metrics
}
