import { measureFragmentPt, measureRun, type CanvasTextTransform, type FontMetrics } from '../fonts'
import type { BreakType } from '../model'
import type { Theme } from '../parser/theme'

import { effectiveFontFamily } from './fontResolution'
import type { EffectiveRunProps, FontResolver, LineBreakInput, LineItem } from './types'

const DEFAULT_DRAWING_SIZE_PT = 96
const DEFAULT_FONT_SIZE_PT = 11
const EMUS_PER_POINT = 12700
const NO_BREAK_SPACE = '\u00a0'
const SOFT_HYPHEN = '\u00ad'

type FontVariant = 'regular' | 'bold' | 'italic' | 'boldItalic'

type SegmenterConstructor = new (
  locales?: string | ReadonlyArray<string>,
  options?: { granularity: 'grapheme' },
) => {
  segment(input: string): Iterable<{ segment: string }>
}

export async function itemizeRuns(
  runs: LineBreakInput['runs'],
  fontResolver: FontResolver,
  theme?: Theme,
): Promise<ReadonlyArray<LineItem>> {
  const items: LineItem[] = []
  const metricsCache = new Map<string, Promise<FontMetrics>>()

  for (const [runIndex, wrappedRun] of runs.entries()) {
    let childOffset = 0

    for (const child of wrappedRun.run.children) {
      if (child.kind === 'text') {
        items.push(
          ...(await itemizeText(
            child.value,
            runIndex,
            childOffset,
            wrappedRun.runProps,
            fontResolver,
            metricsCache,
            theme,
          )),
        )
        childOffset += child.value.length
        continue
      }

      if (child.kind === 'tab') {
        items.push({
          kind: 'tab',
          width: 0,
          runIndex,
          charOffset: childOffset,
        })
        childOffset += 1
        continue
      }

      if (child.kind === 'break') {
        items.push({
          kind: 'break',
          breakKind: normalizeBreakKind(child.breakType),
          runIndex,
        })
        childOffset += 1
        continue
      }

      if (child.kind === 'drawing') {
        // Anchored drawings are painted at their anchor point for now; true
        // floating placement (positionH/V, wrap modes) is not modelled yet.
        items.push({
          kind: 'drawing',
          drawing: child,
          width: resolveDrawingSize(child.extent?.cx),
          height: resolveDrawingSize(child.extent?.cy),
          runIndex,
          charStart: childOffset,
          charEnd: childOffset + 1,
          runProps: wrappedRun.runProps,
        })
        childOffset += 1
      }
    }
  }

  return items
}

async function itemizeText(
  text: string,
  runIndex: number,
  baseOffset: number,
  runProps: EffectiveRunProps,
  fontResolver: FontResolver,
  metricsCache: Map<string, Promise<FontMetrics>>,
  theme: Theme | undefined,
): Promise<ReadonlyArray<LineItem>> {
  const items: LineItem[] = []
  let cursor = 0
  let fragmentStart: number | undefined
  let fragmentText = ''

  while (cursor < text.length) {
    const codepoint = text.codePointAt(cursor)
    if (typeof codepoint !== 'number') {
      break
    }

    const glyph = String.fromCodePoint(codepoint)
    const step = glyph.length

    if (glyph === SOFT_HYPHEN) {
      const hadLeadingFragment = fragmentText.length > 0 && fragmentStart !== undefined
      if (hadLeadingFragment && fragmentStart !== undefined) {
        items.push(
          ...(await createTextItems(
            fragmentText,
            runIndex,
            baseOffset + fragmentStart,
            baseOffset + cursor,
            runProps,
            fontResolver,
            metricsCache,
            theme,
          )),
        )
      }

      if (hadLeadingFragment && hasTrailingFragment(text, cursor + step)) {
        items.push({
          kind: 'hyphen-opportunity',
          text: '',
          penaltyWidth: await measureText('-', runProps, '-', fontResolver, metricsCache, theme),
          runIndex,
          charOffset: baseOffset + cursor,
        })
      }

      fragmentStart = undefined
      fragmentText = ''
      cursor += step
      continue
    }

    if (isBreakableSpace(glyph)) {
      if (fragmentText.length > 0 && fragmentStart !== undefined) {
        items.push(
          ...(await createTextItems(
            fragmentText,
            runIndex,
            baseOffset + fragmentStart,
            baseOffset + cursor,
            runProps,
            fontResolver,
            metricsCache,
            theme,
          )),
        )
      }

      items.push({
        kind: 'space',
        width: await measureText(glyph, runProps, glyph, fontResolver, metricsCache, theme),
        stretchable: true,
        runIndex,
        charOffset: baseOffset + cursor,
      })

      fragmentStart = undefined
      fragmentText = ''
      cursor += step
      continue
    }

    if (fragmentStart === undefined) {
      fragmentStart = cursor
    }

    fragmentText += glyph
    cursor += step
  }

  if (fragmentText.length > 0 && fragmentStart !== undefined) {
    items.push(
      ...(await createTextItems(
        fragmentText,
        runIndex,
        baseOffset + fragmentStart,
        baseOffset + text.length,
        runProps,
        fontResolver,
        metricsCache,
        theme,
      )),
    )
  }

  return items
}

async function createTextItems(
  text: string,
  runIndex: number,
  charStart: number,
  charEnd: number,
  runProps: EffectiveRunProps,
  fontResolver: FontResolver,
  metricsCache: Map<string, Promise<FontMetrics>>,
  theme: Theme | undefined,
): Promise<ReadonlyArray<LineItem>> {
  if (!containsClusterScript(text)) {
    return [
      {
        kind: 'word',
        text,
        width: await measureText(text, runProps, text, fontResolver, metricsCache, theme),
        runIndex,
        charStart,
        charEnd,
        runProps,
      },
    ]
  }

  const items: LineItem[] = []
  let clusterOffset = charStart

  for (const cluster of segmentGraphemes(text)) {
    items.push({
      kind: 'glyph-cluster',
      text: cluster,
      width: await measureText(cluster, runProps, cluster, fontResolver, metricsCache, theme),
      runIndex,
      charStart: clusterOffset,
      charEnd: clusterOffset + cluster.length,
      runProps,
    })
    clusterOffset += cluster.length
  }

  return items
}

async function measureText(
  text: string,
  runProps: EffectiveRunProps,
  sampleText: string,
  fontResolver: FontResolver,
  metricsCache: Map<string, Promise<FontMetrics>>,
  theme: Theme | undefined,
): Promise<number> {
  const family = resolveFontFamily(runProps, sampleText, theme)
  const variant = resolveFontVariant(runProps)
  const sizePt = resolveFontSize(runProps)
  const letterSpacingPt = resolveLetterSpacingPt(runProps)
  const transform = resolveTextTransform(runProps)

  // Prefer browser-derived whole-fragment measurement (preserves kerning /
  // ligatures / OS fallback shaping). Falls back to per-codepoint TTF
  // measurement in Node test envs where canvas is unavailable.
  const fragmentPt = measureFragmentPt(text, family, variant, sizePt, letterSpacingPt, transform)
  if (fragmentPt !== null) {
    return fragmentPt
  }

  const cacheKey = `${family}::${variant}`
  let metricsPromise = metricsCache.get(cacheKey)

  if (!metricsPromise) {
    metricsPromise = fontResolver(family, variant)
    metricsCache.set(cacheKey, metricsPromise)
  }

  return measureRun(text, await metricsPromise, sizePt)
}

const TWIP_PER_POINT = 20

function resolveLetterSpacingPt(runProps: EffectiveRunProps): number {
  if (typeof runProps.spacing === 'number') {
    return runProps.spacing / TWIP_PER_POINT
  }
  return 0
}

function resolveTextTransform(runProps: EffectiveRunProps): CanvasTextTransform {
  if (runProps.smallCaps === true) {
    return 'smallCaps'
  }
  if (runProps.caps === true) {
    return 'uppercase'
  }
  return 'none'
}

function normalizeBreakKind(breakType: BreakType | undefined): 'line' | 'page' | 'column' {
  if (breakType === 'page' || breakType === 'column') {
    return breakType
  }

  return 'line'
}

function resolveDrawingSize(emus: number | undefined): number {
  if (typeof emus === 'number' && emus > 0) {
    return emus / EMUS_PER_POINT
  }

  return DEFAULT_DRAWING_SIZE_PT
}

function resolveFontVariant(runProps: EffectiveRunProps): FontVariant {
  if (runProps.bold && runProps.italic) {
    return 'boldItalic'
  }
  if (runProps.bold) {
    return 'bold'
  }
  if (runProps.italic) {
    return 'italic'
  }

  return 'regular'
}

function resolveFontFamily(
  runProps: EffectiveRunProps,
  text: string,
  theme: Theme | undefined,
): string {
  const hint = containsClusterScript(text) ? 'eastAsian' : runProps.rtl ? 'rtl' : 'latin'
  return effectiveFontFamily(runProps, theme, hint)
}

function resolveFontSize(runProps: EffectiveRunProps): number {
  if (typeof runProps.sz === 'number') {
    return runProps.sz / 2
  }

  if (typeof runProps.szCs === 'number') {
    return runProps.szCs / 2
  }

  return DEFAULT_FONT_SIZE_PT
}

function hasTrailingFragment(text: string, start: number): boolean {
  let cursor = start

  while (cursor < text.length) {
    const codepoint = text.codePointAt(cursor)
    if (typeof codepoint !== 'number') {
      return false
    }

    const glyph = String.fromCodePoint(codepoint)
    if (glyph === SOFT_HYPHEN) {
      cursor += glyph.length
      continue
    }

    return !isBreakableSpace(glyph)
  }

  return false
}

function isBreakableSpace(glyph: string): boolean {
  return glyph !== NO_BREAK_SPACE && /\s/u.test(glyph)
}

function segmentGraphemes(text: string): ReadonlyArray<string> {
  const segmenterCtor = (Intl as unknown as { Segmenter?: SegmenterConstructor }).Segmenter

  if (!segmenterCtor) {
    return Array.from(text)
  }

  const segmenter = new segmenterCtor(undefined, { granularity: 'grapheme' })
  return Array.from(segmenter.segment(text), (segment) => segment.segment)
}

function containsClusterScript(text: string): boolean {
  for (const glyph of text) {
    const codepoint = glyph.codePointAt(0)
    if (typeof codepoint !== 'number') {
      continue
    }

    if (isCodepointInRanges(codepoint, CLUSTER_SCRIPT_RANGES)) {
      return true
    }
  }

  return false
}

function isCodepointInRanges(
  codepoint: number,
  ranges: ReadonlyArray<readonly [number, number]>,
): boolean {
  for (const [start, end] of ranges) {
    if (codepoint >= start && codepoint <= end) {
      return true
    }
  }

  return false
}

const CLUSTER_SCRIPT_RANGES: ReadonlyArray<readonly [number, number]> = [
  [0x1100, 0x11ff],
  [0x2e80, 0x2eff],
  [0x2f00, 0x2fdf],
  [0x3000, 0x303f],
  [0x3040, 0x309f],
  [0x30a0, 0x30ff],
  [0x3100, 0x312f],
  [0x3130, 0x318f],
  [0x31a0, 0x31bf],
  [0x31c0, 0x31ef],
  [0x31f0, 0x31ff],
  [0x3400, 0x4dbf],
  [0x4e00, 0x9fff],
  [0xac00, 0xd7af],
  [0xf900, 0xfaff],
  [0xff66, 0xff9d],
  [0x20000, 0x2ffff],
]
