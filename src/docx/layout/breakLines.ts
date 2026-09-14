import { measureLineMetricsPt, type FontMetrics } from '../fonts'
import type { Indent } from '../model'

import { effectiveFontFamily } from './fontResolution'
import { itemizeRuns } from './itemize'
import type {
  EffectiveParaProps,
  EffectiveRunProps,
  FontResolver,
  LineBox,
  LineBreakInput,
  LineItem,
  TabStop,
} from './types'

const DEFAULT_FONT_SIZE_PT = 11
const DEFAULT_TAB_STOP_PT = 36

type FontVariant = 'regular' | 'bold' | 'italic' | 'boldItalic'

type RawLine = {
  items: ReadonlyArray<LineItem>
  width: number
  lineLimit: number
}

type BreakOpportunity = {
  splitIndex: number
  penaltyWidth: number
}

export async function breakLines(input: LineBreakInput): Promise<ReadonlyArray<LineBox>> {
  const runItems = await itemizeRuns(input.runs, input.fontResolver, input.theme, input.noteMarks)
  const items =
    input.leadingItems !== undefined && input.leadingItems.length > 0
      ? [...input.leadingItems, ...runItems]
      : runItems
  const rawLines = buildRawLines(resolveTabStopWidths(items, input.tabStops), input)
  const metricsCache = new Map<string, Promise<FontMetrics>>()

  return Promise.all(
    rawLines.map((rawLine, lineIndex) =>
      buildLineBox(rawLine, lineIndex === rawLines.length - 1, input, metricsCache),
    ),
  )
}

function buildRawLines(items: ReadonlyArray<LineItem>, input: LineBreakInput): ReadonlyArray<RawLine> {
  const lines: RawLine[] = []
  let currentItems: LineItem[] = []
  let currentWidth = 0
  let lineIndex = 0
  let lastBreakOpportunity: BreakOpportunity | undefined
  let forceTrailingEmptyLine = items.length === 0
  let itemIndex = 0

  while (itemIndex < items.length) {
    const lineLimit = getLineLimit(input.paraProps, input.availableWidth, lineIndex)
    const item = items[itemIndex]

    if (item.kind === 'break') {
      currentItems.push(item)
      lines.push({
        items: currentItems,
        width: computeLineWidth(currentItems, 0),
        lineLimit,
      })
      currentItems = []
      currentWidth = 0
      lineIndex += 1
      lastBreakOpportunity = undefined
      forceTrailingEmptyLine = true
      itemIndex += 1
      continue
    }

    const itemWidth = getItemWidth(item)
    if (currentItems.length > 0 && itemWidth > 0 && currentWidth + itemWidth > lineLimit) {
      if (lastBreakOpportunity) {
        const lineItems = currentItems.slice(0, lastBreakOpportunity.splitIndex)
        const carryItems = currentItems.slice(lastBreakOpportunity.splitIndex)

        lines.push({
          items: lineItems,
          width: computeLineWidth(lineItems, lastBreakOpportunity.penaltyWidth),
          lineLimit,
        })

        currentItems = carryItems
        currentWidth = sumItemWidths(carryItems)
        lineIndex += 1
        lastBreakOpportunity = findLastBreakOpportunity(carryItems)
        forceTrailingEmptyLine = false
        continue
      }

      lines.push({
        items: currentItems,
        width: computeLineWidth(currentItems, 0),
        lineLimit,
      })
      currentItems = []
      currentWidth = 0
      lineIndex += 1
      lastBreakOpportunity = undefined
      forceTrailingEmptyLine = false
      continue
    }

    currentItems.push(item)
    currentWidth += itemWidth
    const opportunity = getBreakOpportunity(item, currentItems.length)
    if (opportunity) {
      lastBreakOpportunity = opportunity
    }
    forceTrailingEmptyLine = false
    itemIndex += 1
  }

  if (currentItems.length > 0) {
    lines.push({
      items: currentItems,
      width: computeLineWidth(currentItems, 0),
      lineLimit: getLineLimit(input.paraProps, input.availableWidth, lineIndex),
    })
  } else if (lines.length === 0 || forceTrailingEmptyLine) {
    lines.push({
      items: [],
      width: 0,
      lineLimit: getLineLimit(input.paraProps, input.availableWidth, lineIndex),
    })
  }

  return lines
}

async function buildLineBox(
  rawLine: RawLine,
  isLastLine: boolean,
  input: LineBreakInput,
  metricsCache: Map<string, Promise<FontMetrics>>,
): Promise<LineBox> {
  const textMetrics = await computeLineMetrics(rawLine.items, input, metricsCache)
  const { ascent, descent, lineHeight, drawingClearancePt } = applyDrawingClearance(
    textMetrics,
    rawLine.items,
    input.paraProps,
  )
  const spaceCount = countJustificationSpaces(rawLine.items)
  const alignment = input.paraProps.jc
  const canJustify = alignment === 'distribute' || (alignment === 'both' && !isLastLine)
  const justificationStretch =
    canJustify && spaceCount > 0
      ? Math.max(rawLine.lineLimit - rawLine.width, 0) / spaceCount
      : 0

  const lastItem = rawLine.items[rawLine.items.length - 1]
  const endsWithPageBreak = lastItem?.kind === 'break' && lastItem.breakKind === 'page'
  const endsWithColumnBreak = lastItem?.kind === 'break' && lastItem.breakKind === 'column'

  return {
    items: rawLine.items,
    width: rawLine.width,
    ascent,
    descent,
    lineHeight,
    isJustified: canJustify && spaceCount > 0,
    justificationStretch,
    ...(drawingClearancePt > 0 ? { drawingClearancePt } : {}),
    ...(endsWithPageBreak ? { endsWithPageBreak: true } : {}),
    ...(endsWithColumnBreak ? { endsWithColumnBreak: true } : {}),
  }
}

type LineMetrics = { ascent: number; descent: number; lineHeight: number }

/**
 * Drawings sit on the text baseline and extend upward. When one is taller
 * than the space above the baseline, the line grows by the difference so the
 * picture never paints over the previous line. The renderer places the text
 * strut below `drawingClearancePt`, keeping the baseline where layout expects.
 * Exact line spacing is honoured as-is (Word clips pictures in that case).
 */
function applyDrawingClearance(
  metrics: LineMetrics,
  items: ReadonlyArray<LineItem>,
  paraProps: EffectiveParaProps,
): LineMetrics & { drawingClearancePt: number } {
  const tallestDrawing = items.reduce(
    (tallest, item) => (item.kind === 'drawing' ? Math.max(tallest, item.height) : tallest),
    0,
  )

  if (tallestDrawing === 0 || paraProps.spacing?.lineRule === 'exact') {
    return { ...metrics, drawingClearancePt: 0 }
  }

  const halfLeading = (metrics.lineHeight - metrics.ascent - metrics.descent) / 2
  const baselineFromTop = halfLeading + metrics.ascent
  const drawingClearancePt = Math.max(0, tallestDrawing - baselineFromTop)

  return {
    ascent: Math.max(metrics.ascent, tallestDrawing),
    descent: metrics.descent,
    lineHeight: metrics.lineHeight + drawingClearancePt,
    drawingClearancePt,
  }
}

async function computeLineMetrics(
  items: ReadonlyArray<LineItem>,
  input: LineBreakInput,
  metricsCache: Map<string, Promise<FontMetrics>>,
): Promise<{ ascent: number; descent: number; lineHeight: number }> {
  const runIndices = new Set<number>()
  const sampleTextByRun = new Map<number, string>()

  for (const item of items) {
    runIndices.add(item.runIndex)
    if ((item.kind === 'word' || item.kind === 'glyph-cluster') && !sampleTextByRun.has(item.runIndex)) {
      sampleTextByRun.set(item.runIndex, item.text)
    }
  }

  if (runIndices.size === 0 && input.runs.length > 0) {
    runIndices.add(0)
  }

  let ascent = 0
  let descent = 0
  let lineGap = 0
  let browserHeight = 0

  if (runIndices.size === 0) {
    const fallback = await resolveRunFontMetrics({}, '', input.fontResolver, metricsCache, input.theme)
    ascent = scaleMetric(fallback.metrics.ascender, fallback.metrics.unitsPerEm, fallback.sizePt)
    descent = Math.abs(scaleMetric(fallback.metrics.descender, fallback.metrics.unitsPerEm, fallback.sizePt))
    lineGap = scaleMetric(fallback.metrics.lineGap, fallback.metrics.unitsPerEm, fallback.sizePt)
    const browser = measureLineMetricsPt(fallback.family, fallback.variant, fallback.sizePt)
    if (browser !== null) {
      ascent = Math.max(ascent, browser.ascentPt)
      descent = Math.max(descent, browser.descentPt)
      browserHeight = Math.max(browserHeight, browser.heightPt)
    }
  } else {
    for (const runIndex of runIndices) {
      const runProps = input.runs[runIndex]?.runProps ?? {}
      const sampleText = sampleTextByRun.get(runIndex) ?? ''
      const resolved = await resolveRunFontMetrics(
        runProps,
        sampleText,
        input.fontResolver,
        metricsCache,
        input.theme,
      )

      ascent = Math.max(
        ascent,
        scaleMetric(resolved.metrics.ascender, resolved.metrics.unitsPerEm, resolved.sizePt),
      )
      descent = Math.max(
        descent,
        Math.abs(scaleMetric(resolved.metrics.descender, resolved.metrics.unitsPerEm, resolved.sizePt)),
      )
      lineGap = Math.max(
        lineGap,
        scaleMetric(resolved.metrics.lineGap, resolved.metrics.unitsPerEm, resolved.sizePt),
      )

      const browser = measureLineMetricsPt(resolved.family, resolved.variant, resolved.sizePt)
      if (browser !== null) {
        ascent = Math.max(ascent, browser.ascentPt)
        descent = Math.max(descent, browser.descentPt)
        browserHeight = Math.max(browserHeight, browser.heightPt)
      }
    }
  }

  // Prefer browser-derived line-box height when available (matches what
  // Chromium actually paints for inline-block word spans); fall back to
  // TTF asc+desc+lineGap with a 1.2 floor in Node tests.
  const naturalLineHeight =
    browserHeight > 0
      ? browserHeight
      : Math.max(ascent + descent + lineGap, ascent * 1.2)

  return {
    ascent,
    descent,
    lineHeight: resolveLineHeight(input.paraProps, naturalLineHeight),
  }
}

async function resolveRunFontMetrics(
  runProps: EffectiveRunProps,
  sampleText: string,
  fontResolver: FontResolver,
  metricsCache: Map<string, Promise<FontMetrics>>,
  theme: LineBreakInput['theme'],
): Promise<{ metrics: FontMetrics; sizePt: number; family: string; variant: FontVariant }> {
  const family = resolveFontFamily(runProps, sampleText, theme)
  const variant = resolveFontVariant(runProps)
  const cacheKey = `${family}::${variant}`
  let metricsPromise = metricsCache.get(cacheKey)

  if (!metricsPromise) {
    metricsPromise = fontResolver(family, variant)
    metricsCache.set(cacheKey, metricsPromise)
  }

  return {
    metrics: await metricsPromise,
    sizePt: resolveFontSize(runProps),
    family,
    variant,
  }
}

/**
 * Resolves every tab item's width (and leader glyph) in one forward pass
 * over the WHOLE item list, before line-breaking runs (D24/DXL-16). Center/
 * right/decimal tab stops need to know how wide the content following the
 * tab is — Word positions that content so its center/end/decimal-point
 * lands exactly at the stop, rather than starting the content flush at the
 * stop the way a left tab does — which means looking ahead past the tab to
 * the next hard stop (another tab, an explicit line break, or the end of
 * the paragraph).
 *
 * This look-ahead deliberately does NOT know where an ordinary (unforced)
 * line wrap will land — that's exactly what line-breaking, running right
 * after this pass, still has to figure out — so it assumes the tab and its
 * following content stay on one line. That holds for the dominant
 * real-world use of non-left tab stops (a single-line TOC/header/footer
 * entry like "Chapter 1........42" or a right-aligned page number); a
 * center/right/decimal tab stop whose content is long enough to wrap onto a
 * second line is a rare combination where this pass's width may end up
 * very slightly off (an accepted, documented approximation).
 */
function resolveTabStopWidths(
  items: ReadonlyArray<LineItem>,
  tabStops: ReadonlyArray<TabStop>,
): ReadonlyArray<LineItem> {
  let currentWidth = 0
  let sawTab = false

  const resolved = items.map((item, index) => {
    if (item.kind === 'break') {
      currentWidth = 0
      return item
    }

    if (item.kind !== 'tab') {
      currentWidth += getItemWidth(item)
      return item
    }

    sawTab = true
    const matchedStop = tabStops.find((tabStop) => tabStop.positionPt > currentWidth)
    const width = resolveTabStopWidth(currentWidth, matchedStop, () => lookaheadTabContent(items, index + 1))
    currentWidth += width
    return { ...item, width, leader: matchedStop?.leader ?? 'none' }
  })

  return sawTab ? resolved : items
}

type TabContentLookahead = {
  readonly totalWidthPt: number
  /** Width of the content up to (not including) its first `.` character, or `undefined` when none is found before the next hard stop. */
  readonly preDecimalWidthPt: number | undefined
}

function lookaheadTabContent(items: ReadonlyArray<LineItem>, startIndex: number): TabContentLookahead {
  let totalWidthPt = 0
  let preDecimalWidthPt: number | undefined

  for (let index = startIndex; index < items.length; index += 1) {
    const item = items[index]
    if (item.kind === 'tab' || item.kind === 'break') {
      break
    }

    const width = getItemWidth(item)
    if (preDecimalWidthPt === undefined && (item.kind === 'word' || item.kind === 'glyph-cluster')) {
      const dotIndex = item.text.indexOf('.')
      if (dotIndex !== -1) {
        const fraction = item.text.length > 0 ? dotIndex / item.text.length : 0
        preDecimalWidthPt = totalWidthPt + width * fraction
      }
    }

    totalWidthPt += width
  }

  return { totalWidthPt, preDecimalWidthPt }
}

function resolveTabStopWidth(
  currentWidth: number,
  matchedStop: TabStop | undefined,
  lookahead: () => TabContentLookahead,
): number {
  if (matchedStop === undefined) {
    const nextDefaultStop = (Math.floor(currentWidth / DEFAULT_TAB_STOP_PT) + 1) * DEFAULT_TAB_STOP_PT
    return Math.max(0, nextDefaultStop - currentWidth)
  }

  if (matchedStop.alignment === 'left') {
    return Math.max(0, matchedStop.positionPt - currentWidth)
  }

  const content = lookahead()

  if (matchedStop.alignment === 'center') {
    return Math.max(0, matchedStop.positionPt - content.totalWidthPt / 2 - currentWidth)
  }

  if (matchedStop.alignment === 'right') {
    return Math.max(0, matchedStop.positionPt - content.totalWidthPt - currentWidth)
  }

  // decimal — falls back to right-alignment (whole content ends at the
  // stop) when no decimal point precedes the next hard stop, matching
  // Word's own documented fallback for a decimal tab with no decimal.
  const anchorWidthPt = content.preDecimalWidthPt ?? content.totalWidthPt
  return Math.max(0, matchedStop.positionPt - anchorWidthPt - currentWidth)
}

function getBreakOpportunity(item: LineItem, splitIndex: number): BreakOpportunity | undefined {
  if (item.kind === 'space' || item.kind === 'tab' || item.kind === 'glyph-cluster') {
    return { splitIndex, penaltyWidth: 0 }
  }

  if (item.kind === 'hyphen-opportunity') {
    return { splitIndex, penaltyWidth: item.penaltyWidth }
  }

  return undefined
}

function findLastBreakOpportunity(items: ReadonlyArray<LineItem>): BreakOpportunity | undefined {
  let result: BreakOpportunity | undefined

  for (const [index, item] of items.entries()) {
    const opportunity = getBreakOpportunity(item, index + 1)
    if (opportunity) {
      result = opportunity
    }
  }

  return result
}

function getLineLimit(
  paraProps: EffectiveParaProps,
  availableWidth: number,
  lineIndex: number,
): number {
  const leftIndent = resolveLeftIndentPt(paraProps.ind)
  const lineIndent = resolveLineIndentExtraPt(paraProps.ind, lineIndex)

  return Math.max(0, availableWidth - leftIndent - lineIndent)
}

/** The paragraph's base left indent (`w:ind/@w:left`, or `@w:start`), in points. */
export function resolveLeftIndentPt(ind: Indent | undefined): number {
  return twipToPt(ind?.left ?? ind?.start)
}

/**
 * Extra horizontal offset for one specific line of a paragraph, beyond its
 * base left indent (`resolveLeftIndentPt`):
 *
 * - `firstLine` PUSHES the first line right (a positive offset), leaving
 *   continuation lines at the base indent — the classic prose first-line
 *   indent.
 * - `hanging` PULLS the first line left (a negative offset) instead, so
 *   continuation lines sit at the base indent while the first line (where a
 *   list marker lives, see D3) hangs out past it — the classic
 *   marker-plus-hanging-indent list layout.
 *
 * `firstLine` and `hanging` are mutually exclusive per the OOXML schema;
 * when a (malformed) paragraph sets both, `firstLine` wins. Used both to
 * size each line's available width (here) and, in `paginate.ts`'s
 * `placeLine`, to position it horizontally — the two MUST stay in sync so a
 * line never renders wider than the space it was measured against.
 */
export function resolveLineIndentExtraPt(ind: Indent | undefined, lineIndex: number): number {
  if (lineIndex !== 0) {
    return 0
  }

  if (typeof ind?.firstLine === 'number') {
    return twipToPt(ind.firstLine)
  }

  if (typeof ind?.hanging === 'number') {
    return -twipToPt(ind.hanging)
  }

  return 0
}

function computeLineWidth(items: ReadonlyArray<LineItem>, penaltyWidth: number): number {
  let width = sumItemWidths(items) + penaltyWidth
  let cursor = items.length - 1

  while (cursor >= 0 && items[cursor].kind === 'break') {
    cursor -= 1
  }

  while (cursor >= 0) {
    const trailing = items[cursor]
    if (trailing.kind !== 'space') {
      break
    }
    width -= trailing.width
    cursor -= 1
  }

  return Math.max(0, width)
}

function sumItemWidths(items: ReadonlyArray<LineItem>): number {
  let width = 0
  for (const item of items) {
    width += getItemWidth(item)
  }
  return width
}

function getItemWidth(item: LineItem): number {
  if (
    item.kind === 'word' ||
    item.kind === 'space' ||
    item.kind === 'tab' ||
    item.kind === 'glyph-cluster' ||
    item.kind === 'drawing'
  ) {
    return item.width
  }

  return 0
}

function countJustificationSpaces(items: ReadonlyArray<LineItem>): number {
  const lastVisibleIndex = findLastVisibleIndex(items)
  if (lastVisibleIndex < 0) {
    return 0
  }

  let count = 0
  for (let index = 0; index <= lastVisibleIndex; index += 1) {
    const item = items[index]
    if (item.kind === 'space' && item.stretchable) {
      count += 1
    }
  }

  return count
}

function findLastVisibleIndex(items: ReadonlyArray<LineItem>): number {
  let cursor = items.length - 1

  while (cursor >= 0 && items[cursor].kind === 'break') {
    cursor -= 1
  }

  while (cursor >= 0 && items[cursor].kind === 'space') {
    cursor -= 1
  }

  return cursor
}

/**
 * Indices of the `space` items within a justified line that should absorb
 * `line.justificationStretch` when rendering (D5) — the exact same set
 * `countJustificationSpaces` counted when computing that stretch amount in
 * the first place, so the renderer must reuse this rather than re-deriving
 * its own notion of "which spaces stretch": trailing whitespace after the
 * last visible glyph is excluded both times, or a render pass that
 * stretched it too would overshoot `lineLimit`.
 */
export function resolveStretchableSpaceIndices(items: ReadonlyArray<LineItem>): ReadonlySet<number> {
  const lastVisibleIndex = findLastVisibleIndex(items)
  if (lastVisibleIndex < 0) {
    return new Set()
  }

  const indices = new Set<number>()
  for (let index = 0; index <= lastVisibleIndex; index += 1) {
    const item = items[index]
    if (item.kind === 'space' && item.stretchable) {
      indices.add(index)
    }
  }

  return indices
}

function resolveLineHeight(paraProps: EffectiveParaProps, naturalLineHeight: number): number {
  const spacing = paraProps.spacing
  const lineRule = spacing?.lineRule ?? 'auto'

  if (lineRule === 'exact') {
    return typeof spacing?.line === 'number' ? spacing.line / 20 : naturalLineHeight
  }

  if (lineRule === 'atLeast') {
    return typeof spacing?.line === 'number'
      ? Math.max(naturalLineHeight, spacing.line / 20)
      : naturalLineHeight
  }

  const multiplier = typeof spacing?.line === 'number' ? spacing.line / 240 : 1
  return naturalLineHeight * multiplier
}

function scaleMetric(metric: number, unitsPerEm: number, sizePt: number): number {
  return (metric / unitsPerEm) * sizePt
}

function twipToPt(value: number | undefined): number {
  return typeof value === 'number' ? value / 20 : 0
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
  sampleText: string,
  theme: LineBreakInput['theme'],
): string {
  const hint = containsClusterScript(sampleText) ? 'eastAsian' : runProps.rtl ? 'rtl' : 'latin'
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
