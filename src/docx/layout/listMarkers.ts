/**
 * Atlas — DOCX list numbering marker generation (D3 / DXP-07, DXL-04)
 *
 * `numbering.ts` already parses `word/numbering.xml` correctly (abstract
 * numbering definitions, per-level formats/text templates, `lvlOverride`),
 * but nothing in the layout pipeline ever turned that into an actual marker
 * — bulleted and numbered paragraphs rendered with no marker at all. This
 * module owns that: per-(numId, ilvl) counter state that increments as
 * paragraphs are laid out in document order, formatting a counter value per
 * the level's `w:numFmt`, and substituting a level's `%N` placeholders in
 * its `lvlText` template.
 */

import type { Document, LvlDef, NumberingDef, NumPr, RunProps } from '../model'

/**
 * Sentinel `LineItem.runIndex` for a list marker's itemized text/tab —
 * chosen to never collide with (or shift) a real 0-based run index, so
 * marker items can be measured through the same itemization pipeline as
 * real content without disturbing revision-tracking or hyperlink lookups
 * keyed by run index.
 */
export const MARKER_RUN_INDEX = -1

// ---------------------------------------------------------------------------
// Counter state
// ---------------------------------------------------------------------------

export type NumberingCounterState = {
  /** numId -> ilvl -> current (already-incremented) counter value. */
  readonly counters: Map<string, Map<number, number>>
}

export function createNumberingCounterState(): NumberingCounterState {
  return { counters: new Map() }
}

export type ListMarker = {
  readonly text: string
  readonly runProps: RunProps | undefined
  /** What separates the marker from the paragraph's own text. */
  readonly suffix: 'tab' | 'space' | 'nothing'
}

/**
 * Advances the (numId, ilvl) counter for one list paragraph and returns its
 * rendered marker, or `undefined` when the paragraph carries no numbering
 * reference, the numbering definition/level can't be found, or the level's
 * format is explicitly `none` (no visible marker, e.g. an outline level used
 * purely for indentation).
 *
 * Must be called exactly once per list paragraph, in document order — the
 * counters are stateful and mutated on every call.
 */
export function resolveListMarker(
  numPr: NumPr | undefined,
  document: Document,
  state: NumberingCounterState,
): ListMarker | undefined {
  if (numPr?.numId === undefined) {
    return undefined
  }

  const def = document.numbering.get(numPr.numId)
  if (def === undefined) {
    return undefined
  }

  const ilvl = numPr.ilvl ?? 0
  const override = def.levelOverrides?.get(ilvl)
  const levelDef = override?.levelDefinition ?? def.levels.get(ilvl)
  if (levelDef === undefined || levelDef.format === 'none') {
    return undefined
  }

  const value = advanceCounter(state, numPr.numId, ilvl, levelDef, override?.startOverride, def)

  const text =
    levelDef.format === 'bullet'
      ? mapSymbolGlyph(levelDef.text?.value ?? DEFAULT_BULLET_GLYPH, levelDef.run)
      : renderLevelText(state, numPr.numId, def, ilvl, levelDef, value)

  return {
    text,
    runProps: levelDef.run,
    suffix: levelDef.suffix ?? 'tab',
  }
}

/**
 * Advances and returns the counter for `ilvl`, resetting every deeper level
 * (per its own `w:lvlRestart`, defaulting to "restart when its immediate
 * parent level is used") so it starts fresh from its own `start`/
 * `startOverride` value the next time it's used. Word's default behavior —
 * absent an explicit `lvlRestart` — is for a level to restart whenever any
 * shallower (numerically lower `ilvl`) level increments.
 */
function advanceCounter(
  state: NumberingCounterState,
  numId: string,
  ilvl: number,
  levelDef: LvlDef,
  startOverride: number | undefined,
  def: NumberingDef,
): number {
  let numCounters = state.counters.get(numId)
  if (numCounters === undefined) {
    numCounters = new Map()
    state.counters.set(numId, numCounters)
  }

  for (const deeperLevel of [...numCounters.keys()]) {
    if (deeperLevel <= ilvl) {
      continue
    }
    const deeperLevelDef =
      def.levelOverrides?.get(deeperLevel)?.levelDefinition ?? def.levels.get(deeperLevel)
    const restartTrigger = deeperLevelDef?.restart ?? deeperLevel - 1
    if (restartTrigger >= ilvl) {
      numCounters.delete(deeperLevel)
    }
  }

  const startValue = startOverride ?? levelDef.start ?? 1
  const current = numCounters.get(ilvl)
  const nextValue = current === undefined ? startValue : current + 1
  numCounters.set(ilvl, nextValue)
  return nextValue
}

/**
 * Substitutes every `%N` placeholder in a level's `lvlText` template (`N` is
 * 1-based; `%1` refers to the outermost level, `%N` to `ilvl === N - 1`)
 * with that ancestor level's CURRENT counter value, formatted per that
 * ancestor's own `numFmt` — this is how a template like "%1.%2." renders a
 * third item under the second first-level item as "2.3.". An ancestor level
 * that hasn't been reached yet in document order (no counter recorded)
 * falls back to its configured `start` (or 1).
 */
function renderLevelText(
  state: NumberingCounterState,
  numId: string,
  def: NumberingDef,
  ilvl: number,
  levelDef: LvlDef,
  currentValue: number,
): string {
  const template = levelDef.text?.value
  if (template === undefined || template === '') {
    return formatCounterValue(currentValue, levelDef.format)
  }

  return template.replace(/%(\d)/g, (match, placeholderDigit: string) => {
    const placeholderLevel = Number.parseInt(placeholderDigit, 10) - 1
    if (Number.isNaN(placeholderLevel)) {
      return match
    }

    if (placeholderLevel === ilvl) {
      return formatCounterValue(currentValue, levelDef.format)
    }

    const ancestorLevelDef =
      def.levelOverrides?.get(placeholderLevel)?.levelDefinition ?? def.levels.get(placeholderLevel)
    const ancestorValue =
      state.counters.get(numId)?.get(placeholderLevel) ?? ancestorLevelDef?.start ?? 1

    return formatCounterValue(ancestorValue, ancestorLevelDef?.format)
  })
}

// ---------------------------------------------------------------------------
// Counter value formatting
// ---------------------------------------------------------------------------

function formatCounterValue(value: number, format: string | undefined): string {
  switch (format) {
    case 'decimalZero':
      return value < 10 ? `0${value}` : String(value)
    case 'upperRoman':
      return toRoman(value).toUpperCase()
    case 'lowerRoman':
      return toRoman(value).toLowerCase()
    case 'upperLetter':
      return toAlpha(value).toUpperCase()
    case 'lowerLetter':
      return toAlpha(value).toLowerCase()
    case 'ordinal':
      return `${value}${ordinalSuffix(value)}`
    case 'none':
      return ''
    case 'decimal':
    default:
      return String(value)
  }
}

const ROMAN_VALUES: ReadonlyArray<readonly [number, string]> = [
  [1000, 'm'],
  [900, 'cm'],
  [500, 'd'],
  [400, 'cd'],
  [100, 'c'],
  [90, 'xc'],
  [50, 'l'],
  [40, 'xl'],
  [10, 'x'],
  [9, 'ix'],
  [5, 'v'],
  [4, 'iv'],
  [1, 'i'],
]

function toRoman(value: number): string {
  if (value <= 0) {
    return String(value)
  }

  let remaining = value
  let result = ''
  for (const [amount, numeral] of ROMAN_VALUES) {
    while (remaining >= amount) {
      result += numeral
      remaining -= amount
    }
  }
  return result
}

/**
 * Word's letter numbering repeats the alphabet in whole blocks after `z`
 * rather than incrementing digit-by-digit like a spreadsheet column: ...,
 * y, z, aa, bb, cc, ..., zz, aaa, ....
 */
function toAlpha(value: number): string {
  if (value <= 0) {
    return String(value)
  }

  const letterIndex = (value - 1) % 26
  const repeatCount = Math.floor((value - 1) / 26) + 1
  return String.fromCharCode(97 + letterIndex).repeat(repeatCount)
}

function ordinalSuffix(value: number): string {
  const remainder100 = value % 100
  if (remainder100 >= 11 && remainder100 <= 13) {
    return 'th'
  }
  switch (value % 10) {
    case 1:
      return 'st'
    case 2:
      return 'nd'
    case 3:
      return 'rd'
    default:
      return 'th'
  }
}

// ---------------------------------------------------------------------------
// Symbol-font glyph mapping
// ---------------------------------------------------------------------------

const DEFAULT_BULLET_GLYPH = '•'

const SYMBOL_FONT_NAMES: ReadonlySet<string> = new Set([
  'wingdings',
  'wingdings 2',
  'wingdings 3',
  'symbol',
  'webdings',
])

/**
 * Word bullets are frequently authored as a single private-use-area
 * codepoint in a symbol font (e.g. Wingdings 0xF0B7 for a round bullet)
 * rather than a real Unicode bullet character — rendered as a box/tofu
 * glyph without that exact font installed. Map the handful of codepoints
 * that account for the vast majority of real-world bulleted lists to their
 * plain-Unicode equivalents; anything unrecognized (or not in a known
 * symbol font) passes through unchanged.
 */
const SYMBOL_GLYPH_MAP: ReadonlyMap<string, string> = new Map([
  ['', '•'], // round bullet -> •
  ['', '▪'], // small square -> ▪
  ['', '▫'], // small square outline -> ▫
  ['', '➔'], // arrow -> ➔
  ['', '✓'], // check mark -> ✓
])

function mapSymbolGlyph(text: string, run: RunProps | undefined): string {
  const fontName = (run?.rFonts?.ascii ?? run?.rFonts?.hAnsi ?? '').trim().toLowerCase()
  if (!SYMBOL_FONT_NAMES.has(fontName)) {
    return text
  }

  let mapped = ''
  for (const glyph of text) {
    mapped += SYMBOL_GLYPH_MAP.get(glyph) ?? glyph
  }
  return mapped
}
