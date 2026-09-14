/**
 * Atlas — PAGE/NUMPAGES/SECTIONPAGES field resolution (D11, header/footer fields)
 *
 * A header/footer's "Page X of Y" is an OOXML *field*, authored as either:
 *
 * - a simple field: `<w:fldSimple w:instr="PAGE"> ... cached result ... </w:fldSimple>`
 * - a complex field: a run-of-siblings sequence —
 *   `<w:fldChar w:fldCharType="begin"/>`, one or more
 *   `<w:instrText>PAGE</w:instrText>`, `<w:fldChar w:fldCharType="separate"/>`,
 *   the cached display text, `<w:fldChar w:fldCharType="end"/>` — which is
 *   what Word itself always generates (confirmed against this wave's own
 *   `header-footer-page-numbers` corpus fixture).
 *
 * Neither shape is parsed into a real model node yet — `src/docx/parser/
 * document.ts` currently falls through to `UnknownNode` for `w:fldChar`/
 * `w:instrText`/`w:fldSimple`, so an instruction string never reaches this
 * layer today. That parser-level field model (and the run-of-UnknownNodes
 * -> single Field node collapsing it needs) is being built on a sibling
 * branch (docx-fields-fonts) alongside a richer field-instruction feature
 * set — deliberately NOT duplicated here to avoid two branches racing to
 * parse the same OOXML shape differently.
 *
 * This module owns the one piece of that problem that belongs to layout
 * rather than parsing: turning an already-extracted instruction string
 * (`"PAGE"`, `"NUMPAGES \* ROMAN"`, `"SECTIONPAGES"`, ...) into resolved
 * display text using pagination's own page-number context — a pure,
 * self-contained function the field-parsing branch can call once its model
 * hands `paginate.ts`/`PageView.tsx` a real instruction string, without
 * needing to know anything about numbering-format switch syntax itself.
 */

/** The subset of `ST_NumberFormat` switch values (`\* SWITCH`) this resolver recognizes; anything else (MERGEFORMAT, CHARFORMAT, CAPS, ...) falls back to plain decimal, matching Word's own behavior for a switch that doesn't affect the numeric text itself. */
export type PageFieldNumberFormat = 'decimal' | 'upperRoman' | 'lowerRoman' | 'upperLetter' | 'lowerLetter'

export type PageFieldContext = {
  /** This page's 1-based number within the WHOLE document — what a bare `PAGE` field resolves to. */
  readonly pageNumber: number
  /** The document's total page count — what `NUMPAGES` resolves to. */
  readonly totalPages: number
  /** This page's 1-based number within its own section, when known — a `PAGE` field with the `\s` switch (page numbering restarted per section) would use this instead; omitted contexts fall back to `pageNumber`. */
  readonly sectionPageNumber?: number
  /** The owning section's own page count, when known — what `SECTIONPAGES` resolves to; omitted contexts fall back to `totalPages`. */
  readonly sectionTotalPages?: number
}

const KEYWORD_PATTERN = /^\s*([A-Za-z]+)/
const NUMERIC_SWITCH_PATTERN = /\\\*\s*(\S+)/

/**
 * Resolves one field instruction to its display text, or `undefined` when
 * the instruction isn't one of the three page-count fields this resolver
 * handles (a caller should fall back to whatever cached result the source
 * document's `w:fldChar`/`w:fldSimple` "separate" content already carried,
 * exactly as Word itself does for a field type it can't recompute).
 *
 * Only the numeric-format switch (`\* ARABIC`/`\* ROMAN`/`\* roman`/
 * `\* ALPHABETIC`/`\* alphabetic`) is honored; every other recognized OOXML
 * field switch (`\# 0.00` custom pictures, `\p` position switches, locking
 * switches like `\* MERGEFORMAT`) is accepted in the instruction text
 * without erroring but has no effect on the returned digits/letters — a
 * deliberately narrow first cut matching this task's scope (see the module
 * doc comment).
 */
export function resolvePageField(instruction: string, ctx: PageFieldContext): string | undefined {
  const keyword = extractFieldKeyword(instruction)
  if (keyword === undefined) {
    return undefined
  }

  const value = resolveFieldValue(keyword, ctx)
  if (value === undefined) {
    return undefined
  }

  return formatPageFieldNumber(value, extractNumericFormatSwitch(instruction))
}

function resolveFieldValue(keyword: string, ctx: PageFieldContext): number | undefined {
  switch (keyword) {
    case 'PAGE':
      return ctx.pageNumber
    case 'NUMPAGES':
      return ctx.totalPages
    case 'SECTIONPAGES':
      return ctx.sectionTotalPages ?? ctx.totalPages
    default:
      return undefined
  }
}

function extractFieldKeyword(instruction: string): string | undefined {
  return KEYWORD_PATTERN.exec(instruction)?.[1]?.toUpperCase()
}

function extractNumericFormatSwitch(instruction: string): PageFieldNumberFormat {
  const token = NUMERIC_SWITCH_PATTERN.exec(instruction)?.[1]

  switch (token) {
    case 'ROMAN':
      return 'upperRoman'
    case 'roman':
      return 'lowerRoman'
    case 'ALPHABETIC':
      return 'upperLetter'
    case 'alphabetic':
      return 'lowerLetter'
    default:
      return 'decimal'
  }
}

function formatPageFieldNumber(value: number, format: PageFieldNumberFormat): string {
  switch (format) {
    case 'upperRoman':
      return toRoman(value).toUpperCase()
    case 'lowerRoman':
      return toRoman(value).toLowerCase()
    case 'upperLetter':
      return toAlpha(value).toUpperCase()
    case 'lowerLetter':
      return toAlpha(value).toLowerCase()
    case 'decimal':
    default:
      return String(value)
  }
}

// Duplicated (rather than imported) from `noteNumbering.ts`'s identical
// helpers: the two modules resolve conceptually unrelated OOXML numbering
// grammars (note-numbering formats vs. field numeric-format switches) that
// only coincidentally share the same roman/letter algorithms, and each is
// small enough that a shared dependency would cost more in coupling than it
// saves in lines.
function toRoman(value: number): string {
  if (value <= 0) {
    return String(value)
  }

  const romanValues: ReadonlyArray<readonly [number, string]> = [
    [1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'],
    [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'],
    [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i'],
  ]

  let remaining = value
  let result = ''
  for (const [amount, numeral] of romanValues) {
    while (remaining >= amount) {
      result += numeral
      remaining -= amount
    }
  }
  return result
}

function toAlpha(value: number): string {
  if (value <= 0) {
    return String(value)
  }

  const letterIndex = (value - 1) % 26
  const repeatCount = Math.floor((value - 1) / 26) + 1
  return String.fromCharCode(97 + letterIndex).repeat(repeatCount)
}
