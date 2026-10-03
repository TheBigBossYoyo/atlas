/**
 * How this application orders and groups cell VALUES — one definition, shared.
 *
 * Sorting (`sortRows`), the column filter's value list (`columnFilter.ts`) and
 * the formula evaluator's comparison operators (`compareValues` in
 * `spreadsheetFormula.ts`) must agree about three questions, because a user
 * sees all three in the same sheet:
 *
 *   - is this cell a number or text?  (numbers order before text)
 *   - is "Paper" the same value as "paper"?  (yes — text is case-insensitive)
 *   - where does an empty cell go?  (last, in BOTH directions)
 *
 * Keeping the answers in one module is not tidiness: two copies drift, and the
 * drift shows up as a sort whose order contradicts a `>` in a formula, or a
 * filter list that offers "Paper" and "paper" as separate entries while the
 * sort treats them as one.
 */

/**
 * The case fold used for every text comparison and grouping.
 *
 * `toUpperCase` rather than `toLowerCase`, matching the formula evaluator — the
 * two must fold identically or the agreement above is lost. (They differ for a
 * handful of scripts; German ß upper-cases to SS, so `ß` and `ss` group
 * together, which is what a German user comparing text in Excel also gets.)
 */
export function foldText(text: string): string {
  return text.toUpperCase()
}

/** Rank 0 sorts first, then 1, then 2 — numbers, then text, then blanks. */
export type CellRank = 0 | 1 | 2

export type CellSortKey = {
  readonly rank: CellRank
  readonly num: number
  readonly str: string
}

/**
 * Orders a cell's value.
 *
 * Numbers before text, blanks last, text compared case-insensitively — the same
 * three rules the formula evaluator's comparison operators use, so a sort and a
 * `>` in a formula cannot disagree about which of two cells is larger. Blanks
 * last in BOTH directions, matching Excel: an empty cell is absence, not a
 * small value, and burying the data under a block of blanks on a descending
 * sort would be useless.
 */
export function sortKey(text: string): CellSortKey {
  const trimmed = text.trim()
  if (trimmed === '') return { rank: 2, num: 0, str: '' }
  const parsed = Number(trimmed)
  if (!Number.isNaN(parsed)) return { rank: 0, num: parsed, str: '' }
  return { rank: 1, num: 0, str: foldText(trimmed) }
}

/**
 * Compares two cells, in `direction`.
 *
 * The direction applies only WITHIN a rank, never to the rank itself: a blank
 * stays last whichever way the sort runs. Signing the whole comparison — which
 * the first version did — puts the blanks at the top of a descending sort and
 * buries the data under them.
 */
export function compareForSort(a: string, b: string, direction: 'asc' | 'desc'): number {
  const ka = sortKey(a)
  const kb = sortKey(b)
  if (ka.rank !== kb.rank) return ka.rank - kb.rank
  const sign = direction === 'asc' ? 1 : -1
  if (ka.rank === 0) return (ka.num === kb.num ? 0 : ka.num < kb.num ? -1 : 1) * sign
  if (ka.rank === 1) return (ka.str === kb.str ? 0 : ka.str < kb.str ? -1 : 1) * sign
  return 0
}
