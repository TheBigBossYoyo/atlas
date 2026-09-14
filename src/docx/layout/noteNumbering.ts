/**
 * Atlas — footnote/endnote reference numbering (D11 milestones 2 + 5)
 *
 * Two independent concerns live here:
 *
 * 1. `formatSequenceNumber` — renders a 1-based counter value per a
 *    footnote/endnote's `w:numFmt` (the same closed set of OOXML numbering
 *    formats `listMarkers.ts` already implements for list markers; kept as
 *    an independent, smaller copy here rather than importing from
 *    `listMarkers.ts` so this module has no dependency on list-numbering
 *    internals it doesn't otherwise need).
 * 2. `NoteNumberingState` — a mutable, document-order counter (mirroring
 *    `listMarkers.ts`'s `NumberingCounterState`) that `paginate.ts` advances
 *    exactly once per footnote/endnote reference as it itemizes paragraphs
 *    in document order. This gives correct numbering for the two most
 *    common restart rules directly:
 *      - `continuous` (the default): never reset, handled by simply never
 *        calling `resetFootnoteCounter`/`resetEndnoteCounter`.
 *      - `eachSect`: `paginate.ts` calls the reset at the start of every
 *        section, before that section's paragraphs are itemized.
 *    `eachPage` cannot be resolved at itemization time — which page a
 *    reference lands on is only known after line-placement — so
 *    `paginate.ts` assigns *placeholder* continuous numbers here (used to
 *    measure/render the superscript mark during layout) and separately
 *    renumbers the *displayed* text per page in a post-layout pass (see
 *    `renumberFootnotesEachPage` below). That pass only ever changes label
 *    text, not layout, so a page with 10+ footnotes could in principle show
 *    a marker measured at one digit's width post-renumbered to two digits —
 *    an accepted, documented approximation (DXL-09 sub-milestone 5).
 */
import type { EndnoteReference, FootnoteReference, HyperlinkChild, ParagraphChild } from '../model'

export type NoteReference = {
  readonly kind: 'footnote' | 'endnote'
  readonly id: string
}

export type NoteNumberingState = {
  footnoteCounter: number
  endnoteCounter: number
  readonly footnoteMarkById: Map<string, string>
  readonly endnoteMarkById: Map<string, string>
  /** Ids in first-referenced (document) order — feeds endnote layout and the `eachPage` renumbering pass. */
  readonly footnoteOrder: string[]
  readonly endnoteOrder: string[]
}

export function createNoteNumberingState(): NoteNumberingState {
  return {
    footnoteCounter: 0,
    endnoteCounter: 0,
    footnoteMarkById: new Map(),
    endnoteMarkById: new Map(),
    footnoteOrder: [],
    endnoteOrder: [],
  }
}

/** Resets the footnote counter back to `start - 1` (next assignment yields `start`) — used for `eachSect`/`eachPage` restart. */
export function resetFootnoteCounter(state: NoteNumberingState, start: number): void {
  state.footnoteCounter = start - 1
}

export function resetEndnoteCounter(state: NoteNumberingState, start: number): void {
  state.endnoteCounter = start - 1
}

/**
 * Assigns (or returns the already-assigned) display mark for one footnote
 * reference. Idempotent per id — a footnote is normally referenced exactly
 * once, but a malformed document referencing the same id twice reuses the
 * first assignment rather than advancing the counter again.
 */
export function assignFootnoteMark(state: NoteNumberingState, id: string, numFmt: string | undefined): string {
  const existing = state.footnoteMarkById.get(id)
  if (existing !== undefined) {
    return existing
  }

  state.footnoteCounter += 1
  const mark = formatSequenceNumber(state.footnoteCounter, numFmt)
  state.footnoteMarkById.set(id, mark)
  state.footnoteOrder.push(id)
  return mark
}

export function assignEndnoteMark(state: NoteNumberingState, id: string, numFmt: string | undefined): string {
  const existing = state.endnoteMarkById.get(id)
  if (existing !== undefined) {
    return existing
  }

  state.endnoteCounter += 1
  const mark = formatSequenceNumber(state.endnoteCounter, numFmt)
  state.endnoteMarkById.set(id, mark)
  state.endnoteOrder.push(id)
  return mark
}

/**
 * Post-layout `eachPage` renumbering: given the per-page ordering of
 * referenced footnote ids (first-appearance order within each page, as
 * `paginate.ts`'s placement pass discovers them), returns a fresh
 * id -> display-mark map where numbering restarts at `start` on every page.
 * Purely a text relabeling — callers apply it to already-placed marker text
 * without re-running layout (see the module doc comment's documented
 * approximation).
 */
export function renumberFootnotesEachPage(
  footnoteIdsByPage: ReadonlyArray<ReadonlyArray<string>>,
  numFmt: string | undefined,
  start: number,
): ReadonlyMap<string, string> {
  const marks = new Map<string, string>()

  for (const pageIds of footnoteIdsByPage) {
    let counter = start - 1
    for (const id of pageIds) {
      if (marks.has(id)) {
        continue
      }
      counter += 1
      marks.set(id, formatSequenceNumber(counter, numFmt))
    }
  }

  return marks
}

// ---------------------------------------------------------------------------
// Reference collection
// ---------------------------------------------------------------------------

/**
 * Walks one paragraph's children (including hyperlink and tracked-change
 * wrappers) collecting every footnote/endnote reference in document order.
 * Mirrors the traversal `paginate.ts`'s `collectParagraphRuns` already does
 * for runs, but footnote/endnote references are markers in their own right
 * (not text content), so they're collected separately here rather than
 * folded into that run-collection pass.
 */
export function collectNoteReferences(children: ReadonlyArray<ParagraphChild>): ReadonlyArray<NoteReference> {
  const refs: NoteReference[] = []

  for (const child of children) {
    if (isNoteReference(child)) {
      refs.push(toNoteReference(child))
      continue
    }

    if (child.kind === 'run') {
      for (const runChild of child.children) {
        if (isNoteReference(runChild)) {
          refs.push(toNoteReference(runChild))
        }
      }
      continue
    }

    if (child.kind === 'hyperlink') {
      refs.push(...collectHyperlinkNoteReferences(child.children))
      continue
    }

    if (child.kind === 'ins-revision' || child.kind === 'del-revision') {
      refs.push(...collectNoteReferences(child.children))
    }
  }

  return refs
}

function collectHyperlinkNoteReferences(children: ReadonlyArray<HyperlinkChild>): ReadonlyArray<NoteReference> {
  const refs: NoteReference[] = []

  for (const child of children) {
    if (isNoteReference(child)) {
      refs.push(toNoteReference(child))
      continue
    }

    if (child.kind === 'run') {
      for (const runChild of child.children) {
        if (isNoteReference(runChild)) {
          refs.push(toNoteReference(runChild))
        }
      }
    }
  }

  return refs
}

function isNoteReference(node: unknown): node is FootnoteReference | EndnoteReference {
  const kind = (node as { kind?: unknown }).kind
  return kind === 'footnote-reference' || kind === 'endnote-reference'
}

function toNoteReference(node: FootnoteReference | EndnoteReference): NoteReference {
  return { kind: node.kind === 'footnote-reference' ? 'footnote' : 'endnote', id: node.id }
}

// ---------------------------------------------------------------------------
// Sequence-number formatting (subset of `listMarkers.ts`'s `formatCounterValue`)
// ---------------------------------------------------------------------------

export function formatSequenceNumber(value: number, numFmt: string | undefined): string {
  switch (numFmt) {
    case 'upperRoman':
      return toRoman(value).toUpperCase()
    case 'lowerRoman':
      return toRoman(value).toLowerCase()
    case 'upperLetter':
      return toAlpha(value).toUpperCase()
    case 'lowerLetter':
      return toAlpha(value).toLowerCase()
    case 'chicago':
      return CHICAGO_MARKS[(value - 1) % CHICAGO_MARKS.length].repeat(
        Math.floor((value - 1) / CHICAGO_MARKS.length) + 1,
      )
    case 'decimal':
    default:
      return String(value)
  }
}

/** Word's "Chicago" footnote/endnote format: *, †, ‡, §, repeating doubled/tripled after one full cycle. */
const CHICAGO_MARKS: ReadonlyArray<string> = ['*', '†', '‡', '§']

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

/** Same whole-alphabet-block convention as `listMarkers.ts`'s `toAlpha` (..., z, aa, bb, ...). */
function toAlpha(value: number): string {
  if (value <= 0) {
    return String(value)
  }

  const letterIndex = (value - 1) % 26
  const repeatCount = Math.floor((value - 1) / 26) + 1
  return String.fromCharCode(97 + letterIndex).repeat(repeatCount)
}
