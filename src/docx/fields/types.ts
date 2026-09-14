/**
 * Atlas — DOCX field evaluation context (DEFER-5 / DXS-20)
 *
 * Everything a field evaluator (`./evaluate.ts`) needs from outside the
 * field's own instruction text, gathered here so evaluators stay pure
 * functions of `(instruction, context)` rather than reaching into
 * `DocxViewer`'s layout/bundle state directly. The viewer builds one of
 * these per "Update field(s)" run (see `buildFieldEvaluationContext` in
 * `DocxViewer.tsx`) from whatever it already has at hand — docProps,
 * bookmark positions, the current pagination result.
 */
export interface FieldEvaluationContext {
  /**
   * "Now", for DATE/TIME. Injectable instead of every evaluator defaulting
   * to `new Date()` so tests are deterministic; the viewer leaves this
   * unset in real use.
   */
  readonly now?: Date
  /** `docProps/core.xml`'s `dc:creator`, for AUTHOR. `undefined` when unknown — AUTHOR then leaves the field's cached text unchanged rather than displaying an empty string. */
  readonly author?: string
  /** `docProps/core.xml`'s `dc:title` (or an equivalent app-level title), for TITLE. */
  readonly title?: string
  /** Bookmark name -> its current plain-text content, for REF's default (text) mode. */
  readonly bookmarkText: ReadonlyMap<string, string>
  /**
   * Bookmark name -> the 1-based page it currently falls on, for PAGEREF
   * (and a REF using the `\p` "above/below" switches, which this scope
   * doesn't attempt beyond falling back to plain text — see
   * `evaluateRefField`'s doc comment). `undefined` when layout hasn't run
   * yet — PAGEREF then leaves the field's cached text unchanged.
   */
  readonly bookmarkPage?: ReadonlyMap<string, number>
  /** Total page count, for NUMPAGES. `undefined` when layout hasn't run yet. */
  readonly pageCount?: number
  /**
   * Resolves a document-body paragraph's page, for a plain PAGE field —
   * `paragraphPath` is `[sectionIndex, blockIndex]` (the paragraph's index
   * among its section's direct blocks), the same shape `updateFields`
   * passes to every field it evaluates in the body. Deliberately never
   * consulted for a field inside a header/footer/footnote/endnote (see
   * `updateFields`'s doc comment for why a single fixed page number
   * doesn't make sense there) — `updateFields` enforces that itself, so an
   * implementation of this only ever needs to handle body paragraphs.
   */
  readonly currentPageOf?: (paragraphPath: ReadonlyArray<number>) => number | undefined
  /**
   * SEQ sequence-name -> its running counter, threaded across an entire
   * "Update field(s)" pass so successive SEQ fields for the same sequence
   * name increment (or `\c`-repeat, or `\r`-reset) the way Word's does.
   * `evaluateSeqField` mutates this map in place — pass a fresh `Map()` for
   * one independent update pass, matching Word's own "no memory between
   * separate Update Fields runs" behavior. Fields are evaluated in document
   * order (`updateFields` walks sections/blocks/paragraphs top to bottom),
   * so counters advance in the order a reader would encounter them.
   */
  readonly sequenceCounters: Map<string, number>
}
