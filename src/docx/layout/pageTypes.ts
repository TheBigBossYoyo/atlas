import type { Document, Table } from '../model'
import type { NotePr } from '../parser/settings'
import type { Theme } from '../parser/theme'

import type { LaidOutTable } from './tableTypes'
import type { FontResolver, LineBox } from './types'

export type Page = {
  sectionIndex: number
  pageIndex: number
  sizePt: {
    width: number
    height: number
  }
  marginsPt: {
    top: number
    right: number
    bottom: number
    left: number
    header: number
    footer: number
    gutter: number
  }
  columns: ReadonlyArray<ColumnBox>
  headerLines: ReadonlyArray<LineBox>
  footerLines: ReadonlyArray<LineBox>
  /**
   * D11 milestone 3 (DXL-09) — this page's bottom-of-page footnote area,
   * flattened across every footnote referenced on the page (in
   * first-reference order), each already positioned relative to the
   * area's own top via `topPt`. Empty when the page references no
   * footnotes — `PageView` renders nothing and reserves no space in that
   * case (mirrors `headerLines`/`footerLines`'s empty-array convention).
   */
  footnoteLines: ReadonlyArray<PageFootnoteLineRef>
  /** Whether to draw the short horizontal separator above the footnote area — true whenever `footnoteLines` is non-empty. */
  hasFootnoteSeparator: boolean
}

/** One line of one footnote's body text, positioned within the page's footnote area (see `Page.footnoteLines`). */
export type PageFootnoteLineRef = {
  noteId: string
  line: LineBox
  /** Top offset in points from the footnote area's own top edge (NOT the page). */
  topPt: number
  leftPt: number
}

export type ColumnBox = {
  widthPt: number
  leftPt: number
  lines: ReadonlyArray<PageLineRef>
  /**
   * Table fragments placed in this column. Each entry is a self-contained
   * slice of a `LaidOutTable` covering one column's worth of rows, with any
   * repeated header rows already duplicated into `slice.rows`. The renderer
   * draws each fragment as a real `<table>` absolutely positioned at
   * `(leftPt, topPt)` inside the column.
   *
   * Empty array when the column has no tables. Stored separately from
   * `lines` so paragraph selection paths and existing line-based rendering
   * are completely unaffected.
   */
  tables: ReadonlyArray<PageTableRef>
}

export type PageLineRef = {
  paragraphPath: ReadonlyArray<number>
  lineIndex: number
  line: LineBox
  topPt: number
  leftPt: number
}

/**
 * A self-contained slice of a `LaidOutTable` positioned on one column.
 *
 * - `topPt` / `leftPt` are absolute coordinates inside the page (same units
 *   as `PageLineRef`) so the renderer can position the fragment without
 *   knowing its parent column.
 * - `rows` already contains repeated header rows (for continuation
 *   fragments) tagged via `isRepeatedHeader`. The renderer never needs the
 *   original `LaidOutTable` to draw this fragment.
 * - `isContinuation` is true when this fragment is not the first slice of
 *   the source table (used by the renderer to emit `data-*` for tooling).
 */
export type PageTableRef = {
  blockPath: ReadonlyArray<number>
  topPt: number
  leftPt: number
  widthPt: number
  heightPt: number
  columnWidthsPt: ReadonlyArray<number>
  borders: LaidOutTable['borders']
  shadingFill: LaidOutTable['shadingFill']
  rows: ReadonlyArray<PageTableRowRef>
  isContinuation: boolean
}

export type PageTableRowRef = {
  sourceRowIndex: number
  isHeaderRow: boolean
  isRepeatedHeader: boolean
  row: LaidOutTable['rows'][number]
}

export type TableLayoutFn = (
  input: {
    table: Table
    availableWidth: number
    sectionIndex: number
    blockIndex: number
  },
) => Promise<ReadonlyArray<number>> | ReadonlyArray<number>

export type PaginationProgress = {
  readonly phase: 'layout'
  readonly completedBlocks: number
  readonly totalBlocks: number
}

/**
 * Thrown by `paginate` when `shouldCancel` returns true. Callers should
 * treat this as a non-error: it just means a newer pagination superseded
 * this one (typical React StrictMode / rapid document edits).
 */
export class PaginationCancelledError extends Error {
  constructor() {
    super('DOCX pagination cancelled')
    this.name = 'PaginationCancelledError'
  }
}

export type PaginatorInput = {
  document: Document
  fontResolver: FontResolver
  /**
   * Overrides (or, for an id this document doesn't otherwise define,
   * supplements) the header/footer content `paginate` itself builds from
   * `document.headers`/`document.footers` (D11 milestone 1) — mainly a
   * test seam at this point; production callers can omit it entirely.
   */
  headerFooterLines?: Map<string, ReadonlyArray<LineBox>>
  tableLayout?: TableLayoutFn
  theme?: Theme
  /**
   * `word/settings.xml`'s `w:evenAndOddHeaders` (D11 milestone 1/DXL-09):
   * when false/absent (the common case), an `even`-typed header/footer
   * reference is never selected even if the document happens to define
   * one — every page uses the `default` reference regardless of parity,
   * matching Word's own behavior for a document that never turned this
   * setting on.
   */
  evenAndOddHeaders?: boolean
  /** `word/settings.xml`'s `w:footnotePr` (D11 milestone 5) — default footnote numbering format/restart/start. */
  footnoteNumbering?: NotePr
  /** `word/settings.xml`'s `w:endnotePr` (D11 milestone 5) — default endnote numbering format/restart/start. */
  endnoteNumbering?: NotePr
  /**
   * Called after each block (paragraph or table) is itemized + measured.
   * Lets the UI render a progress indicator during long pagination on
   * large documents (e.g. 1000+ blocks). Optional; tests omit it.
   */
  onProgress?: (progress: PaginationProgress) => void
  /**
   * Polled before each yield point. When it returns true, `paginate`
   * throws `PaginationCancelledError` so callers can abort cleanly.
   */
  shouldCancel?: () => boolean
}
