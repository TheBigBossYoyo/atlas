import type { Document, Table } from '../model'
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
  headerFooterLines?: Map<string, ReadonlyArray<LineBox>>
  tableLayout?: TableLayoutFn
  theme?: Theme
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
