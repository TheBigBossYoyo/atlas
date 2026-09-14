import type { BorderSet, Style, Table } from '../model'
import type { Theme } from '../parser/theme'

import type { FontResolver, LineBox } from './types'

export type LaidOutTable = {
  widthPt: number
  columnWidthsPt: ReadonlyArray<number>
  rows: ReadonlyArray<LaidOutRow>
  repeatHeaderRowCount: number
  /**
   * Table-level borders resolved from `<w:tblBorders>`. Renderer applies these
   * as a fallback when a cell has no explicit border on a given edge.
   */
  borders?: BorderSet
  /**
   * Table-level shading fill resolved from `<w:shd w:fill="…">`. Already
   * normalized to a CSS-compatible color string (e.g. `#RRGGBB`) or
   * `undefined` when no fill applies.
   */
  shadingFill?: string
}

export type LaidOutRow = {
  heightPt: number
  cells: ReadonlyArray<LaidOutCell>
  isHeader: boolean
}

export type LaidOutCell = {
  widthPt: number
  heightPt: number
  /** Number of horizontal grid columns this cell spans (DOCX `gridSpan`). */
  gridSpan: number
  /** Number of vertical rows this cell spans (resolved from `vMerge`). */
  rowSpan: number
  /** 0-based column index where this cell begins in the table grid. */
  columnStart: number
  /**
   * Whether the renderer should emit a `<td>` for this cell.
   * `false` for `vMerge=continue` cells whose visual region is owned by the
   * `vMerge=restart` cell above. The layout still records the cell so that
   * height contribution and column accounting stay correct.
   */
  shouldRender: boolean
  vMergeStart: boolean
  vMergeContinue: boolean
  contentLines: ReadonlyArray<LineBox>
  paddingPt: {
    top: number
    right: number
    bottom: number
    left: number
  }
  vAlign: 'top' | 'center' | 'bottom'
  /** Cell-level borders resolved from `<w:tcBorders>`; overrides table borders. */
  borders?: BorderSet
  /** Cell-level shading fill as a CSS color, or `undefined` when absent. */
  shadingFill?: string
}

export type TableLayoutInput = {
  table: Table
  availableWidthPt: number
  fontResolver: FontResolver
  theme?: Theme
  /**
   * The document's style map (D7 / DXP-06, DXL-08, DXS-05) — when supplied,
   * `layoutTable` resolves the table's `tblStyle` (basedOn chain +
   * conditional formatting gated by `tblLook`, via `parser/cascadeTable.ts`)
   * and uses it as a fallback under the table/row/cell's own direct
   * formatting, so a built-in banded/colored Word table style actually
   * renders. Omitted (as existing callers that predate this feature do),
   * layout behaves exactly as before — direct formatting only.
   */
  styles?: ReadonlyMap<string, Style>
}
