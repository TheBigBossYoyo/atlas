/**
 * Atlas — DOCX table-style conditional-formatting cascade resolver
 * (D7 / DXP-06, DXL-08, DXS-05)
 *
 * `<w:tblStylePr>` blocks let a table style vary its formatting by table
 * "region" — header row, banded rows/columns, first/last column, and the
 * four corner-cell intersections — the same mechanism Word's own Table
 * Style Options checkboxes (Header Row, Total Row, First/Last Column,
 * Banded Rows/Columns) drive. `parser/styles.ts` parses these into each
 * table `Style`'s `conditionalFormats` map; this module resolves them
 * (walking the `basedOn` chain, same as `cascade.ts` does for paragraph/run
 * styles) against a specific table instance's `tblLook` flags and a cell's
 * position, so `layout/layoutTable.ts` can render actual header shading and
 * row/column banding instead of only ever consulting a table's own direct
 * `w:tblPr`/`w:tcPr` formatting.
 *
 * Deliberately separate from `cascade.ts` (this branch's file-ownership
 * boundary keeps that file untouched): table-style resolution has its own
 * shape (a table style's `basedOn` chain plus a position-gated conditional
 * overlay) that doesn't fit `cascade.ts`'s paragraph/run-focused API.
 *
 * Fidelity scope: resolves table/cell-level formatting (borders, shading,
 * width/layout, cell margins/vertical-alignment) — everything
 * `layoutTable.ts` itself renders. A conditional format's `paragraph`/`run`
 * fields are parsed/serialized for round-trip fidelity (closing DXS-05's
 * "loses all table formatting on save" data loss) and returned by
 * `resolveTableCellStyle` for a future caller, but are not yet consumed by
 * this branch's layout pipeline — paragraph/run style cascade generally
 * isn't wired into the live layout/render path yet (that is a separate,
 * larger undertaking tracked as D1), so this scope boundary matches the
 * rest of the pipeline rather than being a gap specific to table styles.
 */

import { DocxParseError } from './unzip'
import type {
  BorderSet,
  InsetSet,
  ParaProps,
  RunProps,
  Shading,
  Style,
  TableCellProps,
  TableConditionalFormatType,
  TableLook,
  TableRowProps,
  TableStyleProps,
} from '../model'

type Mutable<T> = {
  -readonly [K in keyof T]: T[K]
}

/** Position of one cell within its table, for gating which conditional formats apply. */
export interface TableCellStyleContext {
  /** 0-based row index among all rows in the table, in source order. */
  readonly rowIndex: number
  readonly rowCount: number
  /** 0-based grid column where this cell begins (accounts for earlier gridSpans in the row). */
  readonly columnStart: number
  /** Number of grid columns this cell spans (DOCX `gridSpan`). */
  readonly gridSpan: number
  readonly columnCount: number
}

export interface ResolvedTableCellStyle {
  readonly cell?: TableCellProps
  readonly row?: TableRowProps
  readonly paragraph?: ParaProps
  readonly run?: RunProps
}

const MAX_BASED_ON_DEPTH = 20

/**
 * Resolves the effective table-level formatting for a `tblStyle` reference:
 * each style's own direct `w:tblPr` plus its explicit `w:tblStylePr
 * type="wholeTable"` block (equivalent, rarer, override), walked from the
 * `basedOn` root down to the referenced style so a derived style's
 * settings win over its ancestor's.
 *
 * Returns `undefined` when `styles` isn't available (callers that don't
 * have the document's style map at hand, e.g. layout call sites that
 * predate this feature) or the table has no `tblStyle` / an unresolvable
 * one — callers should fall back to the table's own direct `w:tblPr` in
 * that case, exactly as they did before this resolver existed.
 */
export function resolveTableStyle(
  tblStyleId: string | undefined,
  styles: ReadonlyMap<string, Style> | undefined,
): TableStyleProps | undefined {
  if (styles === undefined) {
    return undefined
  }

  let resolved: TableStyleProps | undefined
  for (const style of resolveTableStyleChain(tblStyleId, styles)) {
    resolved = mergeTableStyleProps(resolved, style.table)
    resolved = mergeTableStyleProps(resolved, style.conditionalFormats?.get('wholeTable')?.table)
  }

  return resolved
}

/**
 * Resolves the effective conditional-formatting overlay for one specific
 * cell: which `TableConditionalFormatType` regions are "active" for its
 * position (gated by `tblLook`'s Header Row/Total Row/First-Last
 * Column/Banded Rows-Columns flags), then layers their `tcPr`/`trPr`
 * (and `pPr`/`rPr`, for a future consumer) in Word's documented precedence
 * — whole table, then column banding, then row banding, then last/first
 * column, then last/first row, then the four corner-cell intersections
 * (most specific, highest precedence) — across the `basedOn` chain.
 *
 * Returns `{}` (every field `undefined`) when `styles` is unavailable, the
 * table has no resolvable `tblStyle`, or no conditional format is active
 * for this cell — callers should treat that as "no conditional formatting
 * applies" and fall back entirely to the cell's own direct properties.
 */
export function resolveTableCellStyle(
  tblStyleId: string | undefined,
  styles: ReadonlyMap<string, Style> | undefined,
  tblLook: TableLook | undefined,
  context: TableCellStyleContext,
): ResolvedTableCellStyle {
  if (styles === undefined) {
    return {}
  }

  const chain = resolveTableStyleChain(tblStyleId, styles)
  if (chain.length === 0) {
    return {}
  }

  const resolvedTableProps = resolveTableStyle(tblStyleId, styles)
  const activeTypes = resolveActiveConditionalTypes(tblLook, context, resolvedTableProps)

  let cell: TableCellProps | undefined
  let row: TableRowProps | undefined
  let paragraph: ParaProps | undefined
  let run: RunProps | undefined

  for (const style of chain) {
    for (const type of activeTypes) {
      const conditional = style.conditionalFormats?.get(type)
      if (conditional === undefined) {
        continue
      }
      cell = mergeTableCellProps(cell, conditional.cell)
      row = mergeTableRowProps(row, conditional.row)
      paragraph = mergeParaPropsShallow(paragraph, conditional.paragraph)
      run = mergeRunPropsShallow(run, conditional.run)
    }
  }

  return {
    ...(cell !== undefined ? { cell } : {}),
    ...(row !== undefined ? { row } : {}),
    ...(paragraph !== undefined ? { paragraph } : {}),
    ...(run !== undefined ? { run } : {}),
  }
}

/** A band size must be a positive integer; anything else falls back to Word's default of 1. */
function normalizeBandSize(value: number | undefined): number {
  return typeof value === 'number' && Number.isInteger(value) && value > 0 ? value : 1
}

function resolveTableStyleChain(
  styleId: string | undefined,
  styles: ReadonlyMap<string, Style>,
  depth: number = 0,
): ReadonlyArray<Style> {
  if (styleId === undefined || styleId === '') {
    return []
  }

  if (depth > MAX_BASED_ON_DEPTH) {
    throw new DocxParseError(`Table style basedOn chain exceeded ${MAX_BASED_ON_DEPTH} levels while resolving "${styleId}"`)
  }

  const style = styles.get(styleId)
  if (style === undefined) {
    return []
  }

  const ancestors = style.basedOn !== undefined ? resolveTableStyleChain(style.basedOn, styles, depth + 1) : []
  return [...ancestors, style]
}

/**
 * Determines which `TableConditionalFormatType` regions apply to a cell at
 * `context`'s position, in Word's documented cascade precedence (each
 * pushed entry overrides everything pushed before it once merged).
 *
 * The precedence order below is Word's actual applied order, NOT the order
 * literally printed in ECMA-376 §17.7.6.6 (whole table, banded columns,
 * banded rows, first/last row, first/last column, corners): MS-OI29500
 * §2.1.250 documents that Word deviates from its own published spec and
 * applies row banding *before* column banding (so column banding overrides
 * row banding on a conflicting field, not the other way around) — verified
 * against Microsoft's own worked example ("OpenXML Styles 101") showing
 * column banding visibly winning over row banding in real Word output.
 * Getting this order backward silently swaps which stripe color wins
 * wherever a table style bands both rows and columns with different fills.
 *
 * `resolvedTableProps`'s `rowBandSize`/`colBandSize` set how many
 * consecutive rows/columns each stripe covers (Word defaults both to 1 —
 * alternate every single row/column — when unset). Banding doesn't exclude
 * header/footer rows from the stripe count the way Word itself does — a
 * documented simplification, consistent with this codebase's existing
 * practice of noting fidelity trade-offs (e.g. `fonts/canvasMetrics.ts`'s
 * small-caps measurement, DXP-17) rather than silently guessing.
 */
function resolveActiveConditionalTypes(
  tblLook: TableLook | undefined,
  context: TableCellStyleContext,
  resolvedTableProps: TableStyleProps | undefined,
): ReadonlyArray<TableConditionalFormatType> {
  const active: TableConditionalFormatType[] = ['wholeTable']

  const bandRowsOn = tblLook?.noHBand !== true
  const bandColsOn = tblLook?.noVBand !== true
  const firstRowOn = tblLook?.firstRow === true
  const lastRowOn = tblLook?.lastRow === true
  const firstColOn = tblLook?.firstColumn === true
  const lastColOn = tblLook?.lastColumn === true

  const isFirstRow = context.rowIndex === 0
  const isLastRow = context.rowIndex === context.rowCount - 1
  const isFirstCol = context.columnStart === 0
  const isLastCol = context.columnStart + context.gridSpan >= context.columnCount

  const rowBandSize = normalizeBandSize(resolvedTableProps?.rowBandSize)
  const colBandSize = normalizeBandSize(resolvedTableProps?.colBandSize)

  // Row banding first, then column banding — column banding wins ties
  // (Word's actual behavior; see this function's doc comment).
  if (bandRowsOn) {
    const stripeIndex = Math.floor(context.rowIndex / rowBandSize)
    active.push(stripeIndex % 2 === 0 ? 'band1Horz' : 'band2Horz')
  }
  if (bandColsOn) {
    const stripeIndex = Math.floor(context.columnStart / colBandSize)
    active.push(stripeIndex % 2 === 0 ? 'band1Vert' : 'band2Vert')
  }
  if (lastColOn && isLastCol) active.push('lastCol')
  if (firstColOn && isFirstCol) active.push('firstCol')
  if (lastRowOn && isLastRow) active.push('lastRow')
  if (firstRowOn && isFirstRow) active.push('firstRow')
  if (lastRowOn && isLastRow && firstColOn && isFirstCol) active.push('swCell')
  if (lastRowOn && isLastRow && lastColOn && isLastCol) active.push('seCell')
  if (firstRowOn && isFirstRow && lastColOn && isLastCol) active.push('neCell')
  if (firstRowOn && isFirstRow && firstColOn && isFirstCol) active.push('nwCell')

  return active
}

// ---------------------------------------------------------------------------
// Merge helpers — field-level shallow merge (override wins), with a nested
// merge for the handful of sub-objects (borders/shading/margins) whose
// individual edges/fields a lower cascade level may still usefully supply
// when a higher level only overrides some of them.
// ---------------------------------------------------------------------------

function mergeTableStyleProps(
  base: TableStyleProps | undefined,
  override: TableStyleProps | undefined,
): TableStyleProps | undefined {
  if (base === undefined && override === undefined) {
    return undefined
  }

  const merged: Mutable<TableStyleProps> = { ...(base ?? {}), ...(override ?? {}) }

  const borders = mergeBorderSet(base?.borders, override?.borders)
  if (borders !== undefined) merged.borders = borders

  const shading = mergeShading(base?.shading, override?.shading)
  if (shading !== undefined) merged.shading = shading

  const cellMargin = mergeInsetSet(base?.cellMargin, override?.cellMargin)
  if (cellMargin !== undefined) merged.cellMargin = cellMargin

  return merged
}

function mergeTableRowProps(
  base: TableRowProps | undefined,
  override: TableRowProps | undefined,
): TableRowProps | undefined {
  if (base === undefined && override === undefined) {
    return undefined
  }
  return { ...(base ?? {}), ...(override ?? {}) }
}

function mergeTableCellProps(
  base: TableCellProps | undefined,
  override: TableCellProps | undefined,
): TableCellProps | undefined {
  if (base === undefined && override === undefined) {
    return undefined
  }

  const merged: Mutable<TableCellProps> = { ...(base ?? {}), ...(override ?? {}) }

  const tcBorders = mergeBorderSet(base?.tcBorders, override?.tcBorders)
  if (tcBorders !== undefined) merged.tcBorders = tcBorders

  const shd = mergeShading(base?.shd, override?.shd)
  if (shd !== undefined) merged.shd = shd

  const tcMar = mergeInsetSet(base?.tcMar, override?.tcMar)
  if (tcMar !== undefined) merged.tcMar = tcMar

  return merged
}

function mergeParaPropsShallow(base: ParaProps | undefined, override: ParaProps | undefined): ParaProps | undefined {
  if (base === undefined && override === undefined) {
    return undefined
  }
  return { ...(base ?? {}), ...(override ?? {}) }
}

function mergeRunPropsShallow(base: RunProps | undefined, override: RunProps | undefined): RunProps | undefined {
  if (base === undefined && override === undefined) {
    return undefined
  }
  return { ...(base ?? {}), ...(override ?? {}) }
}

function mergeBorderSet(base: BorderSet | undefined, override: BorderSet | undefined): BorderSet | undefined {
  if (base === undefined && override === undefined) {
    return undefined
  }
  return { ...(base ?? {}), ...(override ?? {}) }
}

function mergeShading(base: Shading | undefined, override: Shading | undefined): Shading | undefined {
  if (base === undefined && override === undefined) {
    return undefined
  }
  return { ...(base ?? {}), ...(override ?? {}) }
}

function mergeInsetSet(base: InsetSet | undefined, override: InsetSet | undefined): InsetSet | undefined {
  if (base === undefined && override === undefined) {
    return undefined
  }
  return { ...(base ?? {}), ...(override ?? {}) }
}
