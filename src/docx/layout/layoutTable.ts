import type {
  Block,
  HyperlinkChild,
  InsetSet,
  Paragraph,
  ParagraphChild,
  Shading,
  Style,
  TableCellProps,
  TableLook,
  TabAlignment,
  TabLeader,
  Table,
  TableCell,
  TableRow,
  Width,
} from '../model'
import { resolveTableCellStyle, resolveTableStyle } from '../parser/cascadeTable'

import { breakLines } from './breakLines'
import { itemizeRuns } from './itemize'
import type { LaidOutCell, LaidOutRow, LaidOutTable, TableLayoutInput } from './tableTypes'
import type { LineBreakInput, LineItem, TabStop } from './types'

/**
 * Per-table context `layoutRow`/the per-cell loop need to resolve D7's
 * table-style conditional formatting for a given cell position — computed
 * once in `layoutTable` and threaded down rather than re-derived per row.
 */
type TableStyleCascadeContext = {
  tblStyleId: string | undefined
  styles: ReadonlyMap<string, Style> | undefined
  tblLook: TableLook | undefined
  rowCount: number
  columnCount: number
}

type BoxEdges = {
  top: number
  right: number
  bottom: number
  left: number
}

type ColumnConstraint = {
  minWidthPt: number
  maxWidthPt: number
}

type CellLayoutDraft = {
  naturalHeightPt: number
  laidOutCell: Omit<LaidOutCell, 'heightPt'>
  contributesToRowHeight: boolean
}

export async function layoutTable(input: TableLayoutInput): Promise<LaidOutTable> {
  const rows = input.table.rows.filter(isTableRow)
  const repeatHeaderRowCount = countRepeatHeaderRows(rows)
  const columnCount = resolveColumnCount(input.table, rows)
  // D7 / DXP-06, DXL-08, DXS-05: the table's own direct `w:tblPr` always
  // wins; the resolved table style is only a fallback for whatever the
  // table instance doesn't set directly (`?? resolvedTableStyle?.field`
  // below), exactly mirroring how a cell's own direct `w:tcPr` wins over
  // its cell-level conditional formatting further down.
  const resolvedTableStyle = resolveTableStyle(input.table.props?.tblStyle, input.styles)
  const preferredTableWidthPt =
    resolveWidthToPoints(input.table.props?.tblW ?? resolvedTableStyle?.width, input.availableWidthPt)
    ?? clampNonNegative(input.availableWidthPt)
  const fixedColumnWidthsPt = resolveFixedColumnWidths(input.table, columnCount)
  const isFixedLayout = (input.table.props?.tblLayout ?? resolvedTableStyle?.layout) === 'fixed'

  const columnWidthsPt =
    isFixedLayout && fixedColumnWidthsPt.length > 0
      ? fixedColumnWidthsPt
      : await resolveAutofitColumnWidths(rows, columnCount, preferredTableWidthPt, input.fontResolver, input.theme)

  const cascadeContext: TableStyleCascadeContext = {
    tblStyleId: input.table.props?.tblStyle,
    styles: input.styles,
    tblLook: input.table.props?.tblLook,
    rowCount: rows.length,
    columnCount,
  }

  const laidOutRows = await Promise.all(
    rows.map((row, rowIndex) =>
      layoutRow(
        row,
        rowIndex,
        rowIndex < repeatHeaderRowCount,
        columnWidthsPt,
        input.table.props?.tblCellMar,
        input.fontResolver,
        input.theme,
        cascadeContext,
      ),
    ),
  )

  const resolvedRows = resolveVerticalMerges(laidOutRows)

  const effectiveBorders = input.table.props?.tblBorders ?? resolvedTableStyle?.borders
  const effectiveShadingFill =
    resolveShadingFill(input.table.props?.shd) ?? resolveShadingFill(resolvedTableStyle?.shading)

  return {
    widthPt: sumNumbers(columnWidthsPt),
    columnWidthsPt,
    rows: resolvedRows,
    repeatHeaderRowCount,
    ...(effectiveBorders !== undefined ? { borders: effectiveBorders } : {}),
    ...(effectiveShadingFill !== undefined ? { shadingFill: effectiveShadingFill } : {}),
  }
}

/**
 * Walk laid-out rows and resolve DOCX `vMerge` chains so the renderer can emit
 * proper `rowSpan` on the originating cell and skip continuation cells.
 *
 * For every cell with `vMerge=restart`, find the contiguous run of
 * `vMerge=continue` cells beneath it at the same `columnStart` and accumulate
 * the row count into the starter's `rowSpan`. Continuation cells get
 * `shouldRender=false`.
 */
function resolveVerticalMerges(rows: ReadonlyArray<LaidOutRow>): ReadonlyArray<LaidOutRow> {
  const cellsByRow = rows.map((row) => row.cells.map((cell) => ({ ...cell })))

  for (let rowIndex = 0; rowIndex < cellsByRow.length; rowIndex += 1) {
    for (const cell of cellsByRow[rowIndex]) {
      if (!cell.vMergeStart) {
        continue
      }

      let span = 1
      for (let nextRow = rowIndex + 1; nextRow < cellsByRow.length; nextRow += 1) {
        const continuation = cellsByRow[nextRow].find(
          (candidate) => candidate.columnStart === cell.columnStart && candidate.vMergeContinue,
        )
        if (continuation === undefined) {
          break
        }
        continuation.shouldRender = false
        span += 1
      }

      cell.rowSpan = span
    }
  }

  return rows.map((row, rowIndex) => ({
    ...row,
    cells: cellsByRow[rowIndex],
  }))
}

async function resolveAutofitColumnWidths(
  rows: ReadonlyArray<TableRow>,
  columnCount: number,
  preferredTableWidthPt: number,
  fontResolver: TableLayoutInput['fontResolver'],
  theme: TableLayoutInput['theme'],
): Promise<ReadonlyArray<number>> {
  if (columnCount === 0) {
    return []
  }

  const constraints: ColumnConstraint[] = Array.from({ length: columnCount }, () => ({
    minWidthPt: 0,
    maxWidthPt: 0,
  }))

  for (const row of rows) {
    let columnIndex = 0

    for (const child of row.cells) {
      if (!isTableCell(child)) {
        continue
      }

      const gridSpan = normalizeGridSpan(child.props?.gridSpan, columnCount - columnIndex)
      const measuredWidth = await measureCellWidth(child, preferredTableWidthPt, fontResolver, theme)
      applyCellWidthToColumns(constraints, columnIndex, gridSpan, measuredWidth.minWidthPt, measuredWidth.maxWidthPt)
      columnIndex += gridSpan
    }
  }

  return distributeColumnWidths(constraints, preferredTableWidthPt)
}

function resolveFixedColumnWidths(table: Table, columnCount: number): ReadonlyArray<number> {
  const grid = getTableGrid(table)
  if (grid.length === 0) {
    return []
  }

  const widthCount = Math.max(columnCount, grid.length)
  return Array.from({ length: widthCount }, (_, index) => twipToPt(grid[index]))
}

function getTableGrid(table: Table): ReadonlyArray<number> {
  return table.tblGrid ?? []
}

function resolveColumnCount(table: Table, rows: ReadonlyArray<TableRow>): number {
  let maxColumnCount = getTableGrid(table).length

  for (const row of rows) {
    let rowColumnCount = 0
    for (const child of row.cells) {
      if (isTableCell(child)) {
        rowColumnCount += normalizeGridSpan(child.props?.gridSpan)
      }
    }
    maxColumnCount = Math.max(maxColumnCount, rowColumnCount)
  }

  return maxColumnCount
}

async function measureCellWidth(
  cell: TableCell,
  preferredTableWidthPt: number,
  fontResolver: TableLayoutInput['fontResolver'],
  theme: TableLayoutInput['theme'],
): Promise<{ minWidthPt: number; maxWidthPt: number }> {
  const paragraphs = collectParagraphs(cell.blocks)
  let minWidthPt = 0
  let maxWidthPt = 0

  for (const paragraph of paragraphs) {
    const measuredParagraph = await measureParagraphWidth(paragraph, fontResolver, theme)
    minWidthPt = Math.max(minWidthPt, measuredParagraph.minWidthPt)
    maxWidthPt = Math.max(maxWidthPt, measuredParagraph.maxWidthPt)
  }

  const preferredWidthPt = resolveWidthToPoints(cell.props?.tcW, preferredTableWidthPt) ?? 0
  return {
    minWidthPt: Math.max(minWidthPt, preferredWidthPt),
    maxWidthPt: Math.max(maxWidthPt, preferredWidthPt),
  }
}

async function measureParagraphWidth(
  paragraph: Paragraph,
  fontResolver: TableLayoutInput['fontResolver'],
  theme: TableLayoutInput['theme'],
): Promise<{ minWidthPt: number; maxWidthPt: number }> {
  const items = await itemizeRuns(collectRuns(paragraph), fontResolver, theme)
  return measureItems(items)
}

function measureItems(items: ReadonlyArray<LineItem>): { minWidthPt: number; maxWidthPt: number } {
  let minWidthPt = 0
  let maxWidthPt = 0
  let currentSegmentWidthPt = 0
  let trailingWhitespaceWidthPt = 0

  for (const item of items) {
    switch (item.kind) {
      case 'word':
      case 'glyph-cluster':
      case 'drawing': {
        minWidthPt = Math.max(minWidthPt, item.width)
        currentSegmentWidthPt += item.width
        trailingWhitespaceWidthPt = 0
        break
      }
      case 'space':
      case 'tab': {
        currentSegmentWidthPt += item.width
        trailingWhitespaceWidthPt += item.width
        break
      }
      case 'break': {
        maxWidthPt = Math.max(maxWidthPt, currentSegmentWidthPt - trailingWhitespaceWidthPt)
        currentSegmentWidthPt = 0
        trailingWhitespaceWidthPt = 0
        break
      }
      case 'hyphen-opportunity': {
        break
      }
    }
  }

  maxWidthPt = Math.max(maxWidthPt, currentSegmentWidthPt - trailingWhitespaceWidthPt)
  return {
    minWidthPt,
    maxWidthPt,
  }
}

function applyCellWidthToColumns(
  constraints: ColumnConstraint[],
  startColumnIndex: number,
  gridSpan: number,
  minWidthPt: number,
  maxWidthPt: number,
): void {
  if (startColumnIndex >= constraints.length) {
    return
  }

  const effectiveGridSpan = Math.min(gridSpan, constraints.length - startColumnIndex)
  const minSharePt = minWidthPt / effectiveGridSpan
  const maxSharePt = maxWidthPt / effectiveGridSpan

  for (let columnOffset = 0; columnOffset < effectiveGridSpan; columnOffset += 1) {
    const column = constraints[startColumnIndex + columnOffset]
    column.minWidthPt = Math.max(column.minWidthPt, minSharePt)
    column.maxWidthPt = Math.max(column.maxWidthPt, maxSharePt)
  }
}

function distributeColumnWidths(
  constraints: ReadonlyArray<ColumnConstraint>,
  preferredTableWidthPt: number,
): ReadonlyArray<number> {
  const minWidthsPt = constraints.map((constraint) => clampNonNegative(constraint.minWidthPt))
  const maxWidthsPt = constraints.map((constraint, index) =>
    Math.max(minWidthsPt[index], clampNonNegative(constraint.maxWidthPt)),
  )

  const totalMinWidthPt = sumNumbers(minWidthsPt)
  const totalMaxWidthPt = sumNumbers(maxWidthsPt)

  if (totalMaxWidthPt <= preferredTableWidthPt) {
    return maxWidthsPt
  }

  if (totalMinWidthPt >= preferredTableWidthPt) {
    return minWidthsPt
  }

  const totalShrinkPt = totalMaxWidthPt - preferredTableWidthPt
  const shrinkCapacitiesPt = maxWidthsPt.map((maxWidthPt, index) => maxWidthPt - minWidthsPt[index])
  const totalShrinkCapacityPt = sumNumbers(shrinkCapacitiesPt)

  if (totalShrinkCapacityPt <= 0) {
    return minWidthsPt
  }

  return maxWidthsPt.map(
    (maxWidthPt, index) => maxWidthPt - totalShrinkPt * (shrinkCapacitiesPt[index] / totalShrinkCapacityPt),
  )
}

async function layoutRow(
  row: TableRow,
  rowIndex: number,
  isHeader: boolean,
  columnWidthsPt: ReadonlyArray<number>,
  defaultCellMargins: InsetSet | undefined,
  fontResolver: TableLayoutInput['fontResolver'],
  theme: TableLayoutInput['theme'],
  cascadeContext: TableStyleCascadeContext,
): Promise<LaidOutRow> {
  const drafts: CellLayoutDraft[] = []
  let columnIndex = 0

  for (const rawChild of row.cells) {
    if (!isTableCell(rawChild)) {
      continue
    }

    const gridSpan = normalizeGridSpan(rawChild.props?.gridSpan, columnWidthsPt.length - columnIndex)
    const widthPt = sumNumbers(columnWidthsPt.slice(columnIndex, columnIndex + gridSpan))

    drafts.push(
      await layoutTableCellDraft(
        rawChild,
        { rowIndex, columnIndex, gridSpan, widthPt },
        defaultCellMargins,
        fontResolver,
        theme,
        cascadeContext,
      ),
    )

    columnIndex += gridSpan
  }

  const naturalRowHeightPt = drafts.reduce(
    (heightPt, draft) =>
      draft.contributesToRowHeight ? Math.max(heightPt, draft.naturalHeightPt) : heightPt,
    0,
  )
  const heightPt = resolveRowHeight(row, naturalRowHeightPt)

  return {
    heightPt,
    cells: drafts.map((draft) => ({
      ...draft.laidOutCell,
      heightPt,
    })),
    isHeader,
    cantSplit: row.props?.cantSplit === true,
  }
}

type CellPosition = {
  rowIndex: number
  columnIndex: number
  gridSpan: number
  widthPt: number
}

/**
 * Lays out one cell, first resolving D7's table-style conditional
 * formatting (header shading, row/column banding, corner cells) for this
 * cell's position and layering it under the cell's own direct `w:tcPr` —
 * direct formatting always wins, exactly like every other style cascade in
 * this codebase. The resulting merged, "effective" cell is used for every
 * property lookup below; its content (`blocks`) is untouched.
 */
async function layoutTableCellDraft(
  rawChild: TableCell,
  position: CellPosition,
  defaultCellMargins: InsetSet | undefined,
  fontResolver: TableLayoutInput['fontResolver'],
  theme: TableLayoutInput['theme'],
  cascadeContext: TableStyleCascadeContext,
): Promise<CellLayoutDraft> {
  const conditional = resolveTableCellStyle(cascadeContext.tblStyleId, cascadeContext.styles, cascadeContext.tblLook, {
    rowIndex: position.rowIndex,
    rowCount: cascadeContext.rowCount,
    columnStart: position.columnIndex,
    gridSpan: position.gridSpan,
    columnCount: cascadeContext.columnCount,
  })
  const child: TableCell = conditional.cell === undefined
    ? rawChild
    : { ...rawChild, props: mergeCellPropsWithConditional(conditional.cell, rawChild.props) }

  const paddingPt = resolveCellPadding(child, defaultCellMargins, position.widthPt)
  const contentLines = await layoutCellLines(child, position.widthPt, paddingPt, fontResolver, theme)
  const naturalHeightPt = paddingPt.top + paddingPt.bottom + sumNumbers(contentLines.map((line) => line.lineHeight))
  const vMergeStart = child.props?.vMerge === 'restart'
  const vMergeContinue = child.props?.vMerge === 'continue'
  const cellShadingFill = resolveShadingFill(child.props?.shd)

  return {
    naturalHeightPt,
    laidOutCell: {
      widthPt: position.widthPt,
      gridSpan: position.gridSpan,
      rowSpan: 1,
      columnStart: position.columnIndex,
      shouldRender: true,
      vMergeStart,
      vMergeContinue,
      contentLines,
      paddingPt,
      vAlign: resolveVerticalAlign(child),
      ...(child.props?.tcBorders !== undefined ? { borders: child.props.tcBorders } : {}),
      ...(cellShadingFill !== undefined ? { shadingFill: cellShadingFill } : {}),
    },
    contributesToRowHeight: !vMergeContinue,
  }
}

async function layoutCellLines(
  cell: TableCell,
  widthPt: number,
  paddingPt: BoxEdges,
  fontResolver: TableLayoutInput['fontResolver'],
  theme: TableLayoutInput['theme'],
) {
  const contentWidthPt = Math.max(0, widthPt - paddingPt.left - paddingPt.right)
  const paragraphs = collectParagraphs(cell.blocks)
  const contentLines = []

  for (const paragraph of paragraphs) {
    const lines = await breakLines({
      paragraph,
      paraProps: paragraph.props ?? {},
      runs: collectRuns(paragraph),
      availableWidth: contentWidthPt,
      fontResolver,
      tabStops: resolveTabStops(paragraph),
      theme,
    })
    contentLines.push(...lines)
  }

  return contentLines
}

function collectParagraphs(blocks: ReadonlyArray<Block>): ReadonlyArray<Paragraph> {
  return blocks.filter(isParagraph)
}

function collectRuns(paragraph: Paragraph): LineBreakInput['runs'] {
  const runs: LineBreakInput['runs'][number][] = []

  for (const child of paragraph.children) {
    collectRunsFromParagraphChild(child, runs)
  }

  return runs
}

function collectRunsFromParagraphChild(
  child: ParagraphChild,
  runs: LineBreakInput['runs'][number][],
): void {
  if (child.kind === 'run') {
    runs.push({
      run: child,
      runProps: child.props ?? {},
    })
    return
  }

  if (child.kind === 'hyperlink') {
    for (const hyperlinkChild of child.children) {
      collectRunsFromHyperlinkChild(hyperlinkChild, runs)
    }
    return
  }

  if (child.kind === 'field') {
    // DEFER-5 / DXS-20 — a field's cached `result` IS its visible content
    // (see the matching case in `paginate.ts`'s `collectParagraphRuns`):
    // without this, a field inside a table cell (e.g. a DATE/AUTHOR field,
    // or a cross-reference, in a table) would silently contribute no
    // visible text at all.
    for (const resultChild of child.result) {
      collectRunsFromParagraphChild(resultChild, runs)
    }
  }
}

function collectRunsFromHyperlinkChild(
  child: HyperlinkChild,
  runs: LineBreakInput['runs'][number][],
): void {
  if (child.kind === 'run') {
    runs.push({
      run: child,
      runProps: child.props ?? {},
    })
    return
  }

  if (child.kind === 'field') {
    for (const resultChild of child.result) {
      collectRunsFromParagraphChild(resultChild, runs)
    }
  }
}

function resolveTabStops(paragraph: Paragraph): ReadonlyArray<TabStop> {
  return (
    paragraph.props?.tabs?.items.map((tab) => ({
      positionPt: twipToPt(tab.position),
      alignment: resolveTabAlignment(tab.alignment),
      leader: resolveTabLeader(tab.leader),
    })) ?? []
  )
}

function resolveTabAlignment(alignment: TabAlignment | undefined): TabStop['alignment'] {
  switch (alignment) {
    case 'center':
    case 'right':
    case 'decimal':
      return alignment
    default:
      return 'left'
  }
}

function resolveTabLeader(leader: TabLeader | undefined): TabStop['leader'] {
  switch (leader) {
    case 'dot':
    case 'hyphen':
    case 'underscore':
      return leader
    default:
      return 'none'
  }
}

/**
 * D7 / DXP-06, DXL-08, DXS-05: layers a cell's resolved table-style
 * conditional formatting under its own direct `w:tcPr` — a flat,
 * field-level merge (a direct `shd`/`tcBorders`/etc. entirely replaces the
 * conditional one rather than merging their sub-fields), matching how a
 * real table's direct formatting is understood to fully override its
 * style's conditional formatting for whichever properties it actually
 * sets.
 */
function mergeCellPropsWithConditional(
  conditional: TableCellProps,
  direct: TableCellProps | undefined,
): TableCellProps {
  return { ...conditional, ...(direct ?? {}) }
}

function resolveCellPadding(cell: TableCell, defaultCellMargins: InsetSet | undefined, cellWidthPt: number): BoxEdges {
  const cellMargins = cell.props?.tcMar
  return {
    top: resolveInsetWidth(cellMargins?.top ?? defaultCellMargins?.top, cellWidthPt),
    right: resolveInsetWidth(
      cellMargins?.right ?? cellMargins?.end ?? defaultCellMargins?.right ?? defaultCellMargins?.end,
      cellWidthPt,
    ),
    bottom: resolveInsetWidth(cellMargins?.bottom ?? defaultCellMargins?.bottom, cellWidthPt),
    left: resolveInsetWidth(
      cellMargins?.left ?? cellMargins?.start ?? defaultCellMargins?.left ?? defaultCellMargins?.start,
      cellWidthPt,
    ),
  }
}

function resolveInsetWidth(width: Width | undefined, referenceWidthPt: number): number {
  return resolveWidthToPoints(width, referenceWidthPt) ?? 0
}

function resolveWidthToPoints(width: Width | undefined, referenceWidthPt: number): number | undefined {
  if (!width || typeof width.value !== 'number') {
    return undefined
  }

  if (width.type === 'dxa') {
    return twipToPt(width.value)
  }

  if (width.type === 'pct') {
    return (clampNonNegative(referenceWidthPt) * width.value) / 5000
  }

  return undefined
}

function resolveVerticalAlign(cell: TableCell): LaidOutCell['vAlign'] {
  switch (cell.props?.vAlign) {
    case 'center':
      return 'center'
    case 'bottom':
      return 'bottom'
    default:
      return 'top'
  }
}

function resolveRowHeight(row: TableRow, naturalHeightPt: number): number {
  const explicitHeight = row.props?.trHeight
  if (!explicitHeight) {
    return naturalHeightPt
  }

  const explicitHeightPt = twipToPt(explicitHeight.val)
  if (explicitHeight.hRule === 'exact') {
    return explicitHeightPt
  }

  return Math.max(naturalHeightPt, explicitHeightPt)
}

function countRepeatHeaderRows(rows: ReadonlyArray<TableRow>): number {
  let count = 0

  for (const row of rows) {
    if (!row.props?.tblHeader) {
      break
    }

    count += 1
  }

  return count
}

function normalizeGridSpan(gridSpan: number | undefined, maxRemainingColumns?: number): number {
  const normalizedSpan = typeof gridSpan === 'number' && Number.isFinite(gridSpan) ? Math.max(1, Math.floor(gridSpan)) : 1

  if (typeof maxRemainingColumns !== 'number' || maxRemainingColumns <= 0) {
    return normalizedSpan
  }

  return Math.min(normalizedSpan, maxRemainingColumns)
}

function sumNumbers(values: ReadonlyArray<number>): number {
  return values.reduce((sum, value) => sum + value, 0)
}

function clampNonNegative(value: number): number {
  return value >= 0 ? value : 0
}

function twipToPt(value: number | undefined): number {
  return typeof value === 'number' ? value / 20 : 0
}

function isParagraph(block: Block): block is Paragraph {
  return block.kind === 'paragraph'
}

function isTableRow(row: Table['rows'][number]): row is TableRow {
  return row.kind === 'table-row'
}

function isTableCell(cell: TableRow['cells'][number]): cell is TableCell {
  return cell.kind === 'table-cell'
}

/**
 * Normalize a DOCX `<w:shd w:fill="…">` value into a CSS-compatible color
 * string. Returns `undefined` when the shading is missing, automatic, or
 * empty so the renderer can fall back to its default surface color.
 */
function resolveShadingFill(shading: Shading | undefined): string | undefined {
  if (shading === undefined) {
    return undefined
  }
  const fill = shading.fill
  if (fill === undefined || fill === 'auto') {
    return undefined
  }
  if (fill.startsWith('#')) {
    return fill
  }
  return `#${fill}`
}
