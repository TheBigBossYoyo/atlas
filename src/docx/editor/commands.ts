import type {
  Block,
  DelRevision,
  Document,
  Hyperlink,
  Indent,
  InsRevision,
  ParaProps,
  Paragraph,
  ParagraphChild,
  Run,
  RunProps,
  Section,
  Table,
  TableCell,
  TableCellProps,
  TableProps,
  TableRow,
  TextNode,
  Twip,
} from '../model'
import { twip } from '../model'

import type {
  ApplyParaFormatCommand,
  ApplyRunFormatCommand,
  Command,
  DeleteRangeCommand,
  Position,
  Range,
  TrackChangesContext,
} from './commandTypes'

type ResolvedParagraphPath = {
  readonly sectionIndex: number
  readonly blockPath: ReadonlyArray<number>
}

/**
 * DXE-03 — a run's "owner" is either the paragraph directly (a plain run) or
 * a specific `Hyperlink` wrapper it lives inside. Carrying this alongside
 * every run through the edit pipeline (split/format/delete/merge) lets those
 * operations flatten hyperlink content into the same addressable run list a
 * plain-run paragraph already used, while still reconstructing the original
 * `Hyperlink` wrapper (verbatim, minus any interior non-run children) when
 * writing the paragraph back out.
 */
/**
 * DXE-11 — a run's owner additionally tracks whether it lives inside an
 * existing `w:ins`/`w:del` (parsed from the source document, or created by a
 * previous tracked edit in this session). Flattening through these two kinds
 * the same way hyperlinks already flatten (rather than rejecting the whole
 * paragraph — `getEditableRuns`'s previous behavior for any unrecognized
 * child kind) fixes a real editing gap: a paragraph containing so much as one
 * existing tracked change used to be entirely uneditable, silently, since
 * `requireEditableRuns` threw and every caller's `event.preventDefault()`
 * safety net (D12) made that indistinguishable from "nothing happened."
 */
type RunOwner =
  | { readonly kind: 'direct' }
  | { readonly kind: 'hyperlink'; readonly wrapper: Hyperlink }
  | { readonly kind: 'ins-revision'; readonly wrapper: InsRevision }
  | { readonly kind: 'del-revision'; readonly wrapper: DelRevision }

const DIRECT_OWNER: RunOwner = Object.freeze({ kind: 'direct' })

/**
 * A freshly minted `InsRevision`/`DelRevision` wrapper only ever needs its
 * `children` field as a placeholder — `buildParagraphChildren`'s owner
 * grouping (the same mechanism that already reconstructs a `Hyperlink`
 * wrapper from its own template) immediately replaces it with the real
 * grouped runs. Typed the same way the parser's `freezeRevisionChildren`
 * already has to (`ReadonlyArray<ParagraphChild> & ReadonlyArray<Run>`, see
 * the model's own `InsRevision`/`DelRevision` field types).
 */
const EMPTY_REVISION_CHILDREN = Object.freeze([]) as ReadonlyArray<ParagraphChild> & ReadonlyArray<Run>

type RunEntry = {
  readonly run: Run
  readonly owner: RunOwner
}

type EditableRun = RunEntry & {
  readonly text: string
}

type BlocksUpdater = (
  blocks: ReadonlyArray<Block>,
  blockIndex: number,
) => ReadonlyArray<Block> | null

type RunRange = {
  readonly startOffset: number
  readonly endOffset: number
}

type RevisionResolutionCommand = Extract<
  Command,
  { kind: 'accept-revision' | 'reject-revision' }
>

type RevisionTarget = {
  readonly paragraphPath: ReadonlyArray<number>
  readonly childIndex: number
}

export function applyCommand(
  doc: Document,
  cmd: Command,
  trackChanges?: TrackChangesContext,
): {
  document: Document
  inverse: Command
  range?: Range
} {
  switch (cmd.kind) {
    case 'insert-text':
      return applyInsertText(doc, cmd, trackChanges)
    case 'delete-range':
      return applyDeleteRange(doc, cmd, trackChanges)
    case 'insert-paragraph-break':
      return applyInsertParagraphBreak(doc, cmd)
    case 'apply-run-format':
      return applyRunFormat(doc, cmd)
    case 'apply-para-format':
      return applyParaFormat(doc, cmd)
    case 'apply-style':
      return applyStyle(doc, cmd)
    case 'insert-table':
      return applyInsertTable(doc, cmd)
    case 'insert-hyperlink':
      return applyInsertHyperlink(doc, cmd)
    case 'insert-inline':
      return applyInsertInline(doc, cmd)
    case 'composite':
      return applyComposite(doc, cmd, trackChanges)
    case 'replace-blocks':
      return applyReplaceBlocks(doc, cmd)
    case 'insert-list':
      return applyInsertList(doc, cmd)
    case 'change-list-level':
      return applyChangeListLevel(doc, cmd)
    case 'accept-revision':
      return applyRevisionResolution(doc, cmd, 'accept')
    case 'reject-revision':
      return applyRevisionResolution(doc, cmd, 'reject')
    case 'accept-all-revisions':
      return applyAllRevisions(doc, 'accept')
    case 'reject-all-revisions':
      return applyAllRevisions(doc, 'reject')
    case 'insert-table-row':
      return applyInsertTableRow(doc, cmd)
    case 'delete-table-row':
      return applyDeleteTableRow(doc, cmd)
    case 'insert-table-column':
      return applyInsertTableColumn(doc, cmd)
    case 'delete-table-column':
      return applyDeleteTableColumn(doc, cmd)
    case 'delete-table':
      return applyDeleteTable(doc, cmd)
    case 'merge-table-cells':
      return applyMergeTableCells(doc, cmd)
    case 'split-table-cell':
      return applySplitTableCell(doc, cmd)
    case 'resize-table-column':
      return applyResizeTableColumn(doc, cmd)
    case 'replace-table':
      return applyReplaceTable(doc, cmd)
    case 'apply-table-props':
      return applyTableProps(doc, cmd)
  }
}

// ---------------------------------------------------------------------------
// Composite / structural primitives (D13, D12 cross-paragraph support)
// ---------------------------------------------------------------------------

/**
 * DXE-11 — `trackChanges`, when given, is forwarded to every sub-command
 * exactly as `applyCommand` would for a lone command: an `insert-text`/
 * `delete-range` inside a batch (e.g. `handleRichPaste`'s pasted-text
 * commands, or Replace All's find-and-replace pairs) is recorded as
 * `w:ins`/`w:del` the same as typed text, while a sub-command kind that
 * doesn't accept tracking (`insert-table`, `apply-run-format`, ...) just
 * ignores the extra argument as it always has — see those functions' own
 * signatures for which kinds actually consult it.
 */
function applyComposite(
  doc: Document,
  cmd: Extract<Command, { kind: 'composite' }>,
  trackChanges?: TrackChangesContext,
): { document: Document; inverse: Command; range?: Range } {
  let workingDocument = doc
  const inverses: Command[] = []
  let lastRange: Range | undefined

  for (const sub of cmd.commands) {
    const result = applyCommand(workingDocument, sub, trackChanges)
    workingDocument = result.document
    inverses.push(result.inverse)
    if (result.range !== undefined) {
      lastRange = result.range
    }
  }

  inverses.reverse()

  return {
    document: workingDocument,
    inverse: inverses.length === 1 ? inverses[0] : { kind: 'composite', commands: inverses },
    ...(lastRange !== undefined ? { range: lastRange } : {}),
  }
}

function applyReplaceBlocks(
  doc: Document,
  cmd: Extract<Command, { kind: 'replace-blocks' }>,
): { document: Document; inverse: Command; range?: Range } {
  const resolvedPath = resolveParagraphPath(doc, cmd.at)
  if (resolvedPath === null || resolvedPath.blockPath.length === 0) {
    throw new Error('ReplaceBlocks target not found')
  }

  const section = doc.sections[resolvedPath.sectionIndex]
  if (section === undefined) {
    throw new Error('ReplaceBlocks target not found')
  }

  const prefix = resolvedPath.blockPath.slice(0, -1)
  const startIndex = resolvedPath.blockPath[resolvedPath.blockPath.length - 1]
  const siblingBlocks = getBlocksAtPrefix(section.blocks, prefix)

  if (
    siblingBlocks === null ||
    startIndex < 0 ||
    cmd.count < 0 ||
    startIndex + cmd.count > siblingBlocks.length
  ) {
    throw new Error('ReplaceBlocks target is out of range')
  }

  const removed = siblingBlocks.slice(startIndex, startIndex + cmd.count)
  const nextSiblingBlocks = freezeArray([
    ...siblingBlocks.slice(0, startIndex),
    ...cmd.blocks,
    ...siblingBlocks.slice(startIndex + cmd.count),
  ])
  const nextDocument = setBlocksAtPrefix(doc, resolvedPath.sectionIndex, prefix, nextSiblingBlocks)

  return {
    document: nextDocument,
    inverse: {
      kind: 'replace-blocks',
      at: clonePath(cmd.at),
      count: cmd.blocks.length,
      blocks: removed,
    },
    ...(cmd.cursor !== undefined ? { range: cmd.cursor } : {}),
  }
}

function applyInsertInline(
  doc: Document,
  cmd: Extract<Command, { kind: 'insert-inline' }>,
): { document: Document; inverse: Command; range: Range } {
  const paragraph = requireParagraph(doc, cmd.at.paragraphPath)
  const editableRuns = requireEditableRuns(paragraph)
  const split = splitEntriesAtPosition(editableRuns, cmd.at)

  const inlineRun: Run = Object.freeze({ kind: 'run', children: freezeArray([cmd.child]) })
  const inlineEntry: RunEntry = { run: inlineRun, owner: DIRECT_OWNER }

  const nextEntries = [...split.beforeEntries, inlineEntry, ...split.afterEntries]
  const nextParagraph = cloneParagraph(paragraph, buildParagraphChildren(nextEntries))
  const nextDocument = replaceParagraphOrThrow(doc, cmd.at.paragraphPath, nextParagraph)

  const resolvedPath = resolveParagraphPath(doc, cmd.at.paragraphPath)
  if (resolvedPath === null) {
    throw new Error('Paragraph not found')
  }

  const cursor = createPosition(cmd.at.paragraphPath, split.beforeEntries.length + 1, 0)

  return {
    document: nextDocument,
    inverse: {
      kind: 'replace-blocks',
      at: freezeArray([resolvedPath.sectionIndex, ...resolvedPath.blockPath]),
      count: 1,
      blocks: [paragraph],
      cursor: { anchor: cmd.at, focus: cmd.at },
    },
    range: { anchor: cursor, focus: cursor },
  }
}

function applyInsertHyperlink(
  doc: Document,
  cmd: Extract<Command, { kind: 'insert-hyperlink' }>,
): { document: Document; inverse: Command; range: Range } {
  const range = normalizeRange(doc, cmd.range)
  if (!sameParagraphPath(doc, range.anchor.paragraphPath, range.focus.paragraphPath)) {
    throw new Error('InsertHyperlink across multiple paragraphs is not yet implemented')
  }

  const paragraphPath = range.anchor.paragraphPath
  const paragraph = requireParagraph(doc, paragraphPath)
  const editableRuns = requireEditableRuns(paragraph)
  const resolvedPath = resolveParagraphPath(doc, paragraphPath)
  if (resolvedPath === null) {
    throw new Error('Paragraph not found')
  }

  const wrapper: Hyperlink = Object.freeze({
    kind: 'hyperlink',
    relationshipId: cmd.relationshipId,
    history: true,
    children: freezeArray<Run>([]),
  })

  let nextEntries: ReadonlyArray<RunEntry>
  let selectionEnd: Position

  if (comparePositions(doc, range.anchor, range.focus) === 0) {
    const target = resolveInsertTarget(editableRuns, range.anchor)
    const targetPosition = createPosition(paragraphPath, target.runIndex, target.charOffset)
    const split = splitEntriesAtPosition(editableRuns, targetPosition)
    const linkRun: Run = Object.freeze({ kind: 'run', children: freezeArray([createTextNode(cmd.url)]) })

    nextEntries = freezeArray([
      ...split.beforeEntries,
      { run: linkRun, owner: { kind: 'hyperlink', wrapper } },
      ...split.afterEntries,
    ])
    selectionEnd = createPosition(paragraphPath, split.beforeEntries.length, cmd.url.length)
  } else {
    const offsets = getRangeOffsets(paragraph, range)
    const sliced = sliceEntriesByOffsets(editableRuns, offsets)
    if (sliced.within.length === 0) {
      throw new Error('InsertHyperlink selection is empty')
    }

    const wrapped = sliced.within.map((entry): RunEntry => ({ run: entry.run, owner: { kind: 'hyperlink', wrapper } }))
    nextEntries = freezeArray([...sliced.before, ...wrapped, ...sliced.after])
    selectionEnd = createPosition(paragraphPath, sliced.before.length + wrapped.length, 0)
  }

  const nextChildren = buildParagraphChildren(nextEntries)
  const nextParagraph = cloneParagraph(paragraph, nextChildren)
  const nextDocument = replaceParagraphOrThrow(doc, paragraphPath, nextParagraph)

  const inverseAt = freezeArray([resolvedPath.sectionIndex, ...resolvedPath.blockPath])

  return {
    document: nextDocument,
    inverse: {
      kind: 'replace-blocks',
      at: inverseAt,
      count: 1,
      blocks: [paragraph],
      cursor: { anchor: range.anchor, focus: range.focus },
    },
    range: { anchor: selectionEnd, focus: selectionEnd },
  }
}

/** DXE-19 — builds the default empty `rows` x `cols` grid `applyInsertTable`
 * uses when the caller (the toolbar's table-size picker) doesn't supply a
 * fully-built `table` of its own (paste fidelity's own content-bearing
 * table, inserted verbatim instead — see `InsertTableCommand`'s doc
 * comment). */
function buildEmptyTable(rows: number, cols: number): Table {
  const TOTAL_WIDTH_TWIPS = 9000
  const columnWidth = twip(Math.max(1, Math.floor(TOTAL_WIDTH_TWIPS / cols)))
  const tblGrid = freezeArray(Array.from({ length: cols }, () => columnWidth))

  const makeCell = (): TableCell =>
    Object.freeze({ kind: 'table-cell', blocks: freezeArray<Block>([emptyParagraph()]) })
  const makeRow = (): TableRow =>
    Object.freeze({ kind: 'table-row', cells: freezeArray(Array.from({ length: cols }, makeCell)) })

  return Object.freeze({
    kind: 'table',
    tblGrid,
    rows: freezeArray(Array.from({ length: rows }, makeRow)),
  })
}

function applyInsertTable(
  doc: Document,
  cmd: Extract<Command, { kind: 'insert-table' }>,
): { document: Document; inverse: Command; range: Range } {
  if (cmd.table === undefined && (cmd.rows < 1 || cmd.cols < 1)) {
    throw new Error('InsertTable requires at least one row and one column')
  }
  if (cmd.table !== undefined && cmd.table.rows.length < 1) {
    throw new Error('InsertTable requires at least one row and one column')
  }

  const paragraph = requireParagraph(doc, cmd.at.paragraphPath)
  const editableRuns = requireEditableRuns(paragraph)
  const split = splitEntriesAtPosition(editableRuns, cmd.at)

  const table: Table = cmd.table ?? buildEmptyTable(cmd.rows, cmd.cols)

  const beforeParagraph = cloneParagraph(paragraph, buildParagraphChildren(split.beforeEntries))
  const afterParagraph = cloneParagraph(paragraph, buildParagraphChildren(split.afterEntries))

  const resolvedPath = resolveParagraphPath(doc, cmd.at.paragraphPath)
  if (resolvedPath === null || resolvedPath.blockPath.length === 0) {
    throw new Error('Paragraph not found')
  }

  const section = doc.sections[resolvedPath.sectionIndex]
  if (section === undefined) {
    throw new Error('Paragraph not found')
  }

  const prefix = resolvedPath.blockPath.slice(0, -1)
  const index = resolvedPath.blockPath[resolvedPath.blockPath.length - 1]
  const siblingBlocks = getBlocksAtPrefix(section.blocks, prefix)
  if (siblingBlocks === null) {
    throw new Error('Paragraph not found')
  }

  const nextSiblingBlocks = freezeArray([
    ...siblingBlocks.slice(0, index),
    beforeParagraph,
    table,
    afterParagraph,
    ...siblingBlocks.slice(index + 1),
  ])
  const nextDocument = setBlocksAtPrefix(doc, resolvedPath.sectionIndex, prefix, nextSiblingBlocks)

  const inverseAt = freezeArray([resolvedPath.sectionIndex, ...prefix, index])
  const firstCellPath = freezeArray([resolvedPath.sectionIndex, ...prefix, index + 1, 0, 0, 0])
  const cursor = createPosition(firstCellPath, 0, 0)

  return {
    document: nextDocument,
    inverse: {
      kind: 'replace-blocks',
      at: inverseAt,
      count: 3,
      blocks: [paragraph],
      cursor: { anchor: cmd.at, focus: cmd.at },
    },
    range: { anchor: cursor, focus: cursor },
  }
}

// ---------------------------------------------------------------------------
// Table structural editing (DXE-14) — see commandTypes.ts's module doc
// comment above the command shapes for the addressing/inverse design.
// ---------------------------------------------------------------------------

const DEFAULT_COLUMN_WIDTH_TWIPS = 1440
const MIN_COLUMN_WIDTH_TWIPS = 180

function clamp(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value))
}

function requireTableAt(doc: Document, tablePath: ReadonlyArray<number>): Table {
  const resolved = resolveParagraphPath(doc, tablePath)
  if (resolved === null || resolved.blockPath.length === 0) {
    throw new Error('Table not found')
  }

  const section = doc.sections[resolved.sectionIndex]
  if (section === undefined) {
    throw new Error('Table not found')
  }

  const prefix = resolved.blockPath.slice(0, -1)
  const index = resolved.blockPath[resolved.blockPath.length - 1]
  const siblingBlocks = getBlocksAtPrefix(section.blocks, prefix)
  const block = siblingBlocks?.[index]
  if (block === undefined || block.kind !== 'table') {
    throw new Error('Table not found')
  }

  return block
}

/** Replaces the table at `tablePath` with whatever `transform` returns (or
 * removes it entirely when `transform` returns `null`), reusing the same
 * generic block-path walk `updateDocumentAtParagraphPath` already uses for
 * paragraph edits — it works unchanged for a table nested inside a cell. */
function updateTable(
  doc: Document,
  tablePath: ReadonlyArray<number>,
  transform: (table: Table) => Table | null,
): Document {
  const nextDocument = updateDocumentAtParagraphPath(doc, tablePath, (blocks, blockIndex) => {
    const table = blocks[blockIndex]
    if (table === undefined || table.kind !== 'table') {
      return null
    }

    const nextTable = transform(table)
    if (nextTable === null) {
      return freezeArray([...blocks.slice(0, blockIndex), ...blocks.slice(blockIndex + 1)])
    }

    return replaceArrayItem(blocks, blockIndex, nextTable)
  })

  if (nextDocument === null) {
    throw new Error('Table not found')
  }

  return nextDocument
}

function cloneTableFull(
  table: Table,
  rows: ReadonlyArray<Table['rows'][number]>,
  tblGrid: ReadonlyArray<Twip> | undefined,
): Table {
  return Object.freeze({
    kind: 'table',
    ...(table.props !== undefined ? { props: table.props } : {}),
    ...(tblGrid !== undefined ? { tblGrid } : {}),
    rows,
  })
}

function createEmptyTableCell(): TableCell {
  return Object.freeze({ kind: 'table-cell', blocks: freezeArray<Block>([emptyParagraph()]) })
}

function createEmptyTableRow(columnCount: number): TableRow {
  return Object.freeze({
    kind: 'table-row',
    cells: freezeArray(Array.from({ length: columnCount }, createEmptyTableCell)),
  })
}

/** The number of grid columns a table has — `tblGrid` when present (the
 * common, schema-required case), otherwise the first row's total cell
 * `gridSpan`. */
function tableColumnCount(table: Table): number {
  if (table.tblGrid !== undefined) {
    return table.tblGrid.length
  }

  const firstRow = table.rows.find((row): row is TableRow => row.kind === 'table-row')
  if (firstRow === undefined) {
    return 1
  }

  return firstRow.cells.reduce(
    (sum, cell) => sum + (cell.kind === 'table-cell' ? cell.props?.gridSpan ?? 1 : 1),
    0,
  )
}

function insertGridColumn(
  tblGrid: ReadonlyArray<Twip> | undefined,
  at: number,
  width: Twip,
): ReadonlyArray<Twip> | undefined {
  if (tblGrid === undefined) {
    return undefined
  }

  return freezeArray([...tblGrid.slice(0, at), width, ...tblGrid.slice(at)])
}

function deleteGridColumn(
  tblGrid: ReadonlyArray<Twip> | undefined,
  columnIndex: number,
): ReadonlyArray<Twip> | undefined {
  if (tblGrid === undefined) {
    return undefined
  }

  return freezeArray([...tblGrid.slice(0, columnIndex), ...tblGrid.slice(columnIndex + 1)])
}

function withTableCellProps(cell: TableCell, props: TableCellProps | undefined): TableCell {
  return Object.freeze({
    kind: 'table-cell',
    ...(props !== undefined ? { props } : {}),
    blocks: cell.blocks,
  })
}

/**
 * Inserts a new cell before grid column `columnIndex`, walking the row's
 * cells while tracking accumulated `gridSpan` so the index is interpreted in
 * *grid-column* space, not cell-array space. `columnIndex` landing strictly
 * inside an existing merged cell's span widens that cell by one column
 * instead (matching Word: inserting a column through a merged header cell
 * grows the merge rather than splitting it); landing exactly on a cell
 * boundary — including at the very start or end of the row — inserts a new
 * standalone cell there.
 */
function insertColumnIntoRow(row: TableRow, columnIndex: number, makeCell: () => TableCell): TableRow {
  const nextCells: Array<TableRow['cells'][number]> = []
  let accumulated = 0
  let inserted = false

  for (const cell of row.cells) {
    if (!inserted && columnIndex === accumulated) {
      nextCells.push(makeCell())
      inserted = true
    }

    if (cell.kind !== 'table-cell') {
      nextCells.push(cell)
      accumulated += 1
      continue
    }

    const span = cell.props?.gridSpan ?? 1
    if (!inserted && columnIndex > accumulated && columnIndex < accumulated + span) {
      nextCells.push(withTableCellProps(cell, { ...cell.props, gridSpan: span + 1 }))
      inserted = true
      accumulated += span
      continue
    }

    nextCells.push(cell)
    accumulated += span
  }

  if (!inserted) {
    nextCells.push(makeCell())
  }

  return cloneTableRow(row, freezeArray(nextCells))
}

/**
 * Removes grid column `columnIndex` from a row: a cell exactly one column
 * wide covering it is dropped entirely; a merged cell spanning it is
 * narrowed by one column instead (the mirror image of
 * `insertColumnIntoRow`'s widen case).
 */
function deleteColumnFromRow(row: TableRow, columnIndex: number): TableRow {
  const nextCells: Array<TableRow['cells'][number]> = []
  let accumulated = 0
  let handled = false

  for (const cell of row.cells) {
    if (cell.kind !== 'table-cell') {
      nextCells.push(cell)
      accumulated += 1
      continue
    }

    const span = cell.props?.gridSpan ?? 1
    if (!handled && columnIndex >= accumulated && columnIndex < accumulated + span) {
      handled = true
      if (span > 1) {
        nextCells.push(withTableCellProps(cell, { ...cell.props, gridSpan: span - 1 }))
      }
      accumulated += span
      continue
    }

    nextCells.push(cell)
    accumulated += span
  }

  return cloneTableRow(row, freezeArray(nextCells))
}

function resizeColumnInRow(row: TableRow, columnIndex: number, width: Twip): TableRow {
  let accumulated = 0
  const nextCells = row.cells.map((cell) => {
    if (cell.kind !== 'table-cell') {
      accumulated += 1
      return cell
    }

    const span = cell.props?.gridSpan ?? 1
    const covers = columnIndex >= accumulated && columnIndex < accumulated + span
    accumulated += span

    // A merged cell's own width is the sum of every grid column it spans;
    // resizing one of those columns without re-deriving every affected
    // cell's total width is a documented gap (see DXE-14's remaining-work
    // note) — only an unspanned (single-column) cell's `tcW` is updated here.
    if (!covers || span !== 1) {
      return cell
    }

    return withTableCellProps(cell, { ...cell.props, tcW: { type: 'dxa', value: width } })
  })

  return cloneTableRow(row, freezeArray(nextCells))
}

function evenlyDistributeSpan(total: number, into: number): ReadonlyArray<number> {
  const base = Math.floor(total / into)
  const remainder = total % into
  return Array.from({ length: into }, (_, index) => base + (index < remainder ? 1 : 0))
}

function withoutGridSpan(props: TableCellProps | undefined): TableCellProps | undefined {
  if (props === undefined || props.gridSpan === undefined) {
    return props
  }

  const rest = Object.fromEntries(
    Object.entries(props).filter(([key]) => key !== 'gridSpan'),
  ) as TableCellProps
  return Object.keys(rest).length > 0 ? rest : undefined
}

function applyInsertTableRow(
  doc: Document,
  cmd: Extract<Command, { kind: 'insert-table-row' }>,
): { document: Document; inverse: Command } {
  const table = requireTableAt(doc, cmd.tablePath)
  const at = clamp(cmd.at, 0, table.rows.length)
  const newRow = createEmptyTableRow(tableColumnCount(table))

  const nextDocument = updateTable(doc, cmd.tablePath, (t) =>
    cloneTableFull(t, freezeArray([...t.rows.slice(0, at), newRow, ...t.rows.slice(at)]), t.tblGrid),
  )

  return {
    document: nextDocument,
    inverse: { kind: 'delete-table-row', tablePath: cmd.tablePath, rowIndex: at },
  }
}

function applyDeleteTableRow(
  doc: Document,
  cmd: Extract<Command, { kind: 'delete-table-row' }>,
): { document: Document; inverse: Command } {
  const table = requireTableAt(doc, cmd.tablePath)
  if (cmd.rowIndex < 0 || cmd.rowIndex >= table.rows.length) {
    throw new Error('DeleteTableRow row index is out of range')
  }
  if (table.rows.length <= 1) {
    throw new Error('Cannot delete a table\'s last row — delete the table instead')
  }

  const nextDocument = updateTable(doc, cmd.tablePath, (t) =>
    cloneTableFull(
      t,
      freezeArray([...t.rows.slice(0, cmd.rowIndex), ...t.rows.slice(cmd.rowIndex + 1)]),
      t.tblGrid,
    ),
  )

  return {
    document: nextDocument,
    inverse: { kind: 'replace-table', tablePath: cmd.tablePath, table },
  }
}

function applyInsertTableColumn(
  doc: Document,
  cmd: Extract<Command, { kind: 'insert-table-column' }>,
): { document: Document; inverse: Command } {
  const table = requireTableAt(doc, cmd.tablePath)
  const at = clamp(cmd.at, 0, tableColumnCount(table))
  const width = twip(cmd.widthTwips ?? DEFAULT_COLUMN_WIDTH_TWIPS)

  const nextDocument = updateTable(doc, cmd.tablePath, (t) => {
    const nextRows = t.rows.map((row) =>
      row.kind === 'table-row' ? insertColumnIntoRow(row, at, createEmptyTableCell) : row,
    )
    return cloneTableFull(t, freezeArray(nextRows), insertGridColumn(t.tblGrid, at, width))
  })

  return {
    document: nextDocument,
    inverse: { kind: 'delete-table-column', tablePath: cmd.tablePath, columnIndex: at },
  }
}

function applyDeleteTableColumn(
  doc: Document,
  cmd: Extract<Command, { kind: 'delete-table-column' }>,
): { document: Document; inverse: Command } {
  const table = requireTableAt(doc, cmd.tablePath)
  const columnCount = tableColumnCount(table)
  if (cmd.columnIndex < 0 || cmd.columnIndex >= columnCount) {
    throw new Error('DeleteTableColumn column index is out of range')
  }
  if (columnCount <= 1) {
    throw new Error('Cannot delete a table\'s last column — delete the table instead')
  }

  const nextDocument = updateTable(doc, cmd.tablePath, (t) => {
    const nextRows = t.rows.map((row) =>
      row.kind === 'table-row' ? deleteColumnFromRow(row, cmd.columnIndex) : row,
    )
    return cloneTableFull(t, freezeArray(nextRows), deleteGridColumn(t.tblGrid, cmd.columnIndex))
  })

  return {
    document: nextDocument,
    inverse: { kind: 'replace-table', tablePath: cmd.tablePath, table },
  }
}

/** Deletes the whole table — a thin, table-specific entry point over the
 * same general `replace-blocks` primitive cross-paragraph delete already
 * uses, so its inverse (re-inserting the exact table verbatim) comes for
 * free rather than needing its own command kind. */
function applyDeleteTable(
  doc: Document,
  cmd: Extract<Command, { kind: 'delete-table' }>,
): { document: Document; inverse: Command; range?: Range } {
  return applyReplaceBlocks(doc, { kind: 'replace-blocks', at: cmd.tablePath, count: 1, blocks: [] })
}

/**
 * Merges cells `fromCellIndex..toCellIndex` (inclusive) of one row into a
 * single cell: `gridSpan` becomes the sum of the merged cells' spans, and
 * the resulting cell's content is every merged cell's paragraphs
 * concatenated in order (Word has no more principled way to combine them
 * either). Horizontal-only — see commandTypes.ts's module doc comment for
 * why a vertical (`vMerge`, across rows) merge is a documented follow-up
 * rather than implemented here.
 */
function applyMergeTableCells(
  doc: Document,
  cmd: Extract<Command, { kind: 'merge-table-cells' }>,
): { document: Document; inverse: Command } {
  const table = requireTableAt(doc, cmd.tablePath)
  const row = table.rows[cmd.rowIndex]
  if (row === undefined || row.kind !== 'table-row') {
    throw new Error('MergeTableCells target row not found')
  }
  if (
    cmd.fromCellIndex < 0 ||
    cmd.toCellIndex >= row.cells.length ||
    cmd.fromCellIndex >= cmd.toCellIndex
  ) {
    throw new Error('MergeTableCells requires at least two cells in range')
  }

  const targetCells = row.cells.slice(cmd.fromCellIndex, cmd.toCellIndex + 1)
  if (targetCells.some((cell) => cell.kind !== 'table-cell')) {
    throw new Error('MergeTableCells range includes an unsupported cell')
  }
  const cells = targetCells as ReadonlyArray<TableCell>

  const totalSpan = cells.reduce((sum, cell) => sum + (cell.props?.gridSpan ?? 1), 0)
  const mergedCell: TableCell = Object.freeze({
    kind: 'table-cell',
    props: { ...cells[0].props, gridSpan: totalSpan },
    blocks: freezeArray(cells.flatMap((cell) => cell.blocks)),
  })

  const nextDocument = updateTable(doc, cmd.tablePath, (t) => {
    const nextRow = cloneTableRow(
      row,
      freezeArray([
        ...row.cells.slice(0, cmd.fromCellIndex),
        mergedCell,
        ...row.cells.slice(cmd.toCellIndex + 1),
      ]),
    )
    return cloneTableFull(t, replaceArrayItem(t.rows, cmd.rowIndex, nextRow), t.tblGrid)
  })

  return {
    document: nextDocument,
    inverse: { kind: 'replace-table', tablePath: cmd.tablePath, table },
  }
}

/**
 * Splits a merged cell back into `into` (default: its full `gridSpan`)
 * side-by-side cells, distributing the original span as evenly as possible.
 * All of the original cell's content goes into the first resulting cell —
 * matching Word, which has no principled way to redistribute paragraphs
 * across the new cells either — leaving the rest empty.
 */
function applySplitTableCell(
  doc: Document,
  cmd: Extract<Command, { kind: 'split-table-cell' }>,
): { document: Document; inverse: Command } {
  const table = requireTableAt(doc, cmd.tablePath)
  const row = table.rows[cmd.rowIndex]
  if (row === undefined || row.kind !== 'table-row') {
    throw new Error('SplitTableCell target row not found')
  }
  const cell = row.cells[cmd.cellIndex]
  if (cell === undefined || cell.kind !== 'table-cell') {
    throw new Error('SplitTableCell target cell not found')
  }

  const span = cell.props?.gridSpan ?? 1
  const into = cmd.into ?? span
  if (into < 2 || into > span) {
    throw new Error("SplitTableCell must split into 2..the cell's current column span")
  }

  const spans = evenlyDistributeSpan(span, into)
  const newCells: TableCell[] = spans.map((cellSpan, index) => {
    const baseProps = index === 0 ? cell.props : undefined
    const props = cellSpan > 1 ? { ...baseProps, gridSpan: cellSpan } : withoutGridSpan(baseProps)
    return withTableCellProps(
      { kind: 'table-cell', blocks: index === 0 ? cell.blocks : freezeArray<Block>([emptyParagraph()]) },
      props,
    )
  })

  const nextDocument = updateTable(doc, cmd.tablePath, (t) => {
    const nextRow = cloneTableRow(
      row,
      freezeArray([...row.cells.slice(0, cmd.cellIndex), ...newCells, ...row.cells.slice(cmd.cellIndex + 1)]),
    )
    return cloneTableFull(t, replaceArrayItem(t.rows, cmd.rowIndex, nextRow), t.tblGrid)
  })

  return {
    document: nextDocument,
    inverse: { kind: 'replace-table', tablePath: cmd.tablePath, table },
  }
}

function applyResizeTableColumn(
  doc: Document,
  cmd: Extract<Command, { kind: 'resize-table-column' }>,
): { document: Document; inverse: Command } {
  const table = requireTableAt(doc, cmd.tablePath)
  const columnCount = tableColumnCount(table)
  if (cmd.columnIndex < 0 || cmd.columnIndex >= columnCount) {
    throw new Error('ResizeTableColumn column index is out of range')
  }
  if (cmd.widthTwips < MIN_COLUMN_WIDTH_TWIPS) {
    throw new Error('ResizeTableColumn width is too small')
  }

  const width = twip(cmd.widthTwips)
  const nextDocument = updateTable(doc, cmd.tablePath, (t) => {
    const nextGrid =
      t.tblGrid !== undefined
        ? freezeArray(t.tblGrid.map((w, index) => (index === cmd.columnIndex ? width : w)))
        : t.tblGrid
    const nextRows = t.rows.map((row) =>
      row.kind === 'table-row' ? resizeColumnInRow(row, cmd.columnIndex, width) : row,
    )
    return cloneTableFull(t, freezeArray(nextRows), nextGrid)
  })

  return {
    document: nextDocument,
    inverse: { kind: 'replace-table', tablePath: cmd.tablePath, table },
  }
}

function applyReplaceTable(
  doc: Document,
  cmd: Extract<Command, { kind: 'replace-table' }>,
): { document: Document; inverse: Command } {
  const original = requireTableAt(doc, cmd.tablePath)
  const nextDocument = updateTable(doc, cmd.tablePath, () => cmd.table)

  return {
    document: nextDocument,
    inverse: { kind: 'replace-table', tablePath: cmd.tablePath, table: original },
  }
}

/** See `commandTypes.ts`'s `ApplyTablePropsCommand` doc comment for why this
 * replaces `props` wholesale rather than merging like the run/paragraph
 * format commands do. */
function withTableProps(table: Table, props: TableProps | undefined): Table {
  return Object.freeze({
    kind: 'table',
    ...(props !== undefined ? { props } : {}),
    ...(table.tblGrid !== undefined ? { tblGrid: table.tblGrid } : {}),
    rows: table.rows,
  })
}

function applyTableProps(
  doc: Document,
  cmd: Extract<Command, { kind: 'apply-table-props' }>,
): { document: Document; inverse: Command } {
  const original = requireTableAt(doc, cmd.tablePath)
  const nextDocument = updateTable(doc, cmd.tablePath, (t) => withTableProps(t, cmd.props))

  return {
    document: nextDocument,
    inverse: { kind: 'apply-table-props', tablePath: cmd.tablePath, props: original.props },
  }
}

// ---------------------------------------------------------------------------
// Lists / indent (D18)
// ---------------------------------------------------------------------------

const MAX_LIST_LEVEL = 8
const INDENT_STEP_TWIPS = 720

function applyInsertList(
  doc: Document,
  cmd: Extract<Command, { kind: 'insert-list' }>,
): { document: Document; inverse: Command } {
  if (cmd.paragraphPaths.length === 0) {
    return { document: doc, inverse: { kind: 'insert-list', paragraphPaths: [], numId: cmd.numId, level: cmd.level } }
  }

  const numIdStr = String(cmd.numId)
  const allAlreadyListed = cmd.paragraphPaths.every((path) => {
    const paragraph = findParagraph(doc, path)
    return paragraph?.props?.numPr?.numId === numIdStr
  })

  const subCommands: Command[] = cmd.paragraphPaths.map((path) => ({
    kind: 'apply-para-format',
    paragraphPaths: [clonePath(path)],
    format: allAlreadyListed
      ? { numPr: undefined }
      : { numPr: { numId: numIdStr, ilvl: cmd.level } },
  }))

  const result = applyComposite(doc, { kind: 'composite', commands: subCommands })
  return { document: result.document, inverse: result.inverse }
}

function applyChangeListLevel(
  doc: Document,
  cmd: Extract<Command, { kind: 'change-list-level' }>,
): { document: Document; inverse: Command } {
  const paragraph = requireParagraph(doc, cmd.paragraphPath)
  const numPr = paragraph.props?.numPr

  if (numPr?.numId !== undefined) {
    const currentLevel = numPr.ilvl ?? 0
    const nextLevel = Math.min(MAX_LIST_LEVEL, Math.max(0, currentLevel + cmd.delta))
    if (nextLevel === currentLevel) {
      return { document: doc, inverse: noOpParaFormat(cmd.paragraphPath) }
    }

    const nextParagraph = cloneParagraph(
      paragraph,
      paragraph.children,
      applyPropsPatch(paragraph.props, { numPr: { ...numPr, ilvl: nextLevel } }),
    )
    const nextDocument = replaceParagraphOrThrow(doc, cmd.paragraphPath, nextParagraph)

    return {
      document: nextDocument,
      inverse: {
        kind: 'change-list-level',
        paragraphPath: clonePath(cmd.paragraphPath),
        delta: (cmd.delta * -1) as 1 | -1,
      },
    }
  }

  const currentIndent = paragraph.props?.ind?.left ?? 0
  const nextIndent = Math.max(0, currentIndent + cmd.delta * INDENT_STEP_TWIPS)
  if (nextIndent === currentIndent) {
    return { document: doc, inverse: noOpParaFormat(cmd.paragraphPath) }
  }

  const nextParagraph = cloneParagraph(
    paragraph,
    paragraph.children,
    applyPropsPatch(paragraph.props, { ind: withIndentLeft(paragraph.props?.ind, nextIndent) }),
  )
  const nextDocument = replaceParagraphOrThrow(doc, cmd.paragraphPath, nextParagraph)

  return {
    document: nextDocument,
    inverse: {
      kind: 'change-list-level',
      paragraphPath: clonePath(cmd.paragraphPath),
      delta: (cmd.delta * -1) as 1 | -1,
    },
  }
}

/**
 * Sets `ind.left`, but drops the key entirely (returning `undefined`, or the
 * surviving sibling `ind` fields if any) once it reaches zero, so outdenting
 * all the way back to a paragraph's original state doesn't leave a stray
 * `ind: { left: 0 }` behind — `applyPropsPatch` treats an `undefined` value as
 * "delete this key", which is what makes ChangeListLevel's own inverse an
 * exact round-trip back to a paragraph that never had an indent at all.
 */
function withIndentLeft(ind: Indent | undefined, leftTwips: number): Indent | undefined {
  if (leftTwips <= 0) {
    if (ind === undefined) {
      return undefined
    }

    const rest = Object.fromEntries(Object.entries(ind).filter(([key]) => key !== 'left'))
    return Object.keys(rest).length > 0 ? (rest as Indent) : undefined
  }

  return { ...ind, left: twip(leftTwips) }
}

function noOpParaFormat(paragraphPath: ReadonlyArray<number>): Command {
  return { kind: 'apply-para-format', paragraphPaths: [clonePath(paragraphPath)], format: {} }
}

// ---------------------------------------------------------------------------
// Revisions (unchanged)
// ---------------------------------------------------------------------------

function applyRevisionResolution(
  doc: Document,
  cmd: RevisionResolutionCommand,
  mode: 'accept' | 'reject',
): {
  document: Document
  inverse: Command
} {
  const target = resolveRevisionTarget(doc, cmd)
  const paragraph = requireParagraph(doc, target.paragraphPath)
  const child = paragraph.children[target.childIndex]
  if (child === undefined || (child.kind !== 'ins-revision' && child.kind !== 'del-revision')) {
    throw new Error('Revision child not found at index')
  }

  const nextChildren = resolveRevisionChild(paragraph.children, target.childIndex, child, mode)
  const nextParagraph = cloneParagraph(paragraph, nextChildren)
  const nextDocument = replaceParagraphOrThrow(doc, target.paragraphPath, nextParagraph)

  return {
    document: nextDocument,
    inverse: createNoOpInsertText(createPosition(target.paragraphPath, 0, 0)),
  }
}

function resolveRevisionChild(
  children: ReadonlyArray<ParagraphChild>,
  childIndex: number,
  revision: Extract<ParagraphChild, { kind: 'ins-revision' | 'del-revision' }>,
  mode: 'accept' | 'reject',
): ReadonlyArray<ParagraphChild> {
  const keep =
    (revision.kind === 'ins-revision' && mode === 'accept') ||
    (revision.kind === 'del-revision' && mode === 'reject')

  const replacement: ReadonlyArray<ParagraphChild> = keep ? getRevisionChildren(revision) : []
  return freezeArray([
    ...children.slice(0, childIndex),
    ...replacement,
    ...children.slice(childIndex + 1),
  ])
}

function resolveRevisionTarget(
  doc: Document,
  cmd: RevisionResolutionCommand,
): RevisionTarget {
  if ('id' in cmd) {
    const resolved = findRevisionTargetById(doc, cmd.id)
    if (resolved === null) {
      throw new Error(`Revision not found for id ${cmd.id}`)
    }
    return resolved
  }

  return {
    paragraphPath: clonePath(cmd.paragraphPath),
    childIndex: cmd.childIndex,
  }
}

function findRevisionTargetById(
  doc: Document,
  id: string,
): RevisionTarget | null {
  for (let sectionIndex = 0; sectionIndex < doc.sections.length; sectionIndex += 1) {
    const section = doc.sections[sectionIndex]
    if (section === undefined) {
      continue
    }

    const target = findRevisionTargetInBlocks(section.blocks, [sectionIndex], id)
    if (target !== null) {
      return target
    }
  }

  return null
}

function findRevisionTargetInBlocks(
  blocks: ReadonlyArray<Block>,
  pathPrefix: ReadonlyArray<number>,
  id: string,
): RevisionTarget | null {
  for (let blockIndex = 0; blockIndex < blocks.length; blockIndex += 1) {
    const block = blocks[blockIndex]
    if (block === undefined) {
      continue
    }

    const blockPath = [...pathPrefix, blockIndex]
    if (block.kind === 'paragraph') {
      const childIndex = block.children.findIndex(
        (child) =>
          (child.kind === 'ins-revision' || child.kind === 'del-revision') &&
          child.id === id,
      )

      if (childIndex >= 0) {
        return {
          paragraphPath: freezeArray(blockPath),
          childIndex,
        }
      }

      continue
    }

    if (block.kind !== 'table') {
      continue
    }

    for (let rowIndex = 0; rowIndex < block.rows.length; rowIndex += 1) {
      const row = block.rows[rowIndex]
      if (row === undefined || row.kind !== 'table-row') {
        continue
      }

      for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex += 1) {
        const cell = row.cells[cellIndex]
        if (cell === undefined || cell.kind !== 'table-cell') {
          continue
        }

        const target = findRevisionTargetInBlocks(
          cell.blocks,
          [...blockPath, rowIndex, cellIndex],
          id,
        )
        if (target !== null) {
          return target
        }
      }
    }
  }

  return null
}

function getRevisionChildren(
  revision: Extract<ParagraphChild, { kind: 'ins-revision' | 'del-revision' }>,
): ReadonlyArray<ParagraphChild> {
  return revision.children as ReadonlyArray<ParagraphChild>
}

function applyAllRevisions(
  doc: Document,
  mode: 'accept' | 'reject',
): {
  document: Document
  inverse: Command
} {
  let nextDocument = doc

  for (let sectionIdx = 0; sectionIdx < nextDocument.sections.length; sectionIdx += 1) {
    nextDocument = walkBlocksResolveRevisions(nextDocument, sectionIdx, mode)
  }

  return {
    document: nextDocument,
    inverse: { kind: mode === 'accept' ? 'reject-all-revisions' : 'accept-all-revisions' },
  }
}

function walkBlocksResolveRevisions(
  doc: Document,
  sectionIndex: number,
  mode: 'accept' | 'reject',
): Document {
  const section = doc.sections[sectionIndex]
  if (section === undefined) {
    return doc
  }

  let mutated = false
  const nextBlocks = section.blocks.map((block): Block => {
    if (block.kind !== 'paragraph') {
      return block
    }

    const resolved = resolveAllRevisionsInChildren(block.children, mode)
    if (resolved === block.children) {
      return block
    }
    mutated = true
    return cloneParagraph(block, resolved)
  })

  if (!mutated) {
    return doc
  }

  const nextSection = cloneSection(section, freezeArray(nextBlocks))
  const nextSections = replaceArrayItem(doc.sections, sectionIndex, nextSection)
  return cloneDocument(doc, nextSections)
}

function resolveAllRevisionsInChildren(
  children: ReadonlyArray<ParagraphChild>,
  mode: 'accept' | 'reject',
): ReadonlyArray<ParagraphChild> {
  let mutated = false
  const next: ParagraphChild[] = []

  for (const child of children) {
    if (child.kind !== 'ins-revision' && child.kind !== 'del-revision') {
      next.push(child)
      continue
    }

    mutated = true
    const keep =
      (child.kind === 'ins-revision' && mode === 'accept') ||
      (child.kind === 'del-revision' && mode === 'reject')
    if (keep) {
      next.push(...child.children)
    }
  }

  return mutated ? freezeArray(next) : children
}

// ---------------------------------------------------------------------------
// Paragraph lookup / path resolution
// ---------------------------------------------------------------------------

export function findParagraph(
  doc: Document,
  paragraphPath: ReadonlyArray<number>,
): Paragraph | null {
  const resolvedPath = resolveParagraphPath(doc, paragraphPath)
  if (resolvedPath === null) {
    return null
  }

  const section = doc.sections[resolvedPath.sectionIndex]
  if (section === undefined) {
    return null
  }

  return findParagraphInBlocks(section.blocks, resolvedPath.blockPath)
}

export type EnclosingTable = {
  readonly tablePath: ReadonlyArray<number>
  readonly table: Table
  readonly rowIndex: number
  readonly cellIndex: number
}

/**
 * DXE-14 — given any paragraph path, finds the *innermost* table cell it
 * sits directly inside (`null` when it isn't inside one at all). Used by
 * Input.ts's Tab/Shift+Tab cell navigation and by the toolbar's table
 * commands to resolve "the table/row/cell the cursor is currently in"
 * without duplicating the block-path recursion `updateBlocksAtPath` already
 * walks for edits.
 */
export function findEnclosingTable(
  doc: Document,
  paragraphPath: ReadonlyArray<number>,
): EnclosingTable | null {
  const resolved = resolveParagraphPath(doc, paragraphPath)
  if (resolved === null) {
    return null
  }

  const section = doc.sections[resolved.sectionIndex]
  if (section === undefined) {
    return null
  }

  return findEnclosingTableInBlocks(section.blocks, [resolved.sectionIndex], resolved.blockPath)
}

function findEnclosingTableInBlocks(
  blocks: ReadonlyArray<Block>,
  pathPrefix: ReadonlyArray<number>,
  blockPath: ReadonlyArray<number>,
): EnclosingTable | null {
  if (blockPath.length === 0) {
    return null
  }

  const [blockIndex, ...rest] = blockPath
  const block = blocks[blockIndex]
  if (block === undefined || rest.length === 0 || block.kind !== 'table') {
    return null
  }

  const tablePath = freezeArray([...pathPrefix, blockIndex])
  const [rowIndex, cellIndex, ...childPath] = rest
  const row = block.rows[rowIndex]
  if (row === undefined || row.kind !== 'table-row') {
    return null
  }
  const cell = row.cells[cellIndex]
  if (cell === undefined || cell.kind !== 'table-cell') {
    return null
  }

  // Prefer a more deeply nested table when the path reaches into one.
  const nested = findEnclosingTableInBlocks(cell.blocks, freezeArray([...tablePath, rowIndex, cellIndex]), childPath)
  return nested ?? { tablePath, table: block, rowIndex, cellIndex }
}

export function replaceParagraph(
  doc: Document,
  paragraphPath: ReadonlyArray<number>,
  newPara: Paragraph,
): Document {
  return (
    updateDocumentAtParagraphPath(doc, paragraphPath, (blocks, blockIndex) => {
      const block = blocks[blockIndex]
      if (block === undefined || block.kind !== 'paragraph') {
        return null
      }

      return replaceArrayItem(blocks, blockIndex, newPara)
    }) ?? doc
  )
}

// ---------------------------------------------------------------------------
// Text editing (D12 — cross-paragraph, cross-run, hyperlink-aware)
// ---------------------------------------------------------------------------

/**
 * DXE-11 — decides which owner a freshly-typed run should get. Untracked
 * (`trackChanges` absent/disabled), this always returns `currentOwner`
 * unchanged — exactly the pre-DXE-11 behavior of extending whatever run/
 * hyperlink/owner already sits at the insertion point. Tracked, it also
 * extends when `currentOwner` is already a pending insertion by the *same*
 * author ("typing inside your own pending insertion extends it," rather than
 * nesting a new `w:ins` inside the last one on every keystroke); otherwise it
 * mints a fresh `w:ins` wrapper, so newly-typed text next to plain content,
 * next to someone else's tracked insertion, or next to an already-tracked
 * deletion, always gets its own new revision.
 */
function resolveInsertOwner(
  doc: Document,
  trackChanges: TrackChangesContext | undefined,
  currentOwner: RunOwner,
): RunOwner {
  if (trackChanges === undefined || !trackChanges.enabled) {
    return currentOwner
  }

  if (currentOwner.kind === 'ins-revision' && currentOwner.wrapper.author === trackChanges.author) {
    return currentOwner
  }

  return {
    kind: 'ins-revision',
    wrapper: createInsRevisionWrapper(doc, trackChanges),
  }
}

function createInsRevisionWrapper(doc: Document, trackChanges: TrackChangesContext): InsRevision {
  return Object.freeze({
    kind: 'ins-revision',
    id: nextRevisionId(doc),
    author: trackChanges.author,
    date: trackChanges.date,
    children: EMPTY_REVISION_CHILDREN,
  })
}

function createDelRevisionWrapper(doc: Document, trackChanges: TrackChangesContext): DelRevision {
  return Object.freeze({
    kind: 'del-revision',
    id: nextRevisionId(doc),
    author: trackChanges.author,
    date: trackChanges.date,
    children: EMPTY_REVISION_CHILDREN,
  })
}

/**
 * Every `w:id` already used by an `ins-revision`/`del-revision` anywhere in
 * the document (including inside table cells — mirroring
 * `findRevisionTargetInBlocks`'s own recursion depth), so a freshly minted
 * revision's id can never collide with one carried over from the source
 * file. `w:id` only needs to be unique, not contiguous, so returning
 * `max + 1` (starting at `1` for a document with none yet) is enough.
 */
function nextRevisionId(doc: Document): string {
  let max = 0
  for (const section of doc.sections) {
    max = Math.max(max, maxRevisionIdInBlocks(section.blocks))
  }
  return String(max + 1)
}

function maxRevisionIdInBlocks(blocks: ReadonlyArray<Block>): number {
  let max = 0
  for (const block of blocks) {
    if (block.kind === 'paragraph') {
      for (const child of block.children) {
        if (child.kind === 'ins-revision' || child.kind === 'del-revision') {
          const parsed = Number.parseInt(child.id, 10)
          if (Number.isFinite(parsed)) {
            max = Math.max(max, parsed)
          }
        }
      }
      continue
    }

    if (block.kind !== 'table') {
      continue
    }

    for (const row of block.rows) {
      if (row.kind !== 'table-row') {
        continue
      }

      for (const cell of row.cells) {
        if (cell.kind === 'table-cell') {
          max = Math.max(max, maxRevisionIdInBlocks(cell.blocks))
        }
      }
    }
  }

  return max
}

function applyInsertText(
  doc: Document,
  cmd: Extract<Command, { kind: 'insert-text' }>,
  trackChanges?: TrackChangesContext,
): {
  document: Document
  inverse: Command
  range: Range
} {
  if (cmd.text.length === 0) {
    const collapsed = { anchor: cmd.at, focus: cmd.at }
    return {
      document: doc,
      inverse: createDeleteRangeCommand(cmd.at, cmd.at),
      range: collapsed,
    }
  }

  const paragraph = requireParagraph(doc, cmd.at.paragraphPath)
  const editableRuns = requireEditableRuns(paragraph)

  if (editableRuns.length === 0) {
    if (cmd.at.runIndex !== 0 || cmd.at.charOffset !== 0) {
      throw new Error('InsertText position is outside the paragraph')
    }

    const owner = resolveInsertOwner(doc, trackChanges, DIRECT_OWNER)
    const insertedRun = createRunWithText(undefined, cmd.text)
    const nextParagraph = cloneParagraph(paragraph, buildParagraphChildren([{ run: insertedRun, owner }]))
    const nextDocument = replaceParagraphOrThrow(doc, cmd.at.paragraphPath, nextParagraph)
    const start = createPosition(cmd.at.paragraphPath, 0, 0)
    const end = createPosition(cmd.at.paragraphPath, 0, cmd.text.length)

    return {
      document: nextDocument,
      inverse: createDeleteRangeCommand(start, end),
      range: { anchor: end, focus: end },
    }
  }

  const target = resolveInsertTarget(editableRuns, cmd.at)
  const current = editableRuns[target.runIndex]

  // DXE-18/D18 — `resolveInsertTarget`'s own end-of-paragraph branch resolves
  // "insert at the very end" to `{runIndex: lastIndex, charOffset: 0}` when
  // the last entry is an atomic image/break/tab run (exactly what happens
  // when a user keeps typing right after inserting an image at the end of a
  // paragraph), or (DXE-11) an already-tracked-deletion entry, which is
  // presented with zero width the same way for the same reason (see
  // `getEditableRuns`'s `del-revision` case) even though its *real* run text
  // is non-empty. Splicing a new sibling text run in after it (rather than
  // calling createRunLike on it, which would silently replace its non-text
  // child with a plain text node — see toRunEntryOrNull — or, for a
  // del-revision entry, silently overwrite the deleted text with whatever
  // was just typed) keeps it intact and places the typed text right after
  // it, matching what the user was actually doing. Deliberately checked as
  // two explicit shape tests rather than `current.text.length === 0`: an
  // ordinary, genuinely-empty text run (e.g. an emptied-out paragraph — see
  // `applyDeleteSpan`'s own empty-paragraph fallback) also reports a
  // zero-length `.text` but is not a boundary — it must still take the
  // normal splice path below so typing into it replaces its content in
  // place instead of leaving a stray empty run behind.
  if (isAtomicRun(current.run) || current.owner.kind === 'del-revision') {
    const owner = resolveInsertOwner(doc, trackChanges, current.owner)
    const insertedRun = createRunWithText(undefined, cmd.text)
    const nextEntries: RunEntry[] = [
      ...editableRuns.slice(0, target.runIndex + 1).map((entry) => ({ run: entry.run, owner: entry.owner })),
      { run: insertedRun, owner },
      ...editableRuns.slice(target.runIndex + 1).map((entry) => ({ run: entry.run, owner: entry.owner })),
    ]

    const nextParagraph = cloneParagraph(paragraph, buildParagraphChildren(nextEntries))
    const nextDocument = replaceParagraphOrThrow(doc, cmd.at.paragraphPath, nextParagraph)
    const start = createPosition(cmd.at.paragraphPath, target.runIndex + 1, 0)
    const end = createPosition(cmd.at.paragraphPath, target.runIndex + 1, cmd.text.length)

    return {
      document: nextDocument,
      inverse: createDeleteRangeCommand(start, end),
      range: { anchor: end, focus: end },
    }
  }

  const owner = resolveInsertOwner(doc, trackChanges, current.owner)

  if (sameOwner(owner, current.owner)) {
    // Untracked, or typing inside your own pending insertion (DXE-11): a
    // simple in-place text splice, exactly the pre-DXE-11 behavior — no new
    // wrapper, no run split.
    const nextText =
      current.text.slice(0, target.charOffset) +
      cmd.text +
      current.text.slice(target.charOffset)

    const nextEntries = editableRuns.map((entry, index): RunEntry =>
      index === target.runIndex
        ? { run: createRunLike(entry.run, nextText, entry.run.props), owner: entry.owner }
        : { run: entry.run, owner: entry.owner },
    )

    const nextParagraph = cloneParagraph(paragraph, buildParagraphChildren(nextEntries))
    const nextDocument = replaceParagraphOrThrow(doc, cmd.at.paragraphPath, nextParagraph)
    const start = createPosition(cmd.at.paragraphPath, target.runIndex, target.charOffset)
    const end = createPosition(cmd.at.paragraphPath, target.runIndex, target.charOffset + cmd.text.length)

    return {
      document: nextDocument,
      inverse: createDeleteRangeCommand(start, end),
      range: { anchor: end, focus: end },
    }
  }

  // DXE-11 — Track Changes is on and the caret sits inside a run that isn't
  // already a pending insertion of ours: split that run at the caret (both
  // halves keep its original owner untouched — plain text stays plain,
  // someone else's tracked insertion stays theirs) and splice the freshly
  // typed text in between as its own new `w:ins`-owned run, inheriting the
  // split run's formatting the same way the untracked splice above always
  // has.
  const leftText = current.text.slice(0, target.charOffset)
  const rightText = current.text.slice(target.charOffset)
  const insertedRun = createRunWithText(current.run.props, cmd.text)
  const insertedIndex = target.runIndex + (leftText.length > 0 ? 1 : 0)

  const nextEntries: RunEntry[] = [
    ...editableRuns.slice(0, target.runIndex).map((entry) => ({ run: entry.run, owner: entry.owner })),
    ...(leftText.length > 0
      ? [{ run: createRunLike(current.run, leftText, current.run.props), owner: current.owner }]
      : []),
    { run: insertedRun, owner },
    ...(rightText.length > 0
      ? [{ run: createRunLike(current.run, rightText, current.run.props), owner: current.owner }]
      : []),
    ...editableRuns.slice(target.runIndex + 1).map((entry) => ({ run: entry.run, owner: entry.owner })),
  ]

  const nextParagraph = cloneParagraph(paragraph, buildParagraphChildren(nextEntries))
  const nextDocument = replaceParagraphOrThrow(doc, cmd.at.paragraphPath, nextParagraph)
  const start = createPosition(cmd.at.paragraphPath, insertedIndex, 0)
  const end = createPosition(cmd.at.paragraphPath, insertedIndex, cmd.text.length)

  return {
    document: nextDocument,
    inverse: createDeleteRangeCommand(start, end),
    range: { anchor: end, focus: end },
  }
}

function applyDeleteRange(
  doc: Document,
  cmd: DeleteRangeCommand,
  trackChanges?: TrackChangesContext,
): {
  document: Document
  inverse: Command
  range: Range
} {
  const range = normalizeRange(doc, cmd.range)
  if (comparePositions(doc, range.anchor, range.focus) === 0) {
    return {
      document: doc,
      inverse: createNoOpInsertText(range.anchor),
      range: { anchor: range.anchor, focus: range.anchor },
    }
  }

  return applyDeleteSpan(doc, range, trackChanges)
}

/**
 * Deletes `range`, whether it sits inside a single run (DXE-05: cross-run
 * within one paragraph), a single paragraph, or spans several consecutive
 * sibling paragraphs (DXE-04: cross-paragraph). All three are the same
 * operation at different scales: keep the content of the first affected
 * paragraph before `range.anchor`, keep the content of the last affected
 * paragraph after `range.focus`, merge those two halves into one paragraph,
 * and splice it in place of every paragraph the range touched.
 *
 * DXE-11 — when Track Changes is on *and* the whole range sits inside one
 * paragraph, the deleted content is kept and wrapped in a fresh `w:del`
 * (`buildTrackedDeleteParagraph`) instead of being dropped. A range spanning
 * several paragraphs still deletes untracked even with Track Changes on: a
 * merge across paragraph boundaries would need to record a deleted
 * *paragraph mark* (`w:pPr/w:rPr/w:del`) to be reversible, which this editor's
 * model doesn't represent yet — a documented scope cut rather than an
 * oversight, left for a follow-up once cross-paragraph tracked deletion is
 * needed.
 */
function applyDeleteSpan(
  doc: Document,
  range: Range,
  trackChanges?: TrackChangesContext,
): {
  document: Document
  inverse: Command
  range: Range
} {
  const startResolved = resolveParagraphPath(doc, range.anchor.paragraphPath)
  const endResolved = resolveParagraphPath(doc, range.focus.paragraphPath)
  if (startResolved === null || endResolved === null) {
    throw new Error('DeleteRange target not found')
  }

  if (startResolved.sectionIndex !== endResolved.sectionIndex) {
    throw new Error('DeleteRange across sections is not yet implemented')
  }

  const startBlockPath = startResolved.blockPath
  const endBlockPath = endResolved.blockPath
  if (
    startBlockPath.length === 0 ||
    startBlockPath.length !== endBlockPath.length
  ) {
    throw new Error('DeleteRange across mismatched block depths is not yet implemented')
  }

  const parentPrefix = startBlockPath.slice(0, -1)
  for (let index = 0; index < parentPrefix.length; index += 1) {
    if (parentPrefix[index] !== endBlockPath[index]) {
      throw new Error('DeleteRange across different table cells is not yet implemented')
    }
  }

  const startIndex = startBlockPath[startBlockPath.length - 1]
  const endIndex = endBlockPath[endBlockPath.length - 1]
  if (endIndex < startIndex) {
    throw new Error('DeleteRange has an invalid (reversed) span')
  }

  const section = doc.sections[startResolved.sectionIndex]
  if (section === undefined) {
    throw new Error('DeleteRange target not found')
  }

  const siblingBlocks = getBlocksAtPrefix(section.blocks, parentPrefix)
  if (siblingBlocks === null) {
    throw new Error('DeleteRange target not found')
  }

  const spanBlocks = siblingBlocks.slice(startIndex, endIndex + 1)
  if (spanBlocks.some((block) => block.kind !== 'paragraph')) {
    throw new Error('DeleteRange across a table is not yet implemented')
  }

  const paragraphs = spanBlocks as ReadonlyArray<Paragraph>
  const firstParagraph = paragraphs[0]
  const lastParagraph = paragraphs[paragraphs.length - 1]

  const mergedParagraph =
    trackChanges?.enabled === true && paragraphs.length === 1
      ? buildTrackedDeleteParagraph(doc, firstParagraph, range, trackChanges)
      : buildPlainDeleteParagraph(firstParagraph, lastParagraph, range)

  const nextSiblingBlocks = freezeArray([
    ...siblingBlocks.slice(0, startIndex),
    mergedParagraph,
    ...siblingBlocks.slice(endIndex + 1),
  ])
  const nextDocument = setBlocksAtPrefix(doc, startResolved.sectionIndex, parentPrefix, nextSiblingBlocks)

  const inverseAt = freezeArray([startResolved.sectionIndex, ...parentPrefix, startIndex])
  const collapsedAt = range.anchor

  return {
    document: nextDocument,
    inverse: {
      kind: 'replace-blocks',
      at: inverseAt,
      count: 1,
      blocks: paragraphs,
      cursor: { anchor: range.anchor, focus: range.focus },
    },
    range: { anchor: collapsedAt, focus: collapsedAt },
  }
}

function buildPlainDeleteParagraph(
  firstParagraph: Paragraph,
  lastParagraph: Paragraph,
  range: Range,
): Paragraph {
  const firstRuns = requireEditableRuns(firstParagraph)
  const lastRuns = requireEditableRuns(lastParagraph)

  const beforeEntries = splitEntriesAtPosition(firstRuns, range.anchor).beforeEntries
  const afterEntries = splitEntriesAtPosition(lastRuns, range.focus).afterEntries

  const mergedEntries = mergeAdjacentEntries([...beforeEntries, ...afterEntries])
  // Deleting every run's text still leaves the paragraph holding one
  // (now-empty) run rather than none, matching the representation the rest of
  // the pipeline (e.g. InsertText's own empty-paragraph branch) already
  // treats as "an empty paragraph" the model was loaded/created with.
  const mergedChildren =
    mergedEntries.length > 0
      ? buildParagraphChildren(mergedEntries)
      : freezeArray<ParagraphChild>([createRunWithText(firstRuns[0]?.run.props, '')])
  return cloneParagraph(firstParagraph, mergedChildren)
}

/**
 * DXE-11 — Track Changes is on and the whole selection sits inside one
 * paragraph: keep the selected text instead of dropping it, marking it
 * `w:del` so Reject can restore it later. The one exception is a slice
 * that's already part of a *pending insertion by this same author*: text
 * that was never actually committed is removed outright, shrinking (or
 * fully removing) that `w:ins` — the deletion counterpart of "typing inside
 * your own pending insertion extends it" — rather than nesting a `w:del`
 * inside a `w:ins` that reflects nothing Word itself does for your own
 * still-open edit. Every other slice (plain text, someone else's tracked
 * insertion, or a hyperlink's runs — whose wrapper is intentionally dropped
 * here, a documented limitation: DXE-11 doesn't extend the del-revision
 * model to carry a nested hyperlink target) is re-owned under one freshly
 * minted `DelRevision` shared across the whole deleted span, so a selection
 * touching several such slices still comes out as a single `<w:del>` rather
 * than one per slice.
 */
function buildTrackedDeleteParagraph(
  doc: Document,
  paragraph: Paragraph,
  range: Range,
  trackChanges: TrackChangesContext,
): Paragraph {
  const editableRuns = requireEditableRuns(paragraph)
  const offsets = getRangeOffsets(paragraph, range)
  const { before, within, after } = sliceEntriesByOffsets(editableRuns, offsets)

  const survivors = within.filter(
    (entry) => !(entry.owner.kind === 'ins-revision' && entry.owner.wrapper.author === trackChanges.author),
  )

  if (survivors.length === 0) {
    const mergedEntries = mergeAdjacentEntries([...before, ...after])
    const mergedChildren =
      mergedEntries.length > 0
        ? buildParagraphChildren(mergedEntries)
        : freezeArray<ParagraphChild>([createRunWithText(editableRuns[0]?.run.props, '')])
    return cloneParagraph(paragraph, mergedChildren)
  }

  const delOwner: RunOwner = { kind: 'del-revision', wrapper: createDelRevisionWrapper(doc, trackChanges) }
  const deletedEntries: RunEntry[] = survivors.map((entry) => ({ run: entry.run, owner: delOwner }))
  const mergedEntries = mergeAdjacentEntries([...before, ...deletedEntries, ...after])
  return cloneParagraph(paragraph, buildParagraphChildren(mergedEntries))
}

function applyInsertParagraphBreak(
  doc: Document,
  cmd: Extract<Command, { kind: 'insert-paragraph-break' }>,
): {
  document: Document
  inverse: Command
  range: Range
} {
  const paragraph = requireParagraph(doc, cmd.at.paragraphPath)
  const editableRuns = requireEditableRuns(paragraph)
  const split = splitEntriesAtPosition(editableRuns, cmd.at)

  const firstParagraph = cloneParagraph(paragraph, buildParagraphChildren(split.beforeEntries))
  const secondParagraph = cloneParagraph(paragraph, buildParagraphChildren(split.afterEntries))
  const nextDocument = insertParagraphAfter(doc, cmd.at.paragraphPath, firstParagraph, secondParagraph)
  const nextPath = incrementParagraphPath(cmd.at.paragraphPath)
  const nextPosition = createPosition(nextPath, 0, 0)

  return {
    document: nextDocument,
    inverse: createDeleteRangeCommand(
      createParagraphEndPosition(cmd.at.paragraphPath, toEditableRunList(split.beforeEntries)),
      nextPosition,
    ),
    range: { anchor: nextPosition, focus: nextPosition },
  }
}

function applyRunFormat(
  doc: Document,
  cmd: ApplyRunFormatCommand,
): {
  document: Document
  inverse: Command
  range?: Range
} {
  const range = normalizeRange(doc, cmd.range)
  if (comparePositions(doc, range.anchor, range.focus) === 0 || Object.keys(cmd.format).length === 0) {
    return {
      document: doc,
      inverse: createNoOpRunFormat(range),
      range,
    }
  }

  if (sameParagraphPath(doc, range.anchor.paragraphPath, range.focus.paragraphPath)) {
    return applyRunFormatSingleParagraph(doc, range, cmd.format)
  }

  return applyRunFormatAcrossParagraphs(doc, range, cmd.format)
}

function applyRunFormatSingleParagraph(
  doc: Document,
  range: Range,
  format: Partial<RunProps>,
): {
  document: Document
  inverse: Command
  range: Range
} {
  const paragraph = requireParagraph(doc, range.anchor.paragraphPath)
  const editableRuns = requireEditableRuns(paragraph)
  const offsets = getRangeOffsets(paragraph, range)
  const inverseFormat = collectInverseRunFormat(editableRuns, offsets, format)
  const formattedEntries = buildFormattedEntries(editableRuns, offsets, format)
  const mergedEntries = mergeAdjacentEntries(formattedEntries)
  const nextParagraph = cloneParagraph(paragraph, buildParagraphChildren(mergedEntries))
  const nextDocument = replaceParagraphOrThrow(doc, range.anchor.paragraphPath, nextParagraph)
  const nextEditableRuns = requireEditableRuns(nextParagraph)
  const nextRange = createRange(
    offsetToPosition(range.anchor.paragraphPath, nextEditableRuns, offsets.startOffset, 'forward'),
    offsetToPosition(range.anchor.paragraphPath, nextEditableRuns, offsets.endOffset, 'backward'),
  )

  return {
    document: nextDocument,
    inverse: {
      kind: 'apply-run-format',
      range: nextRange,
      format: inverseFormat,
    },
    range: nextRange,
  }
}

/**
 * DXE-04 — formats a selection spanning multiple paragraphs by translating it
 * into one `apply-run-format` per affected paragraph (full paragraph text for
 * every paragraph strictly between the endpoints, the anchor-to-end slice of
 * the first, the start-to-focus slice of the last) and applying them as one
 * composite command. Each sub-command computes its own homogeneous inverse
 * independently, so paragraphs that started with different formatting still
 * invert correctly — composing them is what keeps the whole multi-paragraph
 * edit a single atomic History step.
 */
function applyRunFormatAcrossParagraphs(
  doc: Document,
  range: Range,
  format: Partial<RunProps>,
): {
  document: Document
  inverse: Command
  range?: Range
} {
  const startResolved = resolveParagraphPath(doc, range.anchor.paragraphPath)
  const endResolved = resolveParagraphPath(doc, range.focus.paragraphPath)
  if (startResolved === null || endResolved === null) {
    throw new Error('ApplyRunFormat target not found')
  }
  if (startResolved.sectionIndex !== endResolved.sectionIndex) {
    throw new Error('ApplyRunFormat across sections is not yet implemented')
  }

  const startBlockPath = startResolved.blockPath
  const endBlockPath = endResolved.blockPath
  if (startBlockPath.length === 0 || startBlockPath.length !== endBlockPath.length) {
    throw new Error('ApplyRunFormat across mismatched block depths is not yet implemented')
  }

  const parentPrefix = startBlockPath.slice(0, -1)
  for (let index = 0; index < parentPrefix.length; index += 1) {
    if (parentPrefix[index] !== endBlockPath[index]) {
      throw new Error('ApplyRunFormat across different table cells is not yet implemented')
    }
  }

  const startIndex = startBlockPath[startBlockPath.length - 1]
  const endIndex = endBlockPath[endBlockPath.length - 1]
  if (endIndex < startIndex) {
    throw new Error('ApplyRunFormat has an invalid (reversed) span')
  }

  const subCommands: Command[] = []
  for (let blockIndex = startIndex; blockIndex <= endIndex; blockIndex += 1) {
    const paragraphPath = freezeArray([startResolved.sectionIndex, ...parentPrefix, blockIndex])
    const paragraph = requireParagraph(doc, paragraphPath)
    const paragraphRuns = requireEditableRuns(paragraph)
    const paragraphEnd = createParagraphEndPosition(paragraphPath, paragraphRuns)

    const anchor = blockIndex === startIndex ? range.anchor : createPosition(paragraphPath, 0, 0)
    const focus = blockIndex === endIndex ? range.focus : paragraphEnd

    if (comparePositions(doc, anchor, focus) === 0) {
      continue
    }

    subCommands.push({ kind: 'apply-run-format', range: { anchor, focus }, format })
  }

  if (subCommands.length === 0) {
    return { document: doc, inverse: createNoOpRunFormat(range), range }
  }

  return applyComposite(doc, { kind: 'composite', commands: subCommands })
}

function applyParaFormat(
  doc: Document,
  cmd: ApplyParaFormatCommand,
): {
  document: Document
  inverse: Command
} {
  if (cmd.paragraphPaths.length === 0 || Object.keys(cmd.format).length === 0) {
    return {
      document: doc,
      inverse: {
        kind: 'apply-para-format',
        paragraphPaths: cmd.paragraphPaths.map(clonePath),
        format: {},
      },
    }
  }

  const paragraphs = cmd.paragraphPaths.map((paragraphPath) => requireParagraph(doc, paragraphPath))
  const inverseFormat = collectInverseParagraphFormat(paragraphs, cmd.format)

  let nextDocument = doc
  for (const paragraphPath of cmd.paragraphPaths) {
    const paragraph = requireParagraph(nextDocument, paragraphPath)
    const nextParagraph = cloneParagraph(
      paragraph,
      paragraph.children,
      applyPropsPatch(paragraph.props, cmd.format),
    )
    nextDocument = replaceParagraphOrThrow(nextDocument, paragraphPath, nextParagraph)
  }

  return {
    document: nextDocument,
    inverse: {
      kind: 'apply-para-format',
      paragraphPaths: cmd.paragraphPaths.map(clonePath),
      format: inverseFormat,
    },
  }
}

function applyStyle(
  doc: Document,
  cmd: Extract<Command, { kind: 'apply-style' }>,
): {
  document: Document
  inverse: Command
  range: Range
} {
  const paragraph = requireParagraph(doc, cmd.paragraphPath)
  const nextParagraph = cloneParagraph(
    paragraph,
    paragraph.children,
    applyPropsPatch(paragraph.props, { pStyle: cmd.styleId }),
  )
  const nextDocument = replaceParagraphOrThrow(doc, cmd.paragraphPath, nextParagraph)
  const cursor = createPosition(cmd.paragraphPath, 0, 0)

  return {
    document: nextDocument,
    inverse: {
      kind: 'apply-para-format',
      paragraphPaths: [clonePath(cmd.paragraphPath)],
      format: { pStyle: paragraph.props?.pStyle },
    },
    range: { anchor: cursor, focus: cursor },
  }
}

function requireParagraph(doc: Document, paragraphPath: ReadonlyArray<number>): Paragraph {
  const paragraph = findParagraph(doc, paragraphPath)
  if (paragraph === null) {
    throw new Error('Paragraph not found')
  }

  return paragraph
}

function replaceParagraphOrThrow(
  doc: Document,
  paragraphPath: ReadonlyArray<number>,
  newParagraph: Paragraph,
): Document {
  const nextDocument = updateDocumentAtParagraphPath(doc, paragraphPath, (blocks, blockIndex) => {
    const block = blocks[blockIndex]
    if (block === undefined || block.kind !== 'paragraph') {
      return null
    }

    return replaceArrayItem(blocks, blockIndex, newParagraph)
  })

  if (nextDocument === null) {
    throw new Error('Paragraph not found')
  }

  return nextDocument
}

function insertParagraphAfter(
  doc: Document,
  paragraphPath: ReadonlyArray<number>,
  currentParagraph: Paragraph,
  nextParagraph: Paragraph,
): Document {
  const nextDocument = updateDocumentAtParagraphPath(doc, paragraphPath, (blocks, blockIndex) => {
    const block = blocks[blockIndex]
    if (block === undefined || block.kind !== 'paragraph') {
      return null
    }

    return freezeArray([
      ...blocks.slice(0, blockIndex),
      currentParagraph,
      nextParagraph,
      ...blocks.slice(blockIndex + 1),
    ])
  })

  if (nextDocument === null) {
    throw new Error('Paragraph not found')
  }

  return nextDocument
}

function findParagraphInBlocks(
  blocks: ReadonlyArray<Block>,
  blockPath: ReadonlyArray<number>,
): Paragraph | null {
  if (blockPath.length === 0) {
    return null
  }

  const [blockIndex, ...rest] = blockPath
  const block = blocks[blockIndex]
  if (block === undefined) {
    return null
  }

  if (rest.length === 0) {
    return block.kind === 'paragraph' ? block : null
  }

  if (block.kind !== 'table' || rest.length < 3) {
    return null
  }

  const [rowIndex, cellIndex, ...childPath] = rest
  const row = block.rows[rowIndex]
  if (row === undefined || row.kind !== 'table-row') {
    return null
  }

  const cell = row.cells[cellIndex]
  if (cell === undefined || cell.kind !== 'table-cell') {
    return null
  }

  return findParagraphInBlocks(cell.blocks, childPath)
}

function updateDocumentAtParagraphPath(
  doc: Document,
  paragraphPath: ReadonlyArray<number>,
  update: BlocksUpdater,
): Document | null {
  const resolvedPath = resolveParagraphPath(doc, paragraphPath)
  if (resolvedPath === null) {
    return null
  }

  const section = doc.sections[resolvedPath.sectionIndex]
  if (section === undefined) {
    return null
  }

  const nextBlocks = updateBlocksAtPath(section.blocks, resolvedPath.blockPath, update)
  if (nextBlocks === null) {
    return null
  }

  if (nextBlocks === section.blocks) {
    return doc
  }

  const nextSection = cloneSection(section, nextBlocks)
  const nextSections = replaceArrayItem(doc.sections, resolvedPath.sectionIndex, nextSection)

  return cloneDocument(doc, nextSections)
}

function updateBlocksAtPath(
  blocks: ReadonlyArray<Block>,
  blockPath: ReadonlyArray<number>,
  update: BlocksUpdater,
): ReadonlyArray<Block> | null {
  if (blockPath.length === 0) {
    return null
  }

  const [blockIndex, ...rest] = blockPath
  const block = blocks[blockIndex]
  if (block === undefined) {
    return null
  }

  if (rest.length === 0) {
    return update(blocks, blockIndex)
  }

  if (block.kind !== 'table' || rest.length < 3) {
    return null
  }

  const [rowIndex, cellIndex, ...childPath] = rest
  const row = block.rows[rowIndex]
  if (row === undefined || row.kind !== 'table-row') {
    return null
  }

  const cell = row.cells[cellIndex]
  if (cell === undefined || cell.kind !== 'table-cell') {
    return null
  }

  const nextCellBlocks = updateBlocksAtPath(cell.blocks, childPath, update)
  if (nextCellBlocks === null) {
    return null
  }

  if (nextCellBlocks === cell.blocks) {
    return blocks
  }

  const nextCell = cloneTableCell(cell, nextCellBlocks)
  const nextCells = replaceArrayItem(row.cells, cellIndex, nextCell)
  const nextRow = cloneTableRow(row, nextCells)
  const nextRows = replaceArrayItem(block.rows, rowIndex, nextRow)
  const nextTable = cloneTable(block, nextRows)

  return replaceArrayItem(blocks, blockIndex, nextTable)
}

/**
 * Reads the sibling-blocks array addressed by `prefix` (a blockPath with its
 * final "which block" segment already stripped) without mutating anything —
 * the read counterpart to `setBlocksAtPrefix`, both used by the range-based
 * structural operations (cross-paragraph delete, table/hyperlink insertion)
 * that need to replace several consecutive sibling blocks at once rather than
 * the single-block-at-a-time shape `updateBlocksAtPath` provides.
 */
function getBlocksAtPrefix(
  blocks: ReadonlyArray<Block>,
  prefix: ReadonlyArray<number>,
): ReadonlyArray<Block> | null {
  if (prefix.length === 0) {
    return blocks
  }

  const [blockIndex, ...rest] = prefix
  const block = blocks[blockIndex]
  if (block === undefined || block.kind !== 'table' || rest.length < 2) {
    return null
  }

  const [rowIndex, cellIndex, ...childPrefix] = rest
  const row = block.rows[rowIndex]
  if (row === undefined || row.kind !== 'table-row') {
    return null
  }

  const cell = row.cells[cellIndex]
  if (cell === undefined || cell.kind !== 'table-cell') {
    return null
  }

  return getBlocksAtPrefix(cell.blocks, childPrefix)
}

function setBlocksAtPrefix(
  doc: Document,
  sectionIndex: number,
  prefix: ReadonlyArray<number>,
  newBlocks: ReadonlyArray<Block>,
): Document {
  const section = doc.sections[sectionIndex]
  if (section === undefined) {
    throw new Error('Section not found')
  }

  const updatedBlocks = setBlocksAtPrefixRec(section.blocks, prefix, newBlocks)
  if (updatedBlocks === null) {
    throw new Error('Block path not found')
  }

  const nextSection = cloneSection(section, updatedBlocks)
  return cloneDocument(doc, replaceArrayItem(doc.sections, sectionIndex, nextSection))
}

function setBlocksAtPrefixRec(
  blocks: ReadonlyArray<Block>,
  prefix: ReadonlyArray<number>,
  newBlocks: ReadonlyArray<Block>,
): ReadonlyArray<Block> | null {
  if (prefix.length === 0) {
    return newBlocks
  }

  const [blockIndex, ...rest] = prefix
  const block = blocks[blockIndex]
  if (block === undefined || block.kind !== 'table' || rest.length < 2) {
    return null
  }

  const [rowIndex, cellIndex, ...childPrefix] = rest
  const row = block.rows[rowIndex]
  if (row === undefined || row.kind !== 'table-row') {
    return null
  }

  const cell = row.cells[cellIndex]
  if (cell === undefined || cell.kind !== 'table-cell') {
    return null
  }

  const nextCellBlocks = setBlocksAtPrefixRec(cell.blocks, childPrefix, newBlocks)
  if (nextCellBlocks === null) {
    return null
  }

  const nextCell = cloneTableCell(cell, nextCellBlocks)
  const nextRow = cloneTableRow(row, replaceArrayItem(row.cells, cellIndex, nextCell))
  const nextTable = cloneTable(block, replaceArrayItem(block.rows, rowIndex, nextRow))

  return replaceArrayItem(blocks, blockIndex, nextTable)
}

function resolveParagraphPath(
  doc: Document,
  paragraphPath: ReadonlyArray<number>,
): ResolvedParagraphPath | null {
  if (paragraphPath.length === 0) {
    return null
  }

  if (paragraphPath.length === 1) {
    return {
      sectionIndex: 0,
      blockPath: clonePath(paragraphPath),
    }
  }

  const [sectionIndex, ...blockPath] = paragraphPath
  if (doc.sections[sectionIndex] !== undefined) {
    return {
      sectionIndex,
      blockPath: freezeArray(blockPath),
    }
  }

  return {
    sectionIndex: 0,
    blockPath: clonePath(paragraphPath),
  }
}

// ---------------------------------------------------------------------------
// Owner-aware run collection (DXE-03 — hyperlink flattening)
// ---------------------------------------------------------------------------

function requireEditableRuns(paragraph: Paragraph): ReadonlyArray<EditableRun> {
  const editableRuns = getEditableRuns(paragraph)
  if (editableRuns === null) {
    throw new Error('This command currently supports paragraphs with direct text runs (optionally inside a hyperlink) only')
  }

  return editableRuns
}

function getEditableRuns(paragraph: Paragraph): ReadonlyArray<EditableRun> | null {
  const editableRuns: EditableRun[] = []

  for (const child of paragraph.children) {
    if (child.kind === 'run') {
      const entry = toRunEntryOrNull(child, DIRECT_OWNER)
      if (entry === null) {
        return null
      }

      editableRuns.push(entry)
      continue
    }

    if (child.kind === 'hyperlink') {
      const owner: RunOwner = { kind: 'hyperlink', wrapper: child }
      for (const grandchild of child.children) {
        if (grandchild.kind !== 'run') {
          return null
        }

        const entry = toRunEntryOrNull(grandchild, owner)
        if (entry === null) {
          return null
        }

        editableRuns.push(entry)
      }
      continue
    }

    if (child.kind === 'ins-revision') {
      const owner: RunOwner = { kind: 'ins-revision', wrapper: child }
      for (const grandchild of child.children) {
        if (grandchild.kind !== 'run') {
          return null
        }

        const entry = toRunEntryOrNull(grandchild, owner)
        if (entry === null) {
          return null
        }

        editableRuns.push(entry)
      }
      continue
    }

    if (child.kind === 'del-revision') {
      // DXE-11 — an already-tracked deletion's text is presented with zero
      // width, exactly like an atomic image/tab/break run (see
      // `toRunEntryOrNull`): the surrounding paragraph stays fully editable,
      // but the struck-through text itself is never a target a caret
      // position can land inside or a delete/insert range can split — it
      // reads as a single boundary you type/delete around, matching how a
      // reviewer reads (not edits) already-rejected-pending content. Each
      // grandchild run must still be text-only (checked via `getRunText`, not
      // just assumed) so a deleted drawing/tab/break — a shape this editor
      // doesn't specifically model as "deleted" — still conservatively
      // rejects the whole paragraph rather than silently mishandling it.
      const owner: RunOwner = { kind: 'del-revision', wrapper: child }
      for (const grandchild of child.children) {
        if (grandchild.kind !== 'run' || getRunText(grandchild) === null) {
          return null
        }

        editableRuns.push({ run: grandchild, owner, text: '' })
      }
      continue
    }

    return null
  }

  return editableRuns
}

/**
 * DXE-21 fix — exposes this module's owner-aware run flattening (hyperlink
 * content and atomic image/break/tab runs included) to editor code outside
 * the Command pipeline that only needs character-offset math over a
 * paragraph's real editable text, without reconstructing paragraph children.
 *
 * Input.ts's word-boundary/cursor helpers used to keep their own narrow
 * "every child must be a plain text-only run" flattener, which rejected the
 * *entire* paragraph — not just the offending run — the moment it contained
 * a hyperlink or an atomic image/page-break/tab run, exactly the shapes D12
 * and D14/D18 add first-class editing support for elsewhere in this module.
 * That rejection was silently reinterpreted by the new cross-paragraph
 * word-boundary-delete logic as "this paragraph is empty," which then jumped
 * the deletion into the *adjacent* paragraph and merged it away entirely —
 * e.g. Ctrl+Backspace to delete one word, in a paragraph that merely
 * contains an inline image anywhere in it, could silently delete the whole
 * previous paragraph instead. Routing through the same flattening logic the
 * rest of this module already trusts for editing fixes that at the source.
 */
export function getFlatTextRuns(
  paragraph: Paragraph,
): ReadonlyArray<{ readonly text: string; readonly run: Run }> | null {
  return getEditableRuns(paragraph)
}

/**
 * DXE-18/D18 — `insertImage`/page-break insertion (`applyInsertInline`) each
 * put their inline leaf (a `Drawing`, `BreakNode`, or `TabNode`) into its own
 * dedicated run, never mixed with text. Before this, `getEditableRuns`
 * rejected the *entire paragraph* the instant it saw any such run (`getRunText`
 * returns `null` for a run with a non-text child) — so a paragraph became
 * permanently uneditable (every later insert/delete/format silently no-opped,
 * see D12's always-preventDefault behavior) the moment it gained an inline
 * image or a page break, including from typing immediately after it — by far
 * the most common thing a user does right after inserting one.
 *
 * Recognizing this one specific, self-produced shape (a run with *exactly
 * one* child that is a known atomic leaf) lets such a run flatten into the
 * list as a zero-length entry instead: `splitEntriesAtPosition` already
 * treats a zero-length entry as an all-or-nothing boundary (never slices its
 * "text"), so cross-run/cross-paragraph delete and insert both work correctly
 * around it for free. `sliceEntriesByOffsets`/`mergeAdjacentEntries` are
 * additionally guarded (see below) so a range that merely *touches* an atomic
 * entry's boundary — formatting or deleting text elsewhere in the same
 * paragraph — can never silently rewrite or drop it via `createRunLike`
 * (which would otherwise replace its non-text child with an empty text node,
 * since `getRunText` returning `null` makes `createRunLike`'s "already
 * matches, reuse as-is" check never match). A run mixing text with a
 * non-text child, or containing any other unrecognized non-text shape, still
 * rejects the whole paragraph exactly as before — this only widens support
 * for the specific shape Atlas's own editor produces.
 */
function toRunEntryOrNull(run: Run, owner: RunOwner): EditableRun | null {
  if (isAtomicRun(run)) {
    return { run, text: '', owner }
  }

  const text = getRunText(run)
  return text === null ? null : { run, text, owner }
}

const ATOMIC_LEAF_KINDS: ReadonlySet<string> = new Set(['break', 'tab', 'drawing'])

function isAtomicRun(run: Run): boolean {
  return run.children.length === 1 && ATOMIC_LEAF_KINDS.has(run.children[0].kind)
}

function toEditableRunList(entries: ReadonlyArray<RunEntry>): ReadonlyArray<EditableRun> {
  return entries.map((entry) => {
    const text = getRunText(entry.run)
    return { ...entry, text: text ?? '' }
  })
}

function getRunText(run: Run): string | null {
  let text = ''

  for (const child of run.children) {
    if (child.kind !== 'text') {
      return null
    }

    text += child.value
  }

  return text
}

function sameOwner(a: RunOwner, b: RunOwner): boolean {
  if (a.kind === 'direct' && b.kind === 'direct') {
    return true
  }

  if (a.kind === 'hyperlink' && b.kind === 'hyperlink') {
    return a.wrapper === b.wrapper
  }

  if (a.kind === 'ins-revision' && b.kind === 'ins-revision') {
    return a.wrapper === b.wrapper
  }

  return a.kind === 'del-revision' && b.kind === 'del-revision' && a.wrapper === b.wrapper
}

/**
 * Rebuilds a paragraph's `children` from a flat, owner-tagged run list:
 * consecutive runs that share a `direct` owner become bare `Run` children in
 * order; consecutive runs that share the *same* hyperlink owner are grouped
 * back into one `Hyperlink` wrapper (reusing that wrapper's original
 * relationship/anchor/tooltip attributes, dropping only its interior
 * non-run children, which `getEditableRuns` already refused to flatten).
 */
function buildParagraphChildren(entries: ReadonlyArray<RunEntry>): ReadonlyArray<ParagraphChild> {
  const children: ParagraphChild[] = []
  let index = 0

  while (index < entries.length) {
    const owner = entries[index].owner
    const group: Run[] = []

    while (index < entries.length && sameOwner(entries[index].owner, owner)) {
      group.push(entries[index].run)
      index += 1
    }

    if (owner.kind === 'direct') {
      children.push(...group)
    } else {
      children.push(Object.freeze({ ...owner.wrapper, children: freezeArray(group) }))
    }
  }

  return freezeArray(children)
}

function resolveInsertTarget(
  editableRuns: ReadonlyArray<EditableRun>,
  at: Position,
): {
  readonly runIndex: number
  readonly charOffset: number
} {
  if (at.runIndex < 0) {
    throw new Error('InsertText position is outside the paragraph')
  }

  if (at.runIndex === editableRuns.length) {
    if (at.charOffset !== 0) {
      throw new Error('InsertText position is outside the paragraph')
    }

    const lastIndex = editableRuns.length - 1
    return {
      runIndex: lastIndex,
      charOffset: editableRuns[lastIndex].text.length,
    }
  }

  const run = editableRuns[at.runIndex]
  if (run === undefined || at.charOffset < 0 || at.charOffset > run.text.length) {
    throw new Error('InsertText position is outside the paragraph')
  }

  return {
    runIndex: at.runIndex,
    charOffset: at.charOffset,
  }
}

function splitEntriesAtPosition(
  editableRuns: ReadonlyArray<EditableRun>,
  at: Position,
): {
  readonly beforeEntries: ReadonlyArray<RunEntry>
  readonly afterEntries: ReadonlyArray<RunEntry>
} {
  const entries: ReadonlyArray<RunEntry> = editableRuns.map((entry) => ({ run: entry.run, owner: entry.owner }))

  if (editableRuns.length === 0) {
    if (at.runIndex !== 0 || at.charOffset !== 0) {
      throw new Error('Position is outside the paragraph')
    }

    return { beforeEntries: freezeArray<RunEntry>([]), afterEntries: freezeArray<RunEntry>([]) }
  }

  if (at.runIndex === editableRuns.length) {
    if (at.charOffset !== 0) {
      throw new Error('Position is outside the paragraph')
    }

    return { beforeEntries: freezeArray(entries.slice()), afterEntries: freezeArray<RunEntry>([]) }
  }

  const target = editableRuns[at.runIndex]
  if (target === undefined || at.charOffset < 0 || at.charOffset > target.text.length) {
    throw new Error('Position is outside the paragraph')
  }

  if (at.charOffset === 0) {
    return {
      beforeEntries: freezeArray(entries.slice(0, at.runIndex)),
      afterEntries: freezeArray(entries.slice(at.runIndex)),
    }
  }

  if (at.charOffset === target.text.length) {
    return {
      beforeEntries: freezeArray(entries.slice(0, at.runIndex + 1)),
      afterEntries: freezeArray(entries.slice(at.runIndex + 1)),
    }
  }

  const leftRun = createRunLike(target.run, target.text.slice(0, at.charOffset), target.run.props)
  const rightRun = createRunLike(target.run, target.text.slice(at.charOffset), target.run.props)

  return {
    beforeEntries: freezeArray([...entries.slice(0, at.runIndex), { run: leftRun, owner: target.owner }]),
    afterEntries: freezeArray([{ run: rightRun, owner: target.owner }, ...entries.slice(at.runIndex + 1)]),
  }
}

/**
 * Slices a flat run list by absolute character offsets into three owner-aware
 * groups (`before`/`within`/`after`), splitting any run that straddles a
 * boundary. Shared by run-formatting (the `within` slice gets the format
 * patch) and hyperlink-wrapping (the `within` slice gets re-owned).
 */
function sliceEntriesByOffsets(
  editableRuns: ReadonlyArray<EditableRun>,
  offsets: RunRange,
): {
  readonly before: ReadonlyArray<RunEntry>
  readonly within: ReadonlyArray<RunEntry>
  readonly after: ReadonlyArray<RunEntry>
} {
  const before: RunEntry[] = []
  const within: RunEntry[] = []
  const after: RunEntry[] = []
  let currentOffset = 0

  for (const entry of editableRuns) {
    const runStart = currentOffset
    const runEnd = currentOffset + entry.text.length
    currentOffset = runEnd

    if (runEnd <= offsets.startOffset) {
      before.push({ run: entry.run, owner: entry.owner })
      continue
    }

    if (runStart >= offsets.endOffset) {
      after.push({ run: entry.run, owner: entry.owner })
      continue
    }

    // An atomic entry (image/break/tab, `text: ''`) has runStart === runEnd,
    // so it can only reach here when the range strictly straddles it
    // (offsets.startOffset < runStart < offsets.endOffset) — e.g. bolding or
    // linking a selection that happens to span an inline image. Since it has
    // no text to slice, the code below would otherwise compute
    // localStart === localEnd === 0 and drop the entry from all three
    // buckets — silently deleting it. Keep it, untouched, on the boundary
    // side it started on rather than ever placing it in `within` (which
    // rewrites entries via `createRunLike`, and `createRunLike` on an atomic
    // run always rebuilds it as a bare text node — see `toRunEntryOrNull`).
    if (entry.text.length === 0) {
      before.push({ run: entry.run, owner: entry.owner })
      continue
    }

    const localStart = Math.max(0, offsets.startOffset - runStart)
    const localEnd = Math.min(entry.text.length, offsets.endOffset - runStart)

    if (localStart > 0) {
      before.push({ run: createRunLike(entry.run, entry.text.slice(0, localStart), entry.run.props), owner: entry.owner })
    }
    if (localEnd > localStart) {
      within.push({
        run: createRunLike(entry.run, entry.text.slice(localStart, localEnd), entry.run.props),
        owner: entry.owner,
      })
    }
    if (localEnd < entry.text.length) {
      after.push({ run: createRunLike(entry.run, entry.text.slice(localEnd), entry.run.props), owner: entry.owner })
    }
  }

  return { before: freezeArray(before), within: freezeArray(within), after: freezeArray(after) }
}

function normalizeRange(doc: Document, range: Range): Range {
  if (comparePositions(doc, range.anchor, range.focus) <= 0) {
    return createRange(range.anchor, range.focus)
  }

  return createRange(range.focus, range.anchor)
}

function sameParagraphPath(
  doc: Document,
  left: ReadonlyArray<number>,
  right: ReadonlyArray<number>,
): boolean {
  const leftPath = normalizePathForComparison(doc, left)
  const rightPath = normalizePathForComparison(doc, right)

  return compareNumberPaths(leftPath, rightPath) === 0
}

function comparePositions(doc: Document, left: Position, right: Position): number {
  const pathCompare = compareNumberPaths(
    normalizePathForComparison(doc, left.paragraphPath),
    normalizePathForComparison(doc, right.paragraphPath),
  )

  if (pathCompare !== 0) {
    return pathCompare
  }

  if (left.runIndex !== right.runIndex) {
    return left.runIndex - right.runIndex
  }

  return left.charOffset - right.charOffset
}

function normalizePathForComparison(
  doc: Document,
  paragraphPath: ReadonlyArray<number>,
): ReadonlyArray<number> {
  const resolvedPath = resolveParagraphPath(doc, paragraphPath)
  if (resolvedPath === null) {
    return clonePath(paragraphPath)
  }

  return freezeArray([resolvedPath.sectionIndex, ...resolvedPath.blockPath])
}

function compareNumberPaths(
  left: ReadonlyArray<number>,
  right: ReadonlyArray<number>,
): number {
  const maxLength = Math.max(left.length, right.length)

  for (let index = 0; index < maxLength; index += 1) {
    const leftValue = left[index]
    const rightValue = right[index]

    if (leftValue === undefined) {
      return -1
    }

    if (rightValue === undefined) {
      return 1
    }

    if (leftValue !== rightValue) {
      return leftValue - rightValue
    }
  }

  return 0
}

function getRangeOffsets(paragraph: Paragraph, range: Range): RunRange {
  const startOffset = positionToOffset(paragraph, range.anchor)
  const endOffset = positionToOffset(paragraph, range.focus)

  return { startOffset, endOffset }
}

function positionToOffset(paragraph: Paragraph, position: Position): number {
  const editableRuns = requireEditableRuns(paragraph)

  if (editableRuns.length === 0) {
    if (position.runIndex === 0 && position.charOffset === 0) {
      return 0
    }

    throw new Error('Position is outside the paragraph')
  }

  let offset = 0
  for (let index = 0; index < editableRuns.length; index += 1) {
    const run = editableRuns[index]
    if (index === position.runIndex) {
      if (position.charOffset < 0 || position.charOffset > run.text.length) {
        throw new Error('Position is outside the run')
      }

      return offset + position.charOffset
    }

    offset += run.text.length
  }

  if (position.runIndex === editableRuns.length && position.charOffset === 0) {
    return offset
  }

  throw new Error('Position is outside the paragraph')
}

function offsetToPosition(
  paragraphPath: ReadonlyArray<number>,
  editableRuns: ReadonlyArray<EditableRun>,
  offset: number,
  bias: 'forward' | 'backward',
): Position {
  if (editableRuns.length === 0) {
    return createPosition(paragraphPath, 0, 0)
  }

  let currentOffset = 0
  for (let index = 0; index < editableRuns.length; index += 1) {
    const run = editableRuns[index]
    const nextOffset = currentOffset + run.text.length

    if (offset < nextOffset) {
      return createPosition(paragraphPath, index, offset - currentOffset)
    }

    if (offset === currentOffset && bias === 'forward') {
      return createPosition(paragraphPath, index, 0)
    }

    if (offset === nextOffset) {
      if (bias === 'forward' && index < editableRuns.length - 1) {
        return createPosition(paragraphPath, index + 1, 0)
      }

      return createPosition(paragraphPath, index, run.text.length)
    }

    currentOffset = nextOffset
  }

  const lastIndex = editableRuns.length - 1
  return createPosition(paragraphPath, lastIndex, editableRuns[lastIndex].text.length)
}

function buildFormattedEntries(
  editableRuns: ReadonlyArray<EditableRun>,
  offsets: RunRange,
  format: Partial<RunProps>,
): ReadonlyArray<RunEntry> {
  const sliced = sliceEntriesByOffsets(editableRuns, offsets)
  const formattedWithin = sliced.within.map((entry): RunEntry => {
    const text = getRunText(entry.run) ?? ''
    return { run: createRunLike(entry.run, text, applyPropsPatch(entry.run.props, format)), owner: entry.owner }
  })

  return freezeArray([...sliced.before, ...formattedWithin, ...sliced.after])
}

function collectInverseRunFormat(
  editableRuns: ReadonlyArray<EditableRun>,
  offsets: RunRange,
  format: Partial<RunProps>,
): Partial<RunProps> {
  const targetProps: Array<RunProps | undefined> = []
  let currentOffset = 0

  for (const entry of editableRuns) {
    const runStart = currentOffset
    const runEnd = currentOffset + entry.text.length

    if (runEnd > offsets.startOffset && runStart < offsets.endOffset) {
      targetProps.push(entry.run.props)
    }

    currentOffset = runEnd
  }

  return collectHomogeneousPatchValues(targetProps, format)
}

function collectInverseParagraphFormat(
  paragraphs: ReadonlyArray<Paragraph>,
  format: Partial<ParaProps>,
): Partial<ParaProps> {
  return collectHomogeneousPatchValues(
    paragraphs.map((paragraph) => paragraph.props),
    format,
  )
}

function collectHomogeneousPatchValues<T extends object>(
  sources: ReadonlyArray<T | undefined>,
  format: Partial<T>,
): Partial<T> {
  const inversePatch: Partial<T> = {}

  for (const key of getPatchKeys(format)) {
    const values = sources.map((source) => source?.[key])
    if (!allValuesMatch(values)) {
      throw new Error('Cannot invert a format command over heterogeneous existing values')
    }

    inversePatch[key] = values[0]
  }

  return inversePatch
}

function getPatchKeys<T extends object>(format: Partial<T>): ReadonlyArray<keyof T> {
  return Object.keys(format) as Array<keyof T>
}

function allValuesMatch(values: ReadonlyArray<unknown>): boolean {
  if (values.length <= 1) {
    return true
  }

  const first = values[0]
  return values.every((value) => isDeepEqual(value, first))
}

function applyPropsPatch<T extends object>(
  base: T | undefined,
  patch: Partial<T>,
): T | undefined {
  const nextEntries = new Map<string, unknown>()

  if (base !== undefined) {
    for (const [key, value] of Object.entries(base)) {
      nextEntries.set(key, value)
    }
  }

  for (const key of getPatchKeys(patch)) {
    const value = patch[key]
    if (value === undefined) {
      nextEntries.delete(String(key))
      continue
    }

    nextEntries.set(String(key), value)
  }

  if (nextEntries.size === 0) {
    return undefined
  }

  const result: Record<string, unknown> = {}
  for (const [key, value] of nextEntries.entries()) {
    result[key] = value
  }

  return Object.freeze(result) as T
}

function mergeAdjacentEntries(entries: ReadonlyArray<RunEntry>): ReadonlyArray<RunEntry> {
  const merged: RunEntry[] = []

  for (const entry of entries) {
    // Atomic entries (image/break/tab) have no text to concatenate with a
    // neighbor, and createRunLike would otherwise destroy them when asked to
    // rebuild one with different "text" (see toRunEntryOrNull) — always keep
    // them standalone rather than attempting to merge into or out of them.
    if (isAtomicRun(entry.run)) {
      merged.push(entry)
      continue
    }

    const text = getRunText(entry.run)
    if (text === null) {
      throw new Error('This command currently supports text-only runs')
    }

    const previous = merged[merged.length - 1]
    if (
      previous === undefined ||
      isAtomicRun(previous.run) ||
      !sameOwner(previous.owner, entry.owner) ||
      !sameRunProps(previous.run.props, entry.run.props)
    ) {
      merged.push(entry)
      continue
    }

    const previousText = getRunText(previous.run)
    if (previousText === null) {
      throw new Error('This command currently supports text-only runs')
    }

    merged[merged.length - 1] = {
      run: createRunLike(previous.run, previousText + text, previous.run.props),
      owner: previous.owner,
    }
  }

  return freezeArray(merged)
}

function sameRunProps(left: RunProps | undefined, right: RunProps | undefined): boolean {
  return isDeepEqual(left, right)
}

function isDeepEqual(left: unknown, right: unknown): boolean {
  if (left === undefined || right === undefined) {
    return left === right
  }

  return JSON.stringify(left) === JSON.stringify(right)
}

function createParagraphEndPosition(
  paragraphPath: ReadonlyArray<number>,
  editableRuns: ReadonlyArray<EditableRun>,
): Position {
  if (editableRuns.length === 0) {
    return createPosition(paragraphPath, 0, 0)
  }

  const lastIndex = editableRuns.length - 1
  return createPosition(paragraphPath, lastIndex, editableRuns[lastIndex].text.length)
}

function createPosition(
  paragraphPath: ReadonlyArray<number>,
  runIndex: number,
  charOffset: number,
): Position {
  return {
    paragraphPath: clonePath(paragraphPath),
    runIndex,
    charOffset,
  }
}

function createRange(anchor: Position, focus: Position): Range {
  return {
    anchor: createPosition(anchor.paragraphPath, anchor.runIndex, anchor.charOffset),
    focus: createPosition(focus.paragraphPath, focus.runIndex, focus.charOffset),
  }
}

function createDeleteRangeCommand(anchor: Position, focus: Position): DeleteRangeCommand {
  return {
    kind: 'delete-range',
    range: createRange(anchor, focus),
  }
}

function createNoOpInsertText(at: Position): Command {
  return {
    kind: 'insert-text',
    at: createPosition(at.paragraphPath, at.runIndex, at.charOffset),
    text: '',
  }
}

function createNoOpRunFormat(range: Range): Command {
  return {
    kind: 'apply-run-format',
    range: createRange(range.anchor, range.focus),
    format: {},
  }
}

function incrementParagraphPath(paragraphPath: ReadonlyArray<number>): ReadonlyArray<number> {
  if (paragraphPath.length === 0) {
    throw new Error('Cannot increment an empty paragraph path')
  }

  const nextPath = paragraphPath.slice()
  nextPath[nextPath.length - 1] += 1
  return freezeArray(nextPath)
}

function clonePath(path: ReadonlyArray<number>): ReadonlyArray<number> {
  return freezeArray(path.slice())
}

function emptyParagraph(): Paragraph {
  return Object.freeze({ kind: 'paragraph', children: freezeArray<ParagraphChild>([]) })
}

function cloneParagraph(
  paragraph: Paragraph,
  children: ReadonlyArray<Run> | ReadonlyArray<ParagraphChild>,
  props?: ParaProps,
): Paragraph {
  const nextProps = arguments.length >= 3 ? props : paragraph.props

  return Object.freeze({
    kind: 'paragraph',
    ...(nextProps !== undefined ? { props: nextProps } : {}),
    children: freezeArray(children.slice()),
  })
}

function cloneSection(section: Section, blocks: ReadonlyArray<Block>): Section {
  return Object.freeze({
    kind: 'section',
    props: section.props,
    blocks,
  })
}

function cloneTable(table: Table, rows: ReadonlyArray<Table['rows'][number]>): Table {
  return Object.freeze({
    kind: 'table',
    ...(table.props !== undefined ? { props: table.props } : {}),
    ...(table.tblGrid !== undefined ? { tblGrid: table.tblGrid } : {}),
    rows,
  })
}

function cloneTableRow(
  row: TableRow,
  cells: ReadonlyArray<TableRow['cells'][number]>,
): TableRow {
  return Object.freeze({
    kind: 'table-row',
    ...(row.props !== undefined ? { props: row.props } : {}),
    cells,
  })
}

function cloneTableCell(cell: TableCell, blocks: ReadonlyArray<Block>): TableCell {
  return Object.freeze({
    kind: 'table-cell',
    ...(cell.props !== undefined ? { props: cell.props } : {}),
    blocks,
  })
}

function cloneDocument(doc: Document, sections: ReadonlyArray<Section>): Document {
  return Object.freeze({
    ...doc,
    sections,
  })
}

function createRunLike(run: Run, text: string, props: RunProps | undefined): Run {
  const currentText = getRunText(run)
  if (currentText === text && sameRunProps(run.props, props)) {
    return run
  }

  return createRunWithText(props, text)
}

function createRunWithText(props: RunProps | undefined, text: string): Run {
  return Object.freeze({
    kind: 'run',
    ...(props !== undefined ? { props } : {}),
    children: freezeArray([createTextNode(text)]),
  })
}

function createTextNode(value: string): TextNode {
  return Object.freeze({
    kind: 'text',
    value,
    ...(/^[\s]|[\s]$/.test(value) ? { preserveSpace: true } : {}),
  })
}

function replaceArrayItem<T>(
  items: ReadonlyArray<T>,
  index: number,
  value: T,
): ReadonlyArray<T> {
  if (items[index] === value) {
    return items
  }

  const nextItems = items.slice()
  nextItems[index] = value
  return freezeArray(nextItems)
}

function freezeArray<T>(items: ReadonlyArray<T>): ReadonlyArray<T> {
  return Object.freeze([...items])
}
