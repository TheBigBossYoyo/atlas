import type { Block, ParaProps, RunChild, RunProps, Table, TableProps } from '../model'

export type Position = {
  readonly paragraphPath: ReadonlyArray<number>
  readonly runIndex: number
  readonly charOffset: number
}

/**
 * DXE-11 — passed alongside a forward `insert-text`/`delete-range` command
 * (never alongside History's own undo/redo replay — see `commands.ts`'s
 * `applyInsertText`/`applyDeleteSpan` doc comments for why an *untracked*
 * replay of the stored inverse is exactly what makes undo/redo of a tracked
 * edit come out coherent for free) to record the edit as `w:ins`/`w:del`
 * instead of mutating the run directly, when the document's Track Changes
 * setting is on.
 */
export type TrackChangesContext = {
  readonly enabled: boolean
  readonly author: string
  /** ISO 8601 timestamp, e.g. `new Date().toISOString()`. */
  readonly date: string
}

export type Range = {
  readonly anchor: Position
  readonly focus: Position
}

export type InsertTextCommand = {
  readonly kind: 'insert-text'
  readonly at: Position
  readonly text: string
}

export type DeleteRangeCommand = {
  readonly kind: 'delete-range'
  readonly range: Range
}

export type InsertParagraphBreakCommand = {
  readonly kind: 'insert-paragraph-break'
  readonly at: Position
}

export type ApplyRunFormatCommand = {
  readonly kind: 'apply-run-format'
  readonly range: Range
  readonly format: Partial<RunProps>
}

export type ApplyParaFormatCommand = {
  readonly kind: 'apply-para-format'
  readonly paragraphPaths: ReadonlyArray<ReadonlyArray<number>>
  readonly format: Partial<ParaProps>
}

export type InsertTableCommand = {
  readonly kind: 'insert-table'
  readonly at: Position
  readonly rows: number
  readonly cols: number
}

export type InsertHyperlinkCommand = {
  readonly kind: 'insert-hyperlink'
  readonly range: Range
  readonly url: string
  /**
   * Relationship id (e.g. `rId7`) the caller has already allocated in the
   * bundle's `word/_rels/document.xml.rels` for this external link. Allocation
   * lives outside the pure Document model (it must avoid colliding with the
   * bundle's existing relationships), so callers build this via
   * `insertHyperlinkIntoBundle` rather than constructing the command directly.
   */
  readonly relationshipId: string
}

/**
 * Inserts a single inline leaf (an image `Drawing`, a page/column `BreakNode`,
 * etc.) at a caret position, splitting the enclosing run if the caret sits
 * mid-run. Used for image insertion (DXE-18) and page breaks (DXE-06) so both
 * go through the same Command/History pipeline and are undoable.
 */
export type InsertInlineCommand = {
  readonly kind: 'insert-inline'
  readonly at: Position
  readonly child: RunChild
}

/** Generic composite: applies each sub-command in order as a single atomic
 * unit and undoes/redoes as one History step (DXE-02/DXE-17). */
export type CompositeCommand = {
  readonly kind: 'composite'
  readonly commands: ReadonlyArray<Command>
}

/**
 * Replaces `count` consecutive sibling blocks starting at `at` (a
 * paragraphPath-shaped `[sectionIndex, ...blockPath]` addressing the first
 * replaced block) with `blocks`. This is the general-purpose exact-inverse
 * primitive for structural edits (cross-paragraph delete, table/hyperlink
 * insertion): the forward edit is applied directly against the model, and its
 * inverse is expressed as a `replace-blocks` command carrying the original
 * blocks verbatim, so undo restores them exactly. `cursor`, when present, is
 * the selection to restore once this command is applied (used by undo to put
 * the caret back where the user was before the edit it is undoing).
 */
export type ReplaceBlocksCommand = {
  readonly kind: 'replace-blocks'
  readonly at: ReadonlyArray<number>
  readonly count: number
  readonly blocks: ReadonlyArray<Block>
  readonly cursor?: Range
}

export type ApplyStyleCommand = {
  readonly kind: 'apply-style'
  readonly paragraphPath: ReadonlyArray<number>
  readonly styleId: string
}

export type InsertListCommand = {
  readonly kind: 'insert-list'
  readonly paragraphPaths: ReadonlyArray<ReadonlyArray<number>>
  readonly numId: number
  readonly level: number
}

export type ChangeListLevelCommand = {
  readonly kind: 'change-list-level'
  readonly paragraphPath: ReadonlyArray<number>
  readonly delta: 1 | -1
}

// ---------------------------------------------------------------------------
// DXE-14 — table structural editing
// ---------------------------------------------------------------------------
//
// Every command below addresses the table itself with `tablePath`: a
// paragraphPath-shaped `[sectionIndex, ...blockPath]` ending exactly at the
// table's own block index (the same addressing `insert-table`'s cursor
// position resolves to, and what `resolveParagraphPath`/`updateBlocksAtPath`
// already walk generically for a *nested* table inside a cell). Structural
// removals (`delete-table-row`/`delete-table-column`, merge, split, resize)
// all invert via `replace-table`, an exact verbatim snapshot of the table
// before the change — the same "capture the original, don't try to compute a
// symmetric inverse operation" philosophy `replace-blocks` already uses for
// paragraph-level structural edits. A pure structural *insert* (a fresh,
// empty row/column) has nothing worth preserving, so its own inverse is just
// the matching delete at the same index.

export type InsertTableRowCommand = {
  readonly kind: 'insert-table-row'
  readonly tablePath: ReadonlyArray<number>
  /** Row index the new row is inserted before; `rows.length` appends. */
  readonly at: number
}

export type DeleteTableRowCommand = {
  readonly kind: 'delete-table-row'
  readonly tablePath: ReadonlyArray<number>
  readonly rowIndex: number
}

export type InsertTableColumnCommand = {
  readonly kind: 'insert-table-column'
  readonly tablePath: ReadonlyArray<number>
  /** Grid column index the new column is inserted before; the column count
   * appends. Falling inside an existing merged (`gridSpan > 1`) cell widens
   * that cell by one column instead of splitting it. */
  readonly at: number
  readonly widthTwips?: number
}

export type DeleteTableColumnCommand = {
  readonly kind: 'delete-table-column'
  readonly tablePath: ReadonlyArray<number>
  readonly columnIndex: number
}

export type DeleteTableCommand = {
  readonly kind: 'delete-table'
  readonly tablePath: ReadonlyArray<number>
}

export type MergeTableCellsCommand = {
  readonly kind: 'merge-table-cells'
  readonly tablePath: ReadonlyArray<number>
  readonly rowIndex: number
  /** Inclusive cell-index range within the row to merge into one cell
   * (horizontal merge only — see the module doc comment for the vertical/
   * `vMerge` follow-up this intentionally leaves out). */
  readonly fromCellIndex: number
  readonly toCellIndex: number
}

export type SplitTableCellCommand = {
  readonly kind: 'split-table-cell'
  readonly tablePath: ReadonlyArray<number>
  readonly rowIndex: number
  readonly cellIndex: number
  /** Number of cells to split into, from 2 up to the cell's current
   * `gridSpan`. Omitted defaults to the full `gridSpan` (undo a merge back
   * into single-column cells). */
  readonly into?: number
}

export type ResizeTableColumnCommand = {
  readonly kind: 'resize-table-column'
  readonly tablePath: ReadonlyArray<number>
  readonly columnIndex: number
  readonly widthTwips: number
}

/** The universal exact inverse for a destructive table-structure edit — see
 * the module doc comment above. */
export type ReplaceTableCommand = {
  readonly kind: 'replace-table'
  readonly tablePath: ReadonlyArray<number>
  readonly table: Table
}

/**
 * DXE-14 — table properties dialog (width/alignment/borders basics).
 * Replaces the table's whole `props` object outright rather than merging
 * (like `apply-run-format`/`apply-para-format` do): the dialog always
 * submits every field it shows, so a full replacement — inverting to the
 * exact original `props` — is simpler and just as exact as computing a
 * field-by-field merge inverse would be.
 */
export type ApplyTablePropsCommand = {
  readonly kind: 'apply-table-props'
  readonly tablePath: ReadonlyArray<number>
  readonly props: TableProps | undefined
}

type RevisionTarget =
  | {
      readonly id: string
    }
  | {
      readonly paragraphPath: ReadonlyArray<number>
      readonly childIndex: number
    }

export type AcceptRevisionCommand = {
  readonly kind: 'accept-revision'
} & RevisionTarget

export type RejectRevisionCommand = {
  readonly kind: 'reject-revision'
} & RevisionTarget

export type AcceptAllRevisionsCommand = {
  readonly kind: 'accept-all-revisions'
}

export type RejectAllRevisionsCommand = {
  readonly kind: 'reject-all-revisions'
}

export type Command =
  | InsertTextCommand
  | DeleteRangeCommand
  | InsertParagraphBreakCommand
  | ApplyRunFormatCommand
  | ApplyParaFormatCommand
  | InsertTableCommand
  | InsertHyperlinkCommand
  | InsertInlineCommand
  | CompositeCommand
  | ReplaceBlocksCommand
  | ApplyStyleCommand
  | InsertListCommand
  | ChangeListLevelCommand
  | AcceptRevisionCommand
  | RejectRevisionCommand
  | AcceptAllRevisionsCommand
  | RejectAllRevisionsCommand
  | InsertTableRowCommand
  | DeleteTableRowCommand
  | InsertTableColumnCommand
  | DeleteTableColumnCommand
  | DeleteTableCommand
  | MergeTableCellsCommand
  | SplitTableCellCommand
  | ResizeTableColumnCommand
  | ReplaceTableCommand
  | ApplyTablePropsCommand

export function acceptRevision(id: string): AcceptRevisionCommand {
  return {
    kind: 'accept-revision',
    id,
  }
}

export function rejectRevision(id: string): RejectRevisionCommand {
  return {
    kind: 'reject-revision',
    id,
  }
}

export function acceptAllRevisions(): AcceptAllRevisionsCommand {
  return {
    kind: 'accept-all-revisions',
  }
}

export function rejectAllRevisions(): RejectAllRevisionsCommand {
  return {
    kind: 'reject-all-revisions',
  }
}
