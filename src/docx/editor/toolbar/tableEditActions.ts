import type { ToolbarCommand } from './toolbarTypes'

/**
 * DXE-14 — the structural table-editing commands, shared between the
 * toolbar's "Table" dropdown (`Toolbar.tsx`) and the right-click context
 * menu rendered inside the DOCX editor (`DocxViewer.tsx`) so the two entry
 * points can never drift on labels or which commands exist. Every entry
 * here resolves against the cursor's current cell — see
 * `toolbarAdapter.ts`'s `resolveTableCommand` — so no extra payload is
 * needed beyond the bare `kind`.
 */
export type TableEditActionKind = Extract<
  ToolbarCommand,
  {
    kind:
      | 'insert-table-row-above'
      | 'insert-table-row-below'
      | 'insert-table-column-left'
      | 'insert-table-column-right'
      | 'delete-table-row'
      | 'delete-table-column'
      | 'delete-table'
      | 'merge-table-cell-right'
      | 'split-table-cell'
  }
>['kind']

export type TableEditAction = {
  readonly kind: TableEditActionKind
  readonly label: string
  /** Renders a divider line above this item — groups insert/merge/delete
   * visually without needing a separate "section" data structure. */
  readonly dividerBefore?: boolean
}

export const TABLE_EDIT_ACTIONS: ReadonlyArray<TableEditAction> = [
  { kind: 'insert-table-row-above', label: 'Insert Row Above' },
  { kind: 'insert-table-row-below', label: 'Insert Row Below' },
  { kind: 'insert-table-column-left', label: 'Insert Column Left' },
  { kind: 'insert-table-column-right', label: 'Insert Column Right' },
  { kind: 'merge-table-cell-right', label: 'Merge Right', dividerBefore: true },
  { kind: 'split-table-cell', label: 'Split Cell' },
  { kind: 'delete-table-row', label: 'Delete Row', dividerBefore: true },
  { kind: 'delete-table-column', label: 'Delete Column' },
  { kind: 'delete-table', label: 'Delete Table' },
]
