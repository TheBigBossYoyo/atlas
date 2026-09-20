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
  /** i18n key (see `src/i18n/messages.en.ts`'s `docx.tableEdit.*` section) —
   * resolved to display text by `TableEditMenuItems.tsx`, the only renderer
   * of this list, so this pure data module stays free of a React/`t()`
   * dependency. */
  readonly labelKey: string
  /** Renders a divider line above this item — groups insert/merge/delete
   * visually without needing a separate "section" data structure. */
  readonly dividerBefore?: boolean
}

export const TABLE_EDIT_ACTIONS: ReadonlyArray<TableEditAction> = [
  { kind: 'insert-table-row-above', labelKey: 'docx.tableEdit.insertRowAbove' },
  { kind: 'insert-table-row-below', labelKey: 'docx.tableEdit.insertRowBelow' },
  { kind: 'insert-table-column-left', labelKey: 'docx.tableEdit.insertColumnLeft' },
  { kind: 'insert-table-column-right', labelKey: 'docx.tableEdit.insertColumnRight' },
  { kind: 'merge-table-cell-right', labelKey: 'docx.tableEdit.mergeRight', dividerBefore: true },
  { kind: 'split-table-cell', labelKey: 'docx.tableEdit.splitCell' },
  { kind: 'delete-table-row', labelKey: 'docx.tableEdit.deleteRow', dividerBefore: true },
  { kind: 'delete-table-column', labelKey: 'docx.tableEdit.deleteColumn' },
  { kind: 'delete-table', labelKey: 'docx.tableEdit.deleteTable' },
]
