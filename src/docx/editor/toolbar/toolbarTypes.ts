export type ToolbarCommand =
  | { kind: 'set-font-family'; family: string }
  | { kind: 'set-font-size'; sizePt: number }
  | { kind: 'toggle-bold' | 'toggle-italic' | 'toggle-underline' | 'toggle-strike' | 'toggle-subscript' | 'toggle-superscript' }
  | { kind: 'set-font-color' | 'set-highlight-color'; colorHex: string }
  | { kind: 'set-alignment'; align: 'left' | 'center' | 'right' | 'justify' }
  | { kind: 'set-line-spacing'; spacing: 1 | 1.15 | 1.5 | 2 }
  | { kind: 'toggle-bullet-list' | 'toggle-numbered-list' }
  | { kind: 'change-indent'; delta: 1 | -1 }
  | { kind: 'apply-style'; styleId: string }
  | { kind: 'insert-table'; rows: number; cols: number }
  | { kind: 'insert-image' | 'insert-hyperlink' | 'insert-header' | 'insert-footer' | 'insert-page-break' | 'insert-comment' }
  | {
      // DXE-14 — resolved against the cursor's current cell (see
      // toolbarAdapter.ts's `resolveTableCommand`); disabled in the UI
      // whenever `ToolbarState.insideTable` is false.
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
  | { kind: 'set-margins'; preset: 'normal' | 'narrow' | 'moderate' | 'wide' }
  | { kind: 'set-orientation'; orientation: 'portrait' | 'landscape' }
  | { kind: 'set-page-size'; preset: 'letter' | 'a4' | 'legal' }
  | { kind: 'set-columns'; count: 1 | 2 | 3 }
  | {
      kind:
        | 'toggle-spell-check'
        | 'toggle-track-changes'
        | 'accept-change'
        | 'reject-change'
        | 'accept-all-changes'
        | 'reject-all-changes'
        | 'open-comments-pane'
    }
  | { kind: 'undo' | 'redo' }
  | { kind: 'open-find-replace' };

export type ToolbarState = Readonly<{
  activeFormats: ReadonlySet<'bold' | 'italic' | 'underline' | 'strike' | 'subscript' | 'superscript'>;
  alignment: 'left' | 'center' | 'right' | 'justify' | null;
  fontFamily: string | null;
  fontSizePt: number | null;
  styleId: string | null;
  trackChanges: boolean;
  spellCheck: boolean;
  /** DXE-14 — true when the cursor sits inside a table cell; gates the
   * table-editing button group. */
  insideTable: boolean;
}>;
