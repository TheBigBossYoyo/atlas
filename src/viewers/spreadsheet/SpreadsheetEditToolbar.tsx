/**
 * Shared editing toolbar segment for SpreadsheetViewer and CsvViewer (wave 3).
 *
 * Purely presentational + event plumbing — every action is a callback prop,
 * so this component has no opinion on what "insert a row" actually means to
 * the document model. Row/column actions operate on `selection` (the
 * currently active cell, tracked by the parent viewer — see that file's
 * comment on why selection tracking lives there and not in `useGridFind`);
 * they're disabled when nothing is selected rather than guessing a target.
 */
import { useState } from 'react'
import { ArrowLeftToLine, ArrowRightToLine, ArrowUpToLine, ArrowDownToLine, ClipboardPaste, Redo2, Save, Trash2, Undo2 } from 'lucide-react'

import { tsvToRows } from './spreadsheetClipboard'

export type SaveFormatOption = {
  readonly id: string
  readonly label: string
}

export type SpreadsheetEditToolbarProps = {
  readonly canUndo: boolean
  readonly canRedo: boolean
  readonly onUndo: () => void
  readonly onRedo: () => void
  readonly selection: { readonly row: number; readonly col: number } | null
  readonly onInsertRowAbove: (row: number) => void
  readonly onDeleteRow: (row: number) => void
  readonly onInsertColumnLeft: (col: number) => void
  readonly onDeleteColumn: (col: number) => void
  readonly onPaste: (row: number, col: number, values: ReadonlyArray<ReadonlyArray<string>>) => void
  readonly onSave: () => void
  readonly saveFormats: ReadonlyArray<SaveFormatOption>
  readonly onSaveAs: (formatId: string) => void
}

/**
 * Reads the OS clipboard as TSV/plain text via the async Clipboard API
 * (user-gesture-gated — this is only ever called from a click handler), then
 * hands it to `spreadsheetClipboard.ts`'s own parser — the single source of
 * truth for TSV parsing, also used directly by its unit tests — rather than
 * re-implementing the same tab/newline splitting here.
 */
async function readClipboardRows(): Promise<string[][] | null> {
  try {
    const text = await navigator.clipboard.readText()
    // An empty clipboard is treated as "nothing to paste" (a no-op), not as
    // `tsvToRows`'s own "one empty cell" reading of an empty string — that
    // reading exists for a genuinely-empty-but-copied cell, not for a
    // clipboard with nothing in it at all.
    if (!text) return null
    return tsvToRows(text)
  } catch {
    // Clipboard read denied/unavailable — the grid's own built-in Ctrl+V
    // paste handling (glide-data-grid's `onPaste`) still works either way;
    // this button is a discoverable extra, not the only path.
    return null
  }
}

export function SpreadsheetEditToolbar({
  canUndo,
  canRedo,
  onUndo,
  onRedo,
  selection,
  onInsertRowAbove,
  onDeleteRow,
  onInsertColumnLeft,
  onDeleteColumn,
  onPaste,
  onSave,
  saveFormats,
  onSaveAs,
}: SpreadsheetEditToolbarProps) {
  const [saveFormat, setSaveFormat] = useState(saveFormats[0]?.id ?? '')

  const handlePasteClick = (): void => {
    void readClipboardRows().then((rows) => {
      if (rows) onPaste(selection?.row ?? 0, selection?.col ?? 0, rows)
    })
  }

  return (
    <div className="spreadsheet-edit-toolbar">
      <div className="spreadsheet-edit-toolbar__group">
        <button type="button" onClick={onUndo} disabled={!canUndo} title="Undo (Ctrl+Z)" aria-label="Undo">
          <Undo2 size={14} />
        </button>
        <button type="button" onClick={onRedo} disabled={!canRedo} title="Redo (Ctrl+Y)" aria-label="Redo">
          <Redo2 size={14} />
        </button>
      </div>

      <div className="spreadsheet-edit-toolbar__group">
        <button
          type="button"
          onClick={() => selection && onInsertRowAbove(selection.row)}
          disabled={!selection}
          title="Insert row above the selected cell"
          aria-label="Insert row above"
        >
          <ArrowUpToLine size={14} />
        </button>
        <button
          type="button"
          onClick={() => selection && onDeleteRow(selection.row)}
          disabled={!selection}
          title="Delete the selected row"
          aria-label="Delete row"
        >
          <ArrowDownToLine size={14} />
          <Trash2 size={10} />
        </button>
        <button
          type="button"
          onClick={() => selection && onInsertColumnLeft(selection.col)}
          disabled={!selection}
          title="Insert column left of the selected cell"
          aria-label="Insert column left"
        >
          <ArrowLeftToLine size={14} />
        </button>
        <button
          type="button"
          onClick={() => selection && onDeleteColumn(selection.col)}
          disabled={!selection}
          title="Delete the selected column"
          aria-label="Delete column"
        >
          <ArrowRightToLine size={14} />
          <Trash2 size={10} />
        </button>
        <button type="button" onClick={handlePasteClick} title="Paste from clipboard" aria-label="Paste">
          <ClipboardPaste size={14} />
        </button>
      </div>

      <div className="spreadsheet-edit-toolbar__group spreadsheet-edit-toolbar__save">
        <button type="button" onClick={onSave} title="Save (Ctrl+S)" aria-label="Save">
          <Save size={14} />
          <span>Save</span>
        </button>
        {saveFormats.length > 1 && (
          <>
            <select
              aria-label="Save As format"
              value={saveFormat}
              onChange={(e) => setSaveFormat(e.target.value)}
            >
              {saveFormats.map((f) => (
                <option key={f.id} value={f.id}>
                  {f.label}
                </option>
              ))}
            </select>
            <button
              type="button"
              onClick={() => onSaveAs(saveFormat)}
              title="Save As…"
              aria-label="Save As"
            >
              Save As…
            </button>
          </>
        )}
      </div>
    </div>
  )
}
