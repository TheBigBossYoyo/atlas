/**
 * glide-data-grid's `DataEditor` as used by SpreadsheetViewer and CsvViewer,
 * loaded lazily by both (this module statically pulls in the grid and its CSS).
 *
 * USR-17 — cell edits typed quickly were silently dropped: the grid's overlay
 * commits Enter/Tab with a value captured from its previous render, so a key
 * typed just before Enter (or the whole edit, when the overlay had not
 * re-rendered yet) never reached `onCellEdited`. The editor below swaps in
 * the same text entry with its own Enter/Tab handler that commits the
 * textarea's live value directly. Escape and click-outside keep the grid's
 * own behavior.
 */
import type { ChangeEvent, KeyboardEvent } from 'react'
import {
  DataEditor,
  GridCellKind,
  TextCellEntry,
  type DataEditorProps,
  type GridCell,
  type ProvideEditorCallback,
  type ProvideEditorComponent,
  type TextCell,
} from '@glideapps/glide-data-grid'
import '@glideapps/glide-data-grid/dist/index.css'

const TextCellEditor: ProvideEditorComponent<TextCell> = ({ value, onChange, onFinishedEditing, isHighlighted, validatedSelection }) => {
  const handleKeyDown = (event: KeyboardEvent<HTMLTextAreaElement>): void => {
    const isCommitKey = (event.key === 'Enter' && !event.shiftKey) || event.key === 'Tab'
    if (!isCommitKey) return
    event.preventDefault()
    event.stopPropagation()
    const movement = event.key === 'Tab' ? ([event.shiftKey ? -1 : 1, 0] as const) : ([0, 1] as const)
    onFinishedEditing({ ...value, data: event.currentTarget.value }, movement)
  }

  return (
    <TextCellEntry
      highlight={isHighlighted}
      autoFocus={value.readonly !== true}
      disabled={value.readonly === true}
      altNewline
      value={value.data}
      validatedSelection={validatedSelection}
      onChange={(event: ChangeEvent<HTMLTextAreaElement>) => onChange({ ...value, data: event.target.value })}
      onKeyDown={handleKeyDown}
    />
  )
}

const provideCellEditor: ProvideEditorCallback<GridCell> = (cell) =>
  cell.kind === GridCellKind.Text ? (TextCellEditor as ProvideEditorComponent<GridCell>) : undefined

export function SpreadsheetDataEditor(props: DataEditorProps) {
  return <DataEditor provideEditor={provideCellEditor} {...props} />
}
