/**
 * Spreadsheet editing + save (wave 3) — CsvViewer/TsvViewer cell edit,
 * row/column insert-delete, undo/redo, and an edit -> save -> reload round
 * trip through the real viewer for both csv and tsv.
 */
import { act, fireEvent, render, screen, waitFor } from '@testing-library/react'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { CompactSelection, type EditableGridCell, type GridSelection, type Item } from '@glideapps/glide-data-grid'

import type { LoadedFile } from '../../formats/types'
import { ShortcutManagerProvider } from '../../hooks/ShortcutManagerProvider'
import { ViewerProvider } from '../shared/ViewerContext'
import { CsvViewer } from '../CsvViewer'
import { TsvViewer } from '../TsvViewer'

type CapturedCell = { readonly displayData: string }
type CapturedProps = {
  readonly getCellContent: (loc: readonly [number, number]) => CapturedCell
  readonly columns: ReadonlyArray<{ readonly title: string }>
  readonly rows: number
  readonly onCellEdited?: (cell: Item, newValue: EditableGridCell) => void
  readonly onGridSelectionChange?: (selection: GridSelection) => void
}

let lastDataEditorProps: CapturedProps | null = null

vi.mock('@glideapps/glide-data-grid', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@glideapps/glide-data-grid')>()
  return {
    ...actual,
    DataEditor: (props: CapturedProps) => {
      lastDataEditorProps = props
      return null
    },
  }
})

function gridRows(props: CapturedProps): string[][] {
  const out: string[][] = []
  for (let r = 0; r < props.rows; r++) {
    const row: string[] = []
    for (let c = 0; c < props.columns.length; c++) {
      row.push(props.getCellContent([c, r]).displayData)
    }
    out.push(row)
  }
  return trimBlankMargin(out)
}

/** Drops the blank rows/columns the viewer draws past the data (USR-17's Excel-like margin). */
function trimBlankMargin(grid: string[][]): string[][] {
  let rowCount = grid.length
  while (rowCount > 0 && grid[rowCount - 1].every((cell) => cell === '')) rowCount--
  const rows = grid.slice(0, rowCount)
  const colCount = rows.reduce((max, row) => {
    let last = row.length
    while (last > 0 && row[last - 1] === '') last--
    return Math.max(max, last)
  }, 0)
  return rows.map((row) => row.slice(0, colCount))
}

function editCell(col: number, row: number, text: string): void {
  const newValue = { kind: 'text', data: text, displayData: text, allowOverlay: true } as EditableGridCell
  lastDataEditorProps!.onCellEdited?.([col, row], newValue)
}

beforeEach(() => {
  lastDataEditorProps = null
  window.electronAPI = {
    getInitialFile: vi.fn(),
    openFileDialog: vi.fn(),
    openFileByPath: vi.fn(),
    saveFile: vi.fn().mockResolvedValue({ saved: true, path: '/tmp/saved.csv', name: 'saved.csv' }),
    saveBinaryFile: vi.fn(),
    onFileOpened: vi.fn(),
    setTheme: vi.fn(),
    openFileBinary: vi.fn(),
    readBinaryByPath: vi.fn(),
    onFileOpenedPath: vi.fn(),
    getPathForFile: vi.fn(),
    registerDroppedPath: vi.fn(),
    requestOpenRecent: vi.fn(),
    revealInFolder: vi.fn(),
    spellcheck: {
      onContextMenu: vi.fn(),
      replaceMisspelling: vi.fn(),
      addWord: vi.fn(),
      getLanguages: vi.fn(),
      setLanguages: vi.fn(),
    },
  } as unknown as typeof window.electronAPI
})

describe('CsvViewer — cell editing', () => {
  it('a committed cell edit updates the grid', async () => {
    const file: LoadedFile = { kind: 'text', content: 'name,value\nAtlas,2\n', path: '/tmp/sample.csv', format: 'csv' }
    render(
      <ViewerProvider filePath={file.path}>
        <CsvViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    act(() => editCell(0, 1, 'Renamed'))

    await waitFor(() => expect(gridRows(lastDataEditorProps!)[1][0]).toBe('Renamed'))
  })
})

describe('CsvViewer — save round trip', () => {
  it('edit -> save -> reload round trip preserves the edit (csv)', async () => {
    const file: LoadedFile = { kind: 'text', content: 'name,value\nAtlas,2\n', path: '/tmp/sample.csv', format: 'csv' }
    render(
      <ViewerProvider filePath={file.path}>
        <CsvViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    act(() => editCell(1, 1, '99'))
    await waitFor(() => expect(gridRows(lastDataEditorProps!)[1][1]).toBe('99'))

    const saveButton = await screen.findByRole('button', { name: 'Save' })
    await act(async () => {
      fireEvent.click(saveButton)
    })

    expect(window.electronAPI!.saveFile).toHaveBeenCalledWith(
      // SHEET-8 — the file ended with a line break, so the save keeps one (this
      // used to assert it was dropped: the bug). No encoding metadata is
      // recorded in this component test, so the historical CRLF output applies.
      expect.objectContaining({ content: 'name,value\r\nAtlas,99\r\n', suggestedName: 'sample.csv', existingPath: '/tmp/sample.csv' }),
    )

    const savedContent = (window.electronAPI!.saveFile as ReturnType<typeof vi.fn>).mock.calls[0][0].content as string

    lastDataEditorProps = null
    const reopened: LoadedFile = { kind: 'text', content: savedContent, path: '/tmp/reopened.csv', format: 'csv' }
    render(
      <ViewerProvider filePath={reopened.path}>
        <CsvViewer file={reopened} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())
    expect(gridRows(lastDataEditorProps!)).toEqual([
      ['name', 'value'],
      ['Atlas', '99'],
    ])
  })

  it('edit -> save -> reload round trip preserves the edit (tsv)', async () => {
    const file: LoadedFile = { kind: 'text', content: 'name\tvalue\nAtlas\t2\n', path: '/tmp/sample.tsv', format: 'tsv' }
    render(
      <ViewerProvider filePath={file.path}>
        <TsvViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    act(() => editCell(1, 1, '42'))
    await waitFor(() => expect(gridRows(lastDataEditorProps!)[1][1]).toBe('42'))

    const saveButton = await screen.findByRole('button', { name: 'Save' })
    await act(async () => {
      fireEvent.click(saveButton)
    })

    expect(window.electronAPI!.saveFile).toHaveBeenCalledWith(
      expect.objectContaining({ content: 'name\tvalue\r\nAtlas\t42\r\n' }),
    )

    const savedContent = (window.electronAPI!.saveFile as ReturnType<typeof vi.fn>).mock.calls[0][0].content as string

    lastDataEditorProps = null
    const reopened: LoadedFile = { kind: 'text', content: savedContent, path: '/tmp/reopened.tsv', format: 'tsv' }
    render(
      <ViewerProvider filePath={reopened.path}>
        <TsvViewer file={reopened} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())
    expect(gridRows(lastDataEditorProps!)).toEqual([
      ['name', 'value'],
      ['Atlas', '42'],
    ])
  })
})

describe('CsvViewer — search filter + editing (row-index mapping)', () => {
  it('edits the correct underlying row when the search filter has dropped other rows', async () => {
    const file: LoadedFile = { kind: 'text', content: 'a\nb\nMATCH\n', path: '/tmp/sample.csv', format: 'csv' }
    render(
      <ViewerProvider filePath={file.path}>
        <CsvViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    fireEvent.change(screen.getByPlaceholderText('Search rows...'), { target: { value: 'MATCH' } })
    await waitFor(() => expect(lastDataEditorProps!.rows).toBe(1))

    // The only visible row is grid-space row 0 — before the fix this landed
    // on the actual row 0 ("a"), not the matched row 2 ("MATCH").
    act(() => editCell(0, 0, 'Edited'))

    fireEvent.change(screen.getByPlaceholderText('Search rows...'), { target: { value: '' } })
    await waitFor(() => expect(gridRows(lastDataEditorProps!)).toEqual([['a'], ['b'], ['Edited']]))
  })
})

describe('CsvViewer — row insert/delete and undo', () => {
  it('inserts and then undoes a row', async () => {
    const file: LoadedFile = { kind: 'text', content: 'a\nb\nc\n', path: '/tmp/sample.csv', format: 'csv' }
    render(
      <ViewerProvider filePath={file.path}>
        <CsvViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    // Select the second row (index 1) via the grid's selection callback,
    // reusing the same shape SpreadsheetViewer's own tests exercise.
    act(() => {
      lastDataEditorProps!.onGridSelectionChange?.({
        current: { cell: [0, 1], range: { x: 0, y: 1, width: 1, height: 1 }, rangeStack: [] },
        columns: CompactSelection.empty(),
        rows: CompactSelection.empty(),
      })
    })

    const insertButton = await screen.findByRole('button', { name: 'Insert row above' })
    await waitFor(() => expect(insertButton).not.toBeDisabled())
    fireEvent.click(insertButton)
    await waitFor(() => expect(gridRows(lastDataEditorProps!)).toEqual([['a'], [''], ['b'], ['c']]))

    const undoButton = await screen.findByRole('button', { name: 'Undo' })
    fireEvent.click(undoButton)
    await waitFor(() => expect(gridRows(lastDataEditorProps!)).toEqual([['a'], ['b'], ['c']]))
  })

  // Regression: the shared `useSpreadsheetEditor` hook's undo/redo was only
  // reachable via the toolbar buttons, not the standard Ctrl+Z/Ctrl+Y keys —
  // see the identical test in SpreadsheetViewer.editing.test.tsx, which this
  // mirrors to confirm the fix (added directly to the shared hook) also
  // covers CsvViewer.
  it('Ctrl+Z / Ctrl+Y perform undo/redo, matching the toolbar buttons', async () => {
    const file: LoadedFile = { kind: 'text', content: 'a\nb\nc\n', path: '/tmp/sample.csv', format: 'csv' }
    render(
      <ShortcutManagerProvider>
        <ViewerProvider filePath={file.path}>
          <CsvViewer file={file} />
        </ViewerProvider>
      </ShortcutManagerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    act(() => editCell(0, 0, 'changed'))
    await waitFor(() => expect(gridRows(lastDataEditorProps!)[0][0]).toBe('changed'))

    const ctrlZ = new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true, cancelable: true })
    act(() => {
      window.dispatchEvent(ctrlZ)
    })
    expect(ctrlZ.defaultPrevented).toBe(true)
    await waitFor(() => expect(gridRows(lastDataEditorProps!)[0][0]).toBe('a'))

    const ctrlY = new KeyboardEvent('keydown', { key: 'y', ctrlKey: true, bubbles: true, cancelable: true })
    act(() => {
      window.dispatchEvent(ctrlY)
    })
    expect(ctrlY.defaultPrevented).toBe(true)
    await waitFor(() => expect(gridRows(lastDataEditorProps!)[0][0]).toBe('changed'))
  })
})
