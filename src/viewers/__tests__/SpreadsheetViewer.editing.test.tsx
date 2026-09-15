/**
 * Spreadsheet editing + save (wave 3) — cell edit, row/column insert-delete,
 * sheet add/rename/delete, undo/redo, frozen rows, and an edit -> save ->
 * reload round trip through the real viewer (not just the underlying
 * document-model unit tests in src/viewers/spreadsheet/__tests__/).
 */
import { readFileSync } from 'node:fs'
import path from 'node:path'
import { act, fireEvent, render, screen, waitFor } from '@testing-library/react'
import JSZip from 'jszip'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import * as XLSX from 'xlsx'
import { CompactSelection, type EditableGridCell, type GridSelection, type Item } from '@glideapps/glide-data-grid'

import type { LoadedFile } from '../../formats/types'
import { ShortcutManagerProvider } from '../../hooks/ShortcutManagerProvider'
import { ViewerProvider } from '../shared/ViewerContext'
import { SpreadsheetViewer } from '../SpreadsheetViewer'
import { OdsViewer } from '../OdsViewer'

type CapturedCell = {
  readonly data: string
  readonly displayData: string
  readonly allowOverlay?: boolean
  readonly themeOverride?: { readonly baseFontStyle?: string }
}
type CapturedProps = {
  readonly getCellContent: (loc: readonly [number, number]) => CapturedCell
  readonly columns: ReadonlyArray<{ readonly title: string }>
  readonly rows: number
  readonly onCellEdited?: (cell: Item, newValue: EditableGridCell) => void
  readonly onPaste?: (target: Item, values: readonly (readonly string[])[]) => boolean
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

const FIXTURES_DIR = path.resolve(process.cwd(), 'src/viewers/__fixtures__')

function readFixtureArrayBuffer(name: string): ArrayBuffer {
  const buf = readFileSync(path.join(FIXTURES_DIR, name))
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer
}

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

/** Simulates the user clicking a cell, for tests exercising toolbar actions that read the current selection. */
function selectCell(col: number, row: number): void {
  const selection: GridSelection = {
    current: { cell: [col, row], range: { x: col, y: row, width: 1, height: 1 }, rangeStack: [] },
    columns: CompactSelection.empty(),
    rows: CompactSelection.empty(),
  }
  lastDataEditorProps!.onGridSelectionChange?.(selection)
}

function buildWorkbookFile(build: (wb: XLSX.WorkBook) => void, path_ = '/tmp/generated.xlsx'): LoadedFile {
  const wb = XLSX.utils.book_new()
  build(wb)
  const content = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer
  return { kind: 'binary', content, path: path_, format: 'xlsx' }
}

beforeEach(() => {
  lastDataEditorProps = null
  window.electronAPI = {
    getInitialFile: vi.fn(),
    openFileDialog: vi.fn(),
    openFileByPath: vi.fn(),
    saveFile: vi.fn().mockResolvedValue({ saved: true, path: '/tmp/saved.xlsx', name: 'saved.xlsx' }),
    saveBinaryFile: vi.fn().mockResolvedValue({ saved: true, path: '/tmp/saved.xlsx', name: 'saved.xlsx' }),
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

describe('SpreadsheetViewer — cell editing', () => {
  it('a committed cell edit updates the grid', async () => {
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Name', 'Score'], ['Alice', 10]]), 'Sheet1')
    })
    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    act(() => editCell(0, 1, 'Bob'))

    await waitFor(() => expect(gridRows(lastDataEditorProps!)[1][0]).toBe('Bob'))
  })

  it('draws blank rows/columns past the data and typing into one grows the sheet (USR-17)', async () => {
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Name', 'Score'], ['Alice', 10]]), 'Sheet1')
    })
    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())
    expect(lastDataEditorProps!.rows).toBeGreaterThanOrEqual(100)
    expect(lastDataEditorProps!.columns.length).toBeGreaterThanOrEqual(26)
    expect(lastDataEditorProps!.getCellContent([4, 5]).allowOverlay).toBe(true)

    act(() => editCell(3, 4, 'far'))

    await waitFor(() =>
      expect(gridRows(lastDataEditorProps!)).toEqual([
        ['Name', 'Score', '', ''],
        ['Alice', '10', '', ''],
        ['', '', '', ''],
        ['', '', '', ''],
        ['', '', '', 'far'],
      ]),
    )
    expect(screen.getByText('5 rows × 4 columns')).toBeInTheDocument()
  })

  it('draws an Excel table header row as a header (USR-17)', async () => {
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Name', 'Qty'], ['Apple', 3]]), 'Data')
    const zip = await JSZip.loadAsync(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer)
    zip.file(
      'xl/tables/table1.xml',
      '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:B2"><tableColumns count="2"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Qty"/></tableColumns></table>',
    )
    zip.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
    )
    const file: LoadedFile = {
      kind: 'binary',
      content: await zip.generateAsync({ type: 'arraybuffer' }),
      path: '/tmp/table.xlsx',
      format: 'xlsx',
    }
    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps!.getCellContent([0, 0]).themeOverride?.baseFontStyle).toMatch(/^600/))
    expect(lastDataEditorProps!.getCellContent([0, 1]).themeOverride?.baseFontStyle).toBeUndefined()
    expect(lastDataEditorProps!.getCellContent([2, 0]).themeOverride?.baseFontStyle).toBeUndefined()
  })

  it('a formula edit shows its computed value', async () => {
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a', 'b', 'sum'], [1, 2, '']]), 'Sheet1')
    })
    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    act(() => editCell(2, 1, '=A2+B2'))

    await waitFor(() => expect(gridRows(lastDataEditorProps!)[1][2]).toBe('3'))
  })

  it('re-opening and committing a formula cell UNCHANGED keeps it a live formula, not a frozen value', async () => {
    // Confirmed data-loss bug this locks in: glide-data-grid's text-cell
    // editor seeds its input from getCellContent's `data` field, not
    // `displayData` (see useSpreadsheetGrid.ts's header). Before that fix,
    // `data` held the already-computed "3", so simply opening this cell and
    // committing with no changes (the most innocuous possible interaction)
    // silently replaced the formula with the static value "3" — invisible
    // in the grid until a dependency changed and the cell no longer updated.
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a', 'b', 'sum'], [1, 2, '']]), 'Sheet1')
    })
    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    act(() => editCell(2, 1, '=A2+B2'))
    await waitFor(() => expect(gridRows(lastDataEditorProps!)[1][2]).toBe('3'))

    // Simulate re-opening the formula cell and committing with NO edits:
    // read back whatever the edit overlay would have been seeded with (the
    // real glide-data-grid editor reads `data`, exercised directly here
    // since DataEditor itself is mocked in this file) and feed it straight
    // back through onCellEdited, exactly like an untouched Enter/blur would.
    const seeded = lastDataEditorProps!.getCellContent([2, 1])
    act(() => editCell(2, 1, seeded.data))
    await waitFor(() => expect(gridRows(lastDataEditorProps!)[1][2]).toBe('3'))

    // The real assertion: the formula must still be LIVE. Changing a
    // dependency (A2) should still flow through to the sum — which would NOT
    // happen if the previous step had frozen "sum" into the plain value "3".
    act(() => editCell(0, 1, '10'))
    await waitFor(() => expect(gridRows(lastDataEditorProps!)[1][2]).toBe('12'))
  })
})

describe('SpreadsheetViewer — save', () => {
  it('Save writes workbook bytes via saveBinaryFile with the existing path', async () => {
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a']]), 'Sheet1')
    }, '/tmp/existing.xlsx')
    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    act(() => editCell(0, 0, 'changed'))
    await waitFor(() => expect(gridRows(lastDataEditorProps!)[0][0]).toBe('changed'))

    const saveButton = await screen.findByRole('button', { name: 'Save' })
    await act(async () => {
      fireEvent.click(saveButton)
    })

    await waitFor(() => expect(window.electronAPI!.saveBinaryFile).toHaveBeenCalled())
    const call = (window.electronAPI!.saveBinaryFile as ReturnType<typeof vi.fn>).mock.calls[0][0]
    expect(call.existingPath).toBe('/tmp/existing.xlsx')
    expect(call.content).toBeInstanceOf(Uint8Array)
  })

  it('edit -> save -> reload round trip preserves the edit (xlsx)', async () => {
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Name'], ['Alice']]), 'Sheet1')
    })
    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    act(() => editCell(0, 1, 'Zoe'))
    await waitFor(() => expect(gridRows(lastDataEditorProps!)[1][0]).toBe('Zoe'))

    const saveButton = await screen.findByRole('button', { name: 'Save' })
    await act(async () => {
      fireEvent.click(saveButton)
    })

    const savedBytes = (window.electronAPI!.saveBinaryFile as ReturnType<typeof vi.fn>).mock.calls[0][0]
      .content as Uint8Array

    lastDataEditorProps = null
    const reopened: LoadedFile = { kind: 'binary', content: savedBytes.buffer as ArrayBuffer, path: file.path, format: 'xlsx' }
    render(
      <ViewerProvider filePath="/tmp/reopened.xlsx">
        <SpreadsheetViewer file={reopened} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())
    expect(gridRows(lastDataEditorProps!)).toEqual([['Name'], ['Zoe']])
  })

  it('edit -> save -> reload round trip preserves the edit (ods)', async () => {
    const original: LoadedFile = {
      kind: 'binary',
      content: readFixtureArrayBuffer('sample.ods'),
      path: '/tmp/sample.ods',
      format: 'ods',
    }
    render(
      <ViewerProvider filePath={original.path}>
        <OdsViewer file={original} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    act(() => editCell(0, 0, 'Edited'))
    await waitFor(() => expect(gridRows(lastDataEditorProps!)[0][0]).toBe('Edited'))

    const saveButton = await screen.findByRole('button', { name: 'Save' })
    await act(async () => {
      fireEvent.click(saveButton)
    })

    const savedBytes = (window.electronAPI!.saveBinaryFile as ReturnType<typeof vi.fn>).mock.calls[0][0]
      .content as Uint8Array

    lastDataEditorProps = null
    const reopened: LoadedFile = { kind: 'binary', content: savedBytes.buffer as ArrayBuffer, path: '/tmp/reopened.ods', format: 'ods' }
    render(
      <ViewerProvider filePath={reopened.path}>
        <OdsViewer file={reopened} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())
    expect(gridRows(lastDataEditorProps!)[0][0]).toBe('Edited')
  })
})

describe('SpreadsheetViewer — legacy formats are "view + save-as-xlsx only" (plan scope)', () => {
  it('defaults Save to .xlsx (not the original .xls) and forces the dialog on the first save', async () => {
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a']]), 'Sheet1')
    }, '/tmp/legacy.xls')
    // format:'xlsx' matches how extensionManifest.ts routes a real .xls
    // file — SpreadsheetViewer itself derives the *actual* extension from
    // file.path, not file.format (see extensionOf/fileExtension).
    const legacyFile: LoadedFile = { ...file, path: '/tmp/legacy.xls' }

    render(
      <ViewerProvider filePath={legacyFile.path}>
        <SpreadsheetViewer file={legacyFile} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    act(() => editCell(0, 0, 'changed'))
    const saveButton = await screen.findByRole('button', { name: 'Save' })
    await act(async () => {
      fireEvent.click(saveButton)
    })

    const call = (window.electronAPI!.saveBinaryFile as ReturnType<typeof vi.fn>).mock.calls[0][0]
    // No existingPath: main process shows the save dialog rather than
    // silently overwriting legacy.xls with re-encoded (xlsx-shaped) bytes.
    expect(call.existingPath).toBeUndefined()
    expect(call.suggestedName).toBe('legacy.xlsx')
    expect(call.filters).toEqual([{ name: 'Excel Workbook', extensions: ['xlsx'] }])
  })

  it('once saved, a second plain Save overwrites the NEW path (not legacy.xls) in place', async () => {
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a']]), 'Sheet1')
    })
    const legacyFile: LoadedFile = { ...file, path: '/tmp/legacy.xls' }

    render(
      <ViewerProvider filePath={legacyFile.path}>
        <SpreadsheetViewer file={legacyFile} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    const saveButton = await screen.findByRole('button', { name: 'Save' })
    await act(async () => {
      fireEvent.click(saveButton) // first save -> dialog resolves to /tmp/saved.xlsx (mocked)
    })
    await act(async () => {
      fireEvent.click(saveButton) // second save -> should now overwrite the chosen path
    })

    const calls = (window.electronAPI!.saveBinaryFile as ReturnType<typeof vi.fn>).mock.calls
    expect(calls[1][0].existingPath).toBe('/tmp/saved.xlsx')
  })
})

describe('SpreadsheetViewer — row/column insert and delete', () => {
  it('inserts a row above the selected cell', async () => {
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a'], ['b'], ['c']]), 'Sheet1')
    })
    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    act(() => selectCell(0, 1))

    const insertRowButton = await screen.findByRole('button', { name: 'Insert row above' })
    await waitFor(() => expect(insertRowButton).not.toBeDisabled())
    fireEvent.click(insertRowButton)

    await waitFor(() => expect(gridRows(lastDataEditorProps!)).toEqual([['a'], [''], ['b'], ['c']]))
  })

  it('deletes the selected row', async () => {
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a'], ['b'], ['c']]), 'Sheet1')
    })
    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    act(() => selectCell(0, 1))

    const deleteRowButton = await screen.findByRole('button', { name: 'Delete row' })
    await waitFor(() => expect(deleteRowButton).not.toBeDisabled())
    fireEvent.click(deleteRowButton)

    await waitFor(() => expect(gridRows(lastDataEditorProps!)).toEqual([['a'], ['c']]))
  })
})

describe('SpreadsheetViewer — sheet management', () => {
  it('adds a new sheet via the + tab button', async () => {
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a']]), 'Sheet1')
    })
    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    fireEvent.click(screen.getByRole('button', { name: 'Add sheet' }))

    expect(await screen.findByText('Sheet2')).toBeInTheDocument()
  })

  it('renames a sheet via double-click', async () => {
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a']]), 'Sheet1')
    })
    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    fireEvent.doubleClick(screen.getByRole('button', { name: /Sheet1/ }))
    const input = screen.getByLabelText('Rename sheet Sheet1')
    fireEvent.change(input, { target: { value: 'Budget' } })
    fireEvent.blur(input)

    expect(await screen.findByText('Budget')).toBeInTheDocument()
  })

  it('deletes a sheet via its delete button (only when more than one sheet exists)', async () => {
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a']]), 'Sheet1')
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['b']]), 'Sheet2')
    })
    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())
    expect(screen.getByText('Sheet2')).toBeInTheDocument()

    fireEvent.click(screen.getByRole('button', { name: 'Delete sheet Sheet2' }))

    await waitFor(() => expect(screen.queryByText('Sheet2')).not.toBeInTheDocument())
  })

  // Regression: `commitRename` never told `activeSheetName` about the new
  // name, so renaming a sheet other than the FIRST one left that piece of
  // state pointing at a name that no longer existed. The render-time
  // "pick a fallback active sheet" adjustment then silently jumped to
  // `visibleSheets[0]` — switching the visible grid to a completely
  // different sheet than the one the user just renamed. The original rename
  // test above only covered a single-sheet workbook, where sheet 0 IS the
  // renamed sheet, masking the bug entirely.
  it('keeps the renamed sheet active (and its own content visible) when it is not the first tab', async () => {
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a']]), 'Sheet1')
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['b']]), 'Sheet2')
    })
    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    fireEvent.click(screen.getByRole('button', { name: 'Sheet2' }))
    await waitFor(() => expect(gridRows(lastDataEditorProps!)[0][0]).toBe('b'))

    fireEvent.doubleClick(screen.getByRole('button', { name: 'Sheet2' }))
    const input = screen.getByLabelText('Rename sheet Sheet2')
    fireEvent.change(input, { target: { value: 'Budget' } })
    fireEvent.blur(input)

    await waitFor(() => expect(screen.getByRole('button', { name: 'Budget' })).toHaveClass('spreadsheet-viewer__tab--active'))
    expect(gridRows(lastDataEditorProps!)[0][0]).toBe('b')
  })
})

describe('SpreadsheetViewer — undo', () => {
  it('reverts a cell edit', async () => {
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a']]), 'Sheet1')
    })
    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    act(() => editCell(0, 0, 'changed'))
    await waitFor(() => expect(gridRows(lastDataEditorProps!)[0][0]).toBe('changed'))

    const undoButton = await screen.findByRole('button', { name: 'Undo' })
    fireEvent.click(undoButton)

    await waitFor(() => expect(gridRows(lastDataEditorProps!)[0][0]).toBe('a'))
  })

  it('Ctrl+Z / Ctrl+Y perform undo/redo, matching the toolbar buttons', async () => {
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a']]), 'Sheet1')
    })
    render(
      <ShortcutManagerProvider>
        <ViewerProvider filePath={file.path}>
          <SpreadsheetViewer file={file} />
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

describe('SpreadsheetViewer — search filter + editing (row-index mapping)', () => {
  it('edits the correct underlying sheet row when the search filter has dropped other rows', async () => {
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a'], ['b'], ['MATCH']]), 'Sheet1')
    })
    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    fireEvent.change(screen.getByPlaceholderText('Search rows...'), { target: { value: 'MATCH' } })
    await waitFor(() => expect(lastDataEditorProps!.rows).toBe(1))

    // The only visible row is grid-space row 0 — before the fix this landed
    // on the sheet's actual row 0 ("a"), not the matched row 2 ("MATCH").
    act(() => editCell(0, 0, 'Edited'))

    fireEvent.change(screen.getByPlaceholderText('Search rows...'), { target: { value: '' } })
    await waitFor(() => expect(gridRows(lastDataEditorProps!)).toEqual([['a'], ['b'], ['Edited']]))
  })

  it('inserts a row above the correct underlying sheet row via the toolbar when filtered', async () => {
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a'], ['b'], ['MATCH']]), 'Sheet1')
    })
    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    fireEvent.change(screen.getByPlaceholderText('Search rows...'), { target: { value: 'MATCH' } })
    await waitFor(() => expect(lastDataEditorProps!.rows).toBe(1))

    act(() => selectCell(0, 0))
    const insertRowButton = await screen.findByRole('button', { name: 'Insert row above' })
    await waitFor(() => expect(insertRowButton).not.toBeDisabled())
    fireEvent.click(insertRowButton)

    fireEvent.change(screen.getByPlaceholderText('Search rows...'), { target: { value: '' } })
    // Before the fix this inserted above sheet row 0 ("a") instead of the
    // matched row 2 ("MATCH").
    await waitFor(() => expect(gridRows(lastDataEditorProps!)).toEqual([['a'], ['b'], [''], ['MATCH']]))
  })
})

describe('SpreadsheetViewer — frozen rows (T4/DAT-10 remainder)', () => {
  it('excludes the frozen row(s) from the main grid body and renders the frozen-rows strip', async () => {
    // Hand-built OOXML zip with a state="frozen" pane (ySplit=1) — see
    // src/viewers/spreadsheet/__tests__/spreadsheetPanes.test.ts for the
    // same fixture-construction approach and why a plain XLSX.write() can't
    // produce this (SheetJS never writes <pane> either).
    const JSZip = (await import('jszip')).default
    const zip = new JSZip()
    zip.file(
      '[Content_Types].xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
    )
    zip.file(
      '_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
    )
    zip.file(
      'xl/workbook.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
    )
    zip.file(
      'xl/worksheets/sheet1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="str"><v>Header</v></c></row>
    <row r="2"><c r="A2" t="str"><v>Row1</v></c></row>
    <row r="3"><c r="A3" t="str"><v>Row2</v></c></row>
  </sheetData>
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
</worksheet>`,
    )
    const buffer = await zip.generateAsync({ type: 'arraybuffer' })

    const file: LoadedFile = { kind: 'binary', content: buffer, path: '/tmp/frozen.xlsx', format: 'xlsx' }
    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())
    // 3 total rows minus 1 frozen row = 2 in the scrollable body.
    expect(gridRows(lastDataEditorProps!)).toEqual([['Row1'], ['Row2']])
    // The frozen header row itself renders in the separate strip, not lost.
    expect(screen.getByDisplayValue('Header')).toBeInTheDocument()
  })
})
