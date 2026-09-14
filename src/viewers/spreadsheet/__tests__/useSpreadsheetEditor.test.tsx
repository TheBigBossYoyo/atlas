/**
 * Spreadsheet editing + save (wave 3) — dirty/save/undo-redo session wiring
 * through the shared ViewerContext document-session contract (P1.1), the
 * same contract DocxViewer's own save/dirty plumbing uses.
 */
import { act, render, screen, waitFor } from '@testing-library/react'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { ViewerProvider } from '../../shared/ViewerContext'
import { useViewerIsDirty, useViewerSave } from '../../shared/useViewerContext'
import { createDocument } from '../spreadsheetDocument'
import { useSpreadsheetEditor, type SpreadsheetSaveTarget } from '../useSpreadsheetEditor'
import type { ParsedSheet } from '../../shared/spreadsheetGrid'

const WORKBOOK_TARGET: SpreadsheetSaveTarget = {
  kind: 'workbook',
  bookType: 'xlsx',
  extension: 'xlsx',
  filterName: 'Excel Workbook',
}

const DELIMITED_TARGET: SpreadsheetSaveTarget = {
  kind: 'delimited',
  delimiter: ',',
  extension: 'csv',
  filterName: 'CSV',
}

function sheetFixture(rows: string[][]): ParsedSheet {
  return {
    name: 'Sheet1',
    hidden: false,
    grid: {
      rows,
      colCount: Math.max(...rows.map((r) => r.length), 0),
      merges: [],
      colWidthsPx: [],
      rowHeightsPx: [],
      formulas: rows.map((r) => r.map(() => undefined)),
    },
  }
}

function ViewerDirtyProbe() {
  const isDirty = useViewerIsDirty()
  return <span data-testid="dirty">{String(isDirty)}</span>
}

function ViewerSaveProbe() {
  const save = useViewerSave()
  return (
    <button type="button" onClick={() => void save()}>
      Context Save
    </button>
  )
}

type HarnessProps = { readonly target: SpreadsheetSaveTarget; readonly onReady: (api: ReturnType<typeof useSpreadsheetEditor>) => void }

function Harness({ target, onReady }: HarnessProps) {
  const doc = createDocument([sheetFixture([['a', 'b'], ['1', '2']])])
  const editor = useSpreadsheetEditor(doc, '/tmp/sample.xlsx', target)
  onReady(editor)
  return (
    <>
      <ViewerDirtyProbe />
      <ViewerSaveProbe />
      {editor.saveError && <span data-testid="save-error">{editor.saveError}</span>}
    </>
  )
}

function renderHarness(target: SpreadsheetSaveTarget) {
  let latest: ReturnType<typeof useSpreadsheetEditor> | null = null
  const utils = render(
    <ViewerProvider filePath="/tmp/sample.xlsx">
      <Harness target={target} onReady={(api) => { latest = api }} />
    </ViewerProvider>,
  )
  return { ...utils, getEditor: () => latest! }
}

beforeEach(() => {
  window.electronAPI = {
    getInitialFile: vi.fn(),
    openFileDialog: vi.fn(),
    openFileByPath: vi.fn(),
    saveFile: vi.fn().mockResolvedValue({ saved: true, path: '/tmp/sample.csv', name: 'sample.csv' }),
    saveBinaryFile: vi.fn().mockResolvedValue({ saved: true, path: '/tmp/sample.xlsx', name: 'sample.xlsx' }),
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

afterEach(() => {
  vi.restoreAllMocks()
})

describe('useSpreadsheetEditor — dirty tracking', () => {
  it('starts clean once the initial document hydrates', async () => {
    renderHarness(WORKBOOK_TARGET)
    await waitFor(() => expect(screen.getByTestId('dirty')).toHaveTextContent('false'))
  })

  it('becomes dirty after a cell edit', async () => {
    const { getEditor } = renderHarness(WORKBOOK_TARGET)
    await waitFor(() => expect(screen.getByTestId('dirty')).toHaveTextContent('false'))

    act(() => getEditor().setCellValue(0, 0, 0, 'changed'))

    await waitFor(() => expect(screen.getByTestId('dirty')).toHaveTextContent('true'))
  })

  it('becomes clean again after a successful save', async () => {
    const { getEditor } = renderHarness(WORKBOOK_TARGET)
    await waitFor(() => expect(screen.getByTestId('dirty')).toHaveTextContent('false'))

    act(() => getEditor().setCellValue(0, 0, 0, 'changed'))
    await waitFor(() => expect(screen.getByTestId('dirty')).toHaveTextContent('true'))

    await act(async () => {
      await getEditor().handleSave()
    })

    await waitFor(() => expect(screen.getByTestId('dirty')).toHaveTextContent('false'))
  })
})

describe('useSpreadsheetEditor — save (workbook target)', () => {
  it('calls saveBinaryFile with workbook bytes, the right suggested name, and the existing path', async () => {
    const { getEditor } = renderHarness(WORKBOOK_TARGET)
    await waitFor(() => expect(getEditor()).toBeTruthy())

    await act(async () => {
      await getEditor().handleSave()
    })

    expect(window.electronAPI!.saveBinaryFile).toHaveBeenCalledWith(
      expect.objectContaining({
        suggestedName: 'sample.xlsx',
        filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }],
      }),
    )
    const call = (window.electronAPI!.saveBinaryFile as ReturnType<typeof vi.fn>).mock.calls[0][0]
    expect(call.content).toBeInstanceOf(Uint8Array)
    expect(call.existingPath).toBeUndefined()
  })

  it('handleSaveAs omits existingPath even after a prior save (forces the dialog)', async () => {
    const { getEditor } = renderHarness(WORKBOOK_TARGET)
    await waitFor(() => expect(getEditor()).toBeTruthy())

    await act(async () => {
      await getEditor().handleSave()
    })
    await act(async () => {
      await getEditor().handleSaveAs()
    })

    const calls = (window.electronAPI!.saveBinaryFile as ReturnType<typeof vi.fn>).mock.calls
    expect(calls[1][0].existingPath).toBeUndefined()
  })

  it('uses the existing saved path (no dialog) on a plain save after a prior save', async () => {
    const { getEditor } = renderHarness(WORKBOOK_TARGET)
    await waitFor(() => expect(getEditor()).toBeTruthy())

    await act(async () => {
      await getEditor().handleSave()
    })
    await act(async () => {
      await getEditor().handleSave()
    })

    const calls = (window.electronAPI!.saveBinaryFile as ReturnType<typeof vi.fn>).mock.calls
    expect(calls[1][0].existingPath).toBe('/tmp/sample.xlsx')
  })

  it('sets saveError and keeps the document dirty when the save fails', async () => {
    window.electronAPI!.saveBinaryFile = vi.fn().mockResolvedValue({ saved: false, error: 'Disk full' })
    const { getEditor } = renderHarness(WORKBOOK_TARGET)
    await waitFor(() => expect(screen.getByTestId('dirty')).toHaveTextContent('false'))

    act(() => getEditor().setCellValue(0, 0, 0, 'changed'))
    await act(async () => {
      const ok = await getEditor().handleSave()
      expect(ok).toBe(false)
    })

    expect(await screen.findByTestId('save-error')).toHaveTextContent('Disk full')
    expect(screen.getByTestId('dirty')).toHaveTextContent('true')
  })
})

describe('useSpreadsheetEditor — save (delimited target)', () => {
  it('calls saveFile with delimited text content', async () => {
    const { getEditor } = renderHarness(DELIMITED_TARGET)
    await waitFor(() => expect(getEditor()).toBeTruthy())

    await act(async () => {
      await getEditor().handleSave()
    })

    expect(window.electronAPI!.saveFile).toHaveBeenCalledWith(
      expect.objectContaining({ content: 'a,b\r\n1,2', suggestedName: 'sample.csv' }),
    )
  })
})

describe('useSpreadsheetEditor — undo/redo', () => {
  it('undo reverts the last edit', async () => {
    const { getEditor } = renderHarness(WORKBOOK_TARGET)
    await waitFor(() => expect(getEditor().document.sheets[0].rows[0][0]).toBe('a'))

    act(() => getEditor().setCellValue(0, 0, 0, 'changed'))
    await waitFor(() => expect(getEditor().document.sheets[0].rows[0][0]).toBe('changed'))

    act(() => getEditor().undo())
    await waitFor(() => expect(getEditor().document.sheets[0].rows[0][0]).toBe('a'))
  })
})

describe('useSpreadsheetEditor — the global Ctrl+S/Save contract still works', () => {
  it('the ViewerContext save() button triggers the same save path', async () => {
    renderHarness(WORKBOOK_TARGET)
    await waitFor(() => expect(screen.getByTestId('dirty')).toHaveTextContent('false'))

    await act(async () => {
      screen.getByText('Context Save').click()
    })

    expect(window.electronAPI!.saveBinaryFile).toHaveBeenCalled()
  })
})
