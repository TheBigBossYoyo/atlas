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
import { translate } from '../../../i18n/translate'
import { forgetTextFileMeta, setTextFileMeta } from '../../../utils/textDecoding'

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

type HarnessProps = {
  readonly target: SpreadsheetSaveTarget
  readonly onReady: (api: ReturnType<typeof useSpreadsheetEditor>) => void
  /** SHEET-4 — the bytes the workbook was "opened" from, threaded straight through to `writeWorkbookThroughOriginal` (USR-17's save-through-original). */
  readonly originalBuffer?: ArrayBuffer | null
}

function Harness({ target, onReady, originalBuffer = null }: HarnessProps) {
  const doc = createDocument([sheetFixture([['a', 'b'], ['1', '2']])])
  const editor = useSpreadsheetEditor(doc, '/tmp/sample.xlsx', target, true, originalBuffer)
  onReady(editor)
  return (
    <>
      <ViewerDirtyProbe />
      <ViewerSaveProbe />
      {editor.saveError && <span data-testid="save-error">{editor.saveError}</span>}
      {editor.saveWarning && <span data-testid="save-warning">{editor.saveWarning}</span>}
    </>
  )
}

function renderHarness(target: SpreadsheetSaveTarget, originalBuffer?: ArrayBuffer | null) {
  let latest: ReturnType<typeof useSpreadsheetEditor> | null = null
  const utils = render(
    <ViewerProvider filePath="/tmp/sample.xlsx">
      <Harness target={target} originalBuffer={originalBuffer} onReady={(api) => { latest = api }} />
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
    // Seeded from the loaded file's own path (mirrors DocxViewer) — the very
    // first Save overwrites in place rather than acting like a Save As.
    expect(call.existingPath).toBe('/tmp/sample.xlsx')
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

// SHEET-4 — an xlsx save silently fell back from the lossless
// save-through-original writer to the lossy fresh-SheetJS one whenever
// `writeWorkbookThroughOriginal` returned `null` (a non-OOXML/corrupt
// buffer, missing relationships, ...), with no way for the user to tell:
// `useSpreadsheetEditor.ts`'s `throughOriginal ?? writeWorkbookBytesWithTables(...)`
// reported plain success either way. `saveWarning` now surfaces that.
describe('useSpreadsheetEditor — save-through-original fallback warning (SHEET-4)', () => {
  it('a forced passthrough failure (a buffer that is not a real xlsx zip) surfaces a save warning, but the save still succeeds', async () => {
    const notAZip = new TextEncoder().encode('this is not a zip file').buffer as ArrayBuffer
    const { getEditor } = renderHarness(WORKBOOK_TARGET, notAZip)
    await waitFor(() => expect(getEditor()).toBeTruthy())

    let ok = false
    await act(async () => {
      ok = await getEditor().handleSave()
    })

    // The save itself still succeeded (the fresh-SheetJS writer produced
    // valid bytes and `saveBinaryFile` reported success) — a fallback is not
    // a failure, so this must stay `true` and `saveError` must stay unset.
    expect(ok).toBe(true)
    expect(getEditor().saveError).toBeNull()

    const expectedWarning = translate('en', 'spreadsheet.saveFallbackWarning')
    expect(await screen.findByTestId('save-warning')).toHaveTextContent(expectedWarning)
    expect(getEditor().saveWarning).toBe(expectedWarning)
  })

  it('a normal passthrough save (no original buffer at all) stays completely silent — no save warning', async () => {
    const { getEditor } = renderHarness(WORKBOOK_TARGET)
    await waitFor(() => expect(getEditor()).toBeTruthy())

    await act(async () => {
      await getEditor().handleSave()
    })

    expect(getEditor().saveWarning).toBeNull()
    expect(screen.queryByTestId('save-warning')).toBeNull()
  })

  it('a save NOT eligible for passthrough (delimited target) never reports the fallback warning', async () => {
    // A CSV/TSV save never goes through `writeWorkbookThroughOriginal` at
    // all — `usedFallbackWriter` must stay `false` here, not conflate "this
    // format was never a passthrough candidate" with "passthrough was
    // attempted and failed".
    const notAZip = new TextEncoder().encode('this is not a zip file').buffer as ArrayBuffer
    const { getEditor } = renderHarness(DELIMITED_TARGET, notAZip)
    await waitFor(() => expect(getEditor()).toBeTruthy())

    await act(async () => {
      await getEditor().handleSave()
    })

    expect(getEditor().saveWarning).toBeNull()
  })

  it('the warning clears on the next save once passthrough is no longer failing', async () => {
    const notAZip = new TextEncoder().encode('this is not a zip file').buffer as ArrayBuffer
    const { getEditor, rerender } = renderHarness(WORKBOOK_TARGET, notAZip)
    await waitFor(() => expect(getEditor()).toBeTruthy())

    await act(async () => {
      await getEditor().handleSave()
    })
    expect(getEditor().saveWarning).not.toBeNull()

    // Re-render with no `originalBuffer` at all (so this next save is no
    // longer even eligible for passthrough) — a stale warning from the
    // PREVIOUS save must not linger forever.
    let latest: ReturnType<typeof useSpreadsheetEditor> | null = null
    rerender(
      <ViewerProvider filePath="/tmp/sample.xlsx">
        <Harness target={WORKBOOK_TARGET} originalBuffer={null} onReady={(api) => { latest = api }} />
      </ViewerProvider>,
    )
    await act(async () => {
      await latest!.handleSave()
    })

    expect(latest!.saveWarning).toBeNull()
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
    // A file never loaded from disk sends no encoding metadata (historical output unchanged).
    expect(vi.mocked(window.electronAPI!.saveFile).mock.calls[0][0]).not.toHaveProperty('meta')
  })

  // SHEET-7/8 — a CSV read from disk hands its recorded encoding/BOM/newline
  // to save-file (which reapplies them) and sends `\n`-only text: sending
  // CRLF here would come out as `\r\r\n` once save-file converts `\n`.
  it('sends a loaded file\'s recorded encoding/BOM/newline with LF-only text, and keeps a final line break', async () => {
    const meta = { encoding: 'utf-8', bom: true, newline: 'crlf' } as const
    setTextFileMeta('/tmp/sample.xlsx', meta)
    try {
      const { getEditor } = renderHarness({ ...DELIMITED_TARGET, trailingNewline: true })
      await waitFor(() => expect(getEditor()).toBeTruthy())

      await act(async () => {
        await getEditor().handleSave()
      })

      expect(window.electronAPI!.saveFile).toHaveBeenCalledWith(
        expect.objectContaining({ content: 'a,b\n1,2\n', meta }),
      )
    } finally {
      forgetTextFileMeta('/tmp/sample.xlsx')
    }
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

  it('a blocked/no-op action (deleting the last remaining row) does not push a spurious undo entry', async () => {
    // The fixture is a 2-row sheet, so first collapse it to exactly one row
    // — deleteRowAt on a one-row sheet is a documented no-op
    // (spreadsheetDocument.ts) that the toolbar can still reach (it only
    // disables "Delete row" when nothing is selected, not when deleting
    // would be a no-op).
    const { getEditor } = renderHarness(WORKBOOK_TARGET)
    await waitFor(() => expect(getEditor().document.sheets[0].rows).toHaveLength(2))

    act(() => getEditor().deleteRowAt(0, 1))
    await waitFor(() => expect(getEditor().document.sheets[0].rows).toHaveLength(1))
    expect(getEditor().canUndo).toBe(true)

    act(() => getEditor().undo())
    await waitFor(() => expect(getEditor().document.sheets[0].rows).toHaveLength(2))
    expect(getEditor().canUndo).toBe(false)

    // The real edit is now fully undone — canUndo must be false, not "one
    // more no-op entry away from false". Confirms the guarded no-op below
    // never reached the history stack.
    act(() => getEditor().deleteRowAt(0, 0))
    await waitFor(() => expect(getEditor().document.sheets[0].rows).toHaveLength(1))
    const docAfterRealDelete = getEditor().document

    // Now try to delete the last remaining row — a genuine no-op.
    act(() => getEditor().deleteRowAt(0, 0))
    expect(getEditor().document).toBe(docAfterRealDelete)
    expect(getEditor().document.sheets[0].rows).toHaveLength(1)

    // A single Undo must revert the one REAL delete above, not silently
    // consume a spurious entry from the blocked one first.
    act(() => getEditor().undo())
    await waitFor(() => expect(getEditor().document.sheets[0].rows).toHaveLength(2))
    expect(getEditor().canUndo).toBe(false)
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
