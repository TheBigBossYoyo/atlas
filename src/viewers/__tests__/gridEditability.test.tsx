/**
 * P1.12 (DAT-05 / DAT-06) regression test.
 *
 * SpreadsheetViewer and CsvViewer used `allowOverlay: true` with no
 * `onCellEdited` wired up, so double-clicking a cell opened a phantom edit
 * box whose typed input silently vanished (DAT-06); neither grid passed
 * `getCellsForSelection`, so Ctrl+C never actually copied a selection
 * (DAT-05). This test locks in the fix: cells report `allowOverlay: false`
 * and the grid is given `getCellsForSelection={true}`.
 */
import { render, waitFor } from '@testing-library/react'
import { beforeEach, describe, expect, it, vi } from 'vitest'

import type { LoadedFile } from '../../formats/types'
import { ViewerProvider } from '../shared/ViewerContext'

type CapturedCell = { readonly allowOverlay: boolean }
type CapturedProps = {
  readonly getCellsForSelection?: unknown
  readonly getCellContent: (loc: readonly [number, number]) => CapturedCell
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

// Only `read` is mocked (a fixed two-row worksheet) — `utils` stays real so
// SpreadsheetViewer's direct `!ref`/`encode_cell` walk (T1/T4) behaves
// exactly as it would against a genuine parsed workbook.
vi.mock('xlsx', async (importOriginal) => {
  const actual = await importOriginal<typeof import('xlsx')>()
  return {
    ...actual,
    read: () => ({
      SheetNames: ['Sheet1'],
      Sheets: {
        Sheet1: {
          '!ref': 'A1:B2',
          A1: { t: 's', v: 'Name', w: 'Name' },
          B1: { t: 's', v: 'Score', w: 'Score' },
          A2: { t: 's', v: 'Alice', w: 'Alice' },
          B2: { t: 's', v: '10', w: '10' },
        },
      },
    }),
  }
})

beforeEach(() => {
  lastDataEditorProps = null
})

describe('grid editability honesty (DAT-05 / DAT-06)', () => {
  it('SpreadsheetViewer: cells are non-editable and selection copy is enabled', async () => {
    const { SpreadsheetViewer } = await import('../SpreadsheetViewer')
    const file: LoadedFile = {
      kind: 'binary',
      content: new ArrayBuffer(0),
      path: '/tmp/example.xlsx',
      format: 'xlsx',
    }

    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(lastDataEditorProps).not.toBeNull()
    })

    expect(lastDataEditorProps?.getCellsForSelection).toBe(true)
    expect(lastDataEditorProps?.getCellContent([0, 0]).allowOverlay).toBe(false)
  })

  it('CsvViewer: cells are non-editable and selection copy is enabled', async () => {
    const { CsvViewer } = await import('../CsvViewer')
    const file: LoadedFile = {
      kind: 'text',
      content: 'Name,Score\nAlice,10\n',
      path: '/tmp/example.csv',
      format: 'csv',
    }

    render(
      <ViewerProvider filePath={file.path}>
        <CsvViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(lastDataEditorProps).not.toBeNull()
    })

    expect(lastDataEditorProps?.getCellsForSelection).toBe(true)
    expect(lastDataEditorProps?.getCellContent([0, 0]).allowOverlay).toBe(false)
  })
})
