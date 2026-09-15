/**
 * P1.12 (DAT-05 / DAT-06) regression test — updated for wave 3.
 *
 * Originally: SpreadsheetViewer and CsvViewer used `allowOverlay: true` with
 * no `onCellEdited` wired up, so double-clicking a cell opened a phantom
 * edit box whose typed input silently vanished (DAT-06); neither grid passed
 * `getCellsForSelection`, so Ctrl+C never actually copied a selection
 * (DAT-05). P1.12's interim fix locked in `allowOverlay: false` (the
 * "disable it and be honest" option the findings register also offered).
 *
 * Wave 3 implements the findings register's OTHER listed resolution for
 * DAT-06 instead — "wire real editing+save" — so the correct assertion is
 * now the opposite of what this file originally checked: cells report
 * `allowOverlay: true` again, but this time a committed edit via
 * `onCellEdited` genuinely reaches the document model and is reflected back
 * out through `getCellContent` on the next render, instead of vanishing.
 * `getCellsForSelection` stays `true` either way (DAT-05 is unrelated to
 * this change).
 */
import { act, render, waitFor } from '@testing-library/react'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { EditableGridCell, Item } from '@glideapps/glide-data-grid'

import type { LoadedFile } from '../../formats/types'
import { ViewerProvider } from '../shared/ViewerContext'

type CapturedCell = { readonly allowOverlay: boolean; readonly displayData: string }
type CapturedProps = {
  readonly getCellsForSelection?: unknown
  readonly getCellContent: (loc: readonly [number, number]) => CapturedCell
  readonly onCellEdited?: (cell: Item, newValue: EditableGridCell) => void
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

// Only `read` is mocked (a fixed two-row worksheet) — `utils`/`write` stay
// real so SpreadsheetViewer's direct `!ref`/`encode_cell` walk (T1/T4) and
// its save path (wave 3) behave exactly as they would against a genuine
// parsed workbook.
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

function commitEdit(col: number, row: number, text: string): void {
  const newValue = { kind: 'text', data: text, displayData: text, allowOverlay: true } as EditableGridCell
  lastDataEditorProps!.onCellEdited?.([col, row], newValue)
}

beforeEach(() => {
  lastDataEditorProps = null
})

describe('grid editability honesty (DAT-05 / DAT-06, closed for real in wave 3)', () => {
  it('SpreadsheetViewer: cells are genuinely editable and selection copy is enabled', async () => {
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
    expect(lastDataEditorProps?.getCellContent([0, 0]).allowOverlay).toBe(true)

    // The critical DAT-06 check: a committed edit must not silently vanish.
    act(() => commitEdit(0, 1, 'Bob'))
    await waitFor(() => expect(lastDataEditorProps?.getCellContent([0, 1]).displayData).toBe('Bob'))
  })

  it('CsvViewer: cells are genuinely editable and selection copy is enabled', async () => {
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
    expect(lastDataEditorProps?.getCellContent([0, 0]).allowOverlay).toBe(true)

    act(() => commitEdit(0, 1, 'Bob'))
    await waitFor(() => expect(lastDataEditorProps?.getCellContent([0, 1]).displayData).toBe('Bob'))
  })
})
