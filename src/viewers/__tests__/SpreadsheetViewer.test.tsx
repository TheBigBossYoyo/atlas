/**
 * T11 (DAT-16) — SpreadsheetViewer render/parse fixture coverage, plus
 * T1/T4/T11 regression coverage for formatted values, merged cells, and the
 * hidden-sheet toggle.
 */
import { readFileSync } from 'node:fs'
import path from 'node:path'
import { render, renderHook, waitFor, fireEvent, screen } from '@testing-library/react'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import * as XLSX from 'xlsx'

import type { LoadedFile } from '../../formats/types'
import { ViewerProvider } from '../shared/ViewerContext'
import { useOpenViewerFind } from '../shared/useViewerContext'

type CapturedCell = { readonly data: string; readonly displayData: string }
type CapturedProps = {
  readonly getCellContent: (loc: readonly [number, number]) => CapturedCell
  readonly columns: ReadonlyArray<{ readonly title: string }>
  readonly rows: number
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

function buildWorkbookFile(build: (wb: XLSX.WorkBook) => void): LoadedFile {
  const wb = XLSX.utils.book_new()
  build(wb)
  const content = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer
  return { kind: 'binary', content, path: '/tmp/generated.xlsx', format: 'xlsx' }
}

beforeEach(() => {
  lastDataEditorProps = null
})

describe('SpreadsheetViewer — fixture render/parse (T11)', () => {
  it('renders a real .xlsx fixture with formatted header + data rows', async () => {
    const { SpreadsheetViewer } = await import('../SpreadsheetViewer')
    const file: LoadedFile = {
      kind: 'binary',
      content: readFixtureArrayBuffer('sample.xlsx'),
      path: '/tmp/sample.xlsx',
      format: 'xlsx',
    }

    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())
    expect(gridRows(lastDataEditorProps!)).toEqual([
      ['Name', 'Value'],
      ['Atlas', '2'],
      ['Phase', '1'],
    ])
  })

  it('renders a real .ods fixture through the same viewer', async () => {
    const { OdsViewer } = await import('../OdsViewer')
    const file: LoadedFile = {
      kind: 'binary',
      content: readFixtureArrayBuffer('sample.ods'),
      path: '/tmp/sample.ods',
      format: 'ods',
    }

    render(
      <ViewerProvider filePath={file.path}>
        <OdsViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())
    expect(lastDataEditorProps!.rows).toBeGreaterThan(0)
  })
})

describe('SpreadsheetViewer — merges, hidden sheets, error/date formatting (T1/T4)', () => {
  it('backfills a merged range instead of leaving it blank outside the top-left cell (DAT-09)', async () => {
    const { SpreadsheetViewer } = await import('../SpreadsheetViewer')
    const file = buildWorkbookFile((wb) => {
      const ws = XLSX.utils.aoa_to_sheet([
        ['Merged Header', '', ''],
        ['a', 'b', 'c'],
      ])
      ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }]
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')
    })

    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())
    const rows = gridRows(lastDataEditorProps!)
    expect(rows[0]).toEqual(['Merged Header', 'Merged Header', 'Merged Header'])
  })

  it('renders a formula-error cell as its error text instead of blank (DAT-04)', async () => {
    const { SpreadsheetViewer } = await import('../SpreadsheetViewer')
    const file = buildWorkbookFile((wb) => {
      const ws = XLSX.utils.aoa_to_sheet([['ok']])
      ws['B1'] = { t: 'e', v: 0x07, w: '#DIV/0!' } as XLSX.CellObject
      ws['!ref'] = 'A1:B1'
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')
    })

    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())
    expect(gridRows(lastDataEditorProps!)[0]).toEqual(['ok', '#DIV/0!'])
  })

  it('hides a hidden sheet by default and reveals it via the "Show hidden sheets" toggle (DAT-11)', async () => {
    const { SpreadsheetViewer } = await import('../SpreadsheetViewer')
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['visible']]), 'Visible')
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['hidden']]), 'Hidden')
      wb.Workbook = { Sheets: [{ Hidden: 0 }, { Hidden: 1 }] }
    })

    render(
      <ViewerProvider filePath={file.path}>
        <SpreadsheetViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())
    expect(screen.queryByText('Hidden')).not.toBeInTheDocument()

    const toggle = screen.getByRole('button', { name: /show hidden sheets/i })
    fireEvent.click(toggle)

    expect(screen.getByText('Hidden')).toBeInTheDocument()
  })
})

describe('SpreadsheetViewer — in-viewer find (openFind capability)', () => {
  it('registers a find handler that opens the find overlay when invoked', async () => {
    const { SpreadsheetViewer } = await import('../SpreadsheetViewer')
    const file = buildWorkbookFile((wb) => {
      XLSX.utils.book_append_sheet(
        wb,
        XLSX.utils.aoa_to_sheet([
          ['Name', 'City'],
          ['Alice', 'Boston'],
        ]),
        'Sheet1',
      )
    })

    const { result } = renderHook(() => useOpenViewerFind(), {
      wrapper: ({ children }) => (
        <ViewerProvider filePath={file.path}>
          <SpreadsheetViewer file={file} />
          {children}
        </ViewerProvider>
      ),
    })

    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())

    expect(screen.queryByPlaceholderText('Search in document...')).not.toBeInTheDocument()

    result.current()

    expect(await screen.findByPlaceholderText('Search in document...')).toBeInTheDocument()
  })
})
