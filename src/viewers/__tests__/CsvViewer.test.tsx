/**
 * T11 (DAT-16) — CsvViewer render/parse fixture coverage, plus T2's loading
 * state and the in-viewer find capability.
 */
import { readFileSync } from 'node:fs'
import path from 'node:path'
import { render, renderHook, waitFor, screen } from '@testing-library/react'
import { beforeEach, describe, expect, it, vi } from 'vitest'

import type { LoadedFile } from '../../formats/types'
import { ViewerProvider } from '../shared/ViewerContext'
import { useOpenViewerFind } from '../shared/useViewerContext'
import { CsvViewer } from '../CsvViewer'
import { TsvViewer } from '../TsvViewer'

type CapturedCell = { readonly displayData: string }
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

function readFixtureText(name: string): string {
  return readFileSync(path.join(FIXTURES_DIR, name), 'utf8')
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

beforeEach(() => {
  lastDataEditorProps = null
})

describe('CsvViewer — fixture render/parse (T11)', () => {
  it('renders a real .csv fixture', async () => {
    const file: LoadedFile = {
      kind: 'text',
      content: readFixtureText('sample.csv'),
      path: '/tmp/sample.csv',
      format: 'csv',
    }

    render(
      <ViewerProvider filePath={file.path}>
        <CsvViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())
    expect(gridRows(lastDataEditorProps!)).toEqual([
      ['name', 'value'],
      ['Atlas', '2'],
      ['Phase', '1'],
    ])
  })

  it('renders a real .tsv fixture through TsvViewer', async () => {
    const file: LoadedFile = {
      kind: 'text',
      content: readFixtureText('sample.tsv'),
      path: '/tmp/sample.tsv',
      format: 'tsv',
    }

    render(
      <ViewerProvider filePath={file.path}>
        <TsvViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())
    expect(gridRows(lastDataEditorProps!)).toEqual([
      ['name', 'value'],
      ['Atlas', '2'],
      ['Phase', '1'],
    ])
  })

  it('shows a loading indicator while parsing, then the grid once ready', async () => {
    const file: LoadedFile = {
      kind: 'text',
      content: 'a,b\n1,2\n',
      path: '/tmp/loading.csv',
      format: 'csv',
    }

    render(
      <ViewerProvider filePath={file.path}>
        <CsvViewer file={file} />
      </ViewerProvider>,
    )

    expect(screen.getByRole('status')).toHaveTextContent(/loading/i)
    await waitFor(() => expect(lastDataEditorProps).not.toBeNull())
  })

  it('reports a friendly error for a non-text file', async () => {
    const file: LoadedFile = {
      kind: 'binary',
      content: new ArrayBuffer(0),
      path: '/tmp/wrong-kind.csv',
      format: 'csv',
    }

    render(
      <ViewerProvider filePath={file.path}>
        <CsvViewer file={file} />
      </ViewerProvider>,
    )

    expect(await screen.findByText(/expected text file/i)).toBeInTheDocument()
  })
})

describe('CsvViewer — in-viewer find (openFind capability)', () => {
  it('registers a find handler that opens the find overlay when invoked', async () => {
    const file: LoadedFile = {
      kind: 'text',
      content: 'Name,City\nAlice,Boston\n',
      path: '/tmp/find.csv',
      format: 'csv',
    }

    const { result } = renderHook(() => useOpenViewerFind(), {
      wrapper: ({ children }) => (
        <ViewerProvider filePath={file.path}>
          <CsvViewer file={file} />
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
