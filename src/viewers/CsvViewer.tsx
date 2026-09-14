import { memo, useEffect, useMemo, useState, useDeferredValue, lazy, Suspense } from 'react'
import { FileSpreadsheet } from 'lucide-react'

import type { ViewerProps } from '../formats/types'
import { useSetNavItems, useSetViewerStats, useRegisterViewerFind } from './shared/useViewerContext'
import { useSpreadsheetGrid } from './shared/useSpreadsheetGrid'
import { useGridFind } from './shared/useGridFind'
import { parseCsv } from './shared/csvParse'
import { ViewerLoading } from '../components/ViewerLoading'
import { SearchOverlay } from '../components/SearchOverlay'
import './__styles__/viewer-spreadsheet.css'

// Lazy load DataEditor and its CSS
const LazyDataEditor = lazy(async () => {
  const mod = await import('@glideapps/glide-data-grid')
  await import('@glideapps/glide-data-grid/dist/index.css')
  return { default: mod.DataEditor }
})

type CsvData = {
  rows: string[][]
  colCount: number
}

type LoadState =
  | { status: 'loading' }
  | { status: 'ready'; data: CsvData }
  | { status: 'error'; error: string }

function CsvViewerBase({ file }: ViewerProps) {
  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()
  const registerFind = useRegisterViewerFind()

  const [state, setState] = useState<LoadState>({ status: 'loading' })
  const [search, setSearch] = useState('')
  const deferredSearch = useDeferredValue(search)

  useEffect(() => {
    let cancelled = false

    // Runs synchronously up to the first `await`, so the "loading"/"wrong
    // kind" reset lands in the same tick as the effect itself — identical
    // timing to setting state directly in the effect body, but nested in a
    // callback so it synchronizes with the external `parseCsv` call rather
    // than reading as derivable state (mirrors DocxViewer's load effect).
    void (async () => {
      if (file.kind !== 'text') {
        setState({ status: 'error', error: 'Expected text file for CSV/TSV viewer.' })
        return
      }

      setState({ status: 'loading' })
      const delimiter = file.format === 'tsv' ? '\t' : undefined

      try {
        const result = await parseCsv(file.content, { delimiter })
        if (cancelled) return
        if (result.errors.length > 0 && result.data.length === 0) {
          setState({ status: 'error', error: result.errors[0].message })
          return
        }
        const rows = result.data
        const colCount = rows.reduce((max: number, row: string[]) => Math.max(max, row.length), 0)
        setState({ status: 'ready', data: { rows, colCount } })
      } catch (err) {
        if (!cancelled) {
          setState({ status: 'error', error: err instanceof Error ? err.message : String(err) })
        }
      }
    })()

    return () => {
      cancelled = true
    }
  }, [file])

  useEffect(() => {
    setNavItems([])
  }, [setNavItems])

  const data = state.status === 'ready' ? state.data : null

  const filteredRows = useMemo(() => {
    if (!data) return []
    const q = deferredSearch.toLowerCase()
    if (!q) return data.rows
    return data.rows.filter((row) => row.some((cell) => String(cell).toLowerCase().includes(q)))
  }, [data, deferredSearch])

  useEffect(() => {
    if (data) {
      setStats({
        kind: 'spreadsheet',
        sheet: 'data',
        rows: filteredRows.length,
        cols: data.colCount,
      })
    } else {
      setStats(null)
    }
  }, [setStats, data, filteredRows.length])

  const { columns, getCellContent, onColumnResize, onItemHovered, theme } = useSpreadsheetGrid({
    rows: filteredRows,
    colCount: data?.colCount ?? 0,
  })

  const gridFind = useGridFind(filteredRows)
  useEffect(() => {
    registerFind(gridFind.open)
    return () => registerFind(null)
  }, [registerFind, gridFind.open])

  if (state.status === 'error') {
    return <div className="csv-viewer__error">{state.error}</div>
  }

  if (state.status === 'loading') {
    return <ViewerLoading format={file.format} />
  }

  return (
    <div className="csv-viewer">
      <SearchOverlay
        isOpen={gridFind.isOpen}
        query={gridFind.query}
        matchCount={gridFind.matchCount}
        currentMatch={gridFind.currentMatch}
        onQueryChange={gridFind.setQuery}
        onNext={() => gridFind.goToMatch('next')}
        onPrev={() => gridFind.goToMatch('prev')}
        onClose={gridFind.close}
      />
      <div className="csv-viewer__toolbar">
        <div className="csv-viewer__stats">
          {filteredRows.length} rows × {data?.colCount ?? 0} columns
        </div>
        <div className="csv-viewer__search">
          <input
            type="search"
            placeholder="Search rows..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
        </div>
      </div>
      <div className="csv-viewer__grid">
        {filteredRows.length === 0 ? (
          <div className="csv-viewer__empty">
            <FileSpreadsheet size={48} />
            <p>{search ? 'No rows match your search.' : 'This file is empty.'}</p>
          </div>
        ) : (
          <Suspense fallback={null}>
            <LazyDataEditor
              getCellContent={getCellContent}
              getCellsForSelection={true}
              columns={columns}
              rows={filteredRows.length}
              theme={theme}
              width="100%"
              height="100%"
              smoothScrollX
              smoothScrollY
              rowMarkers="number"
              // Fixed UX default — CSV/TSV have no file-level freeze-pane
              // concept at all (that's an xlsx/ods-only feature); see the
              // "Frozen panes" note in shared/spreadsheetGrid.ts.
              freezeColumns={1}
              headerHeight={36}
              rowHeight={32}
              onItemHovered={onItemHovered}
              onColumnResize={onColumnResize}
              gridSelection={gridFind.gridSelection}
              onGridSelectionChange={gridFind.onGridSelectionChange}
            />
          </Suspense>
        )}
      </div>
    </div>
  )
}

export const CsvViewer = memo(CsvViewerBase)
