import { memo, useEffect, useMemo, useState, useCallback, lazy, Suspense, useDeferredValue } from 'react'
import { FileSpreadsheet } from 'lucide-react'

import type { ViewerProps } from '../formats/types'
import { useSetNavItems, useSetViewerStats } from './shared/useViewerContext'
import { useGridTheme } from './shared/useGridTheme'
import './__styles__/viewer-spreadsheet.css'

import type { GridCell, GridMouseEventArgs, GridColumn } from '@glideapps/glide-data-grid'
import type { ParseResult } from 'papaparse'

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

function CsvViewerBase({ file }: ViewerProps) {
  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()
  const theme = useGridTheme()

  const [data, setData] = useState<CsvData | null>(null)
  const [error, setError] = useState<string | null>(null)
  
  const [search, setSearch] = useState('')
  const deferredSearch = useDeferredValue(search)
  const [hoveredRow, setHoveredRow] = useState<number | undefined>()

  useEffect(() => {
    let cancelled = false

    if (file.kind !== 'text') {
      setError('Expected text file for CSV/TSV viewer.')
      return
    }

    async function loadCsv() {
      const content = file.content as string
      try {
        const Papa = (await import('papaparse')).default
        if (cancelled) return

        const delimiter = file.format === 'tsv' ? '\t' : undefined

        const result = Papa.parse(content, {
          worker: false,
          header: false,
          skipEmptyLines: true,
          delimiter,
        }) as unknown as ParseResult<string[]>

        if (cancelled) return

        if (result.errors.length > 0 && result.data.length === 0) {
          setError(result.errors[0].message)
          return
        }

        const rows = result.data
        const colCount = rows.reduce((max: number, row: string[]) => Math.max(max, row.length), 0)

        setData({ rows, colCount })
      } catch (err) {
        if (!cancelled) {
          setError(err instanceof Error ? err.message : String(err))
        }
      }
    }

    loadCsv()

    return () => {
      cancelled = true
    }
  }, [file])

  useEffect(() => {
    setNavItems([])
  }, [setNavItems])

  const filteredRows = useMemo(() => {
    if (!data) return []
    const q = deferredSearch.toLowerCase()
    if (!q) return data.rows
    return data.rows.filter(row => row.some(cell => String(cell).toLowerCase().includes(q)))
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

  const getCellContent = useCallback(
    ([col, row]: readonly [number, number]): GridCell => {
      if (!data) {
        return { kind: 'text' as const, data: '', displayData: '', allowOverlay: false } as GridCell
      }
      const cellValue = filteredRows[row]?.[col] ?? ''
      
      const isOdd = row % 2 !== 0
      const isHovered = row === hoveredRow
      let bgCell = isOdd ? theme.bgRowOdd : theme.bgCell
      if (isHovered) bgCell = theme.bgRowHover
      
      return {
        kind: 'text' as const,
        data: cellValue,
        displayData: String(cellValue),
        allowOverlay: true,
        themeOverride: { bgCell },
      } as GridCell
    },
    [data, filteredRows, theme, hoveredRow]
  )

  const [colWidths, setColWidths] = useState<Record<string, number>>({})

  const columns = useMemo(() => {
    if (!data) return []
    return Array.from({ length: data.colCount }).map((_, i) => {
      let title = ''
      let n = i
      while (n >= 0) {
        title = String.fromCharCode(65 + (n % 26)) + title
        n = Math.floor(n / 26) - 1
      }
      
      // Sample first 50 non-empty rows
      let numericCount = 0
      let totalSampled = 0
      for (const row of data.rows) {
        const val = row[i]
        if (val !== undefined && val !== null && String(val).trim() !== '') {
          totalSampled++
          if (!isNaN(Number(val))) {
            numericCount++
          }
        }
        if (totalSampled >= 50) break
      }
      
      const isNumeric = totalSampled > 0 && (numericCount / totalSampled) >= 0.8
      
      return { 
        title, 
        id: String(i), 
        width: colWidths[i] ?? 120, 
        grow: colWidths[i] ? undefined : 1,
        contentAlign: isNumeric ? 'right' as const : undefined
      }
    })
  }, [data, colWidths])

  const onColumnResize = useCallback((column: GridColumn, newSize: number) => {
    setColWidths(prev => ({ ...prev, [column.id ?? '']: newSize }))
  }, [])

  const onItemHovered = useCallback((args: GridMouseEventArgs) => {
    if (args.location[1] >= 0) {
      setHoveredRow(args.location[1])
    } else {
      setHoveredRow(undefined)
    }
  }, [])

  if (error) {
    return <div className="csv-viewer__error">{error}</div>
  }

  if (!data) {
    return <div className="csv-viewer" />
  }

  return (
    <div className="csv-viewer">
      <div className="csv-viewer__toolbar">
        <div className="csv-viewer__stats">
          {filteredRows.length} rows × {data.colCount} columns
        </div>
        <div className="csv-viewer__search">
          <input
            type="search"
            placeholder="Search rows..."
            value={search}
            onChange={e => setSearch(e.target.value)}
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
              columns={columns}
              rows={filteredRows.length}
              theme={theme}
              width="100%"
              height="100%"
              smoothScrollX
              smoothScrollY
              rowMarkers="number"
              freezeColumns={1}
              headerHeight={36}
              rowHeight={32}
              onItemHovered={onItemHovered}
              onColumnResize={onColumnResize}
            />
          </Suspense>
        )}
      </div>
    </div>
  )
}

export const CsvViewer = memo(CsvViewerBase)
