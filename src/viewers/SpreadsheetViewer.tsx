import { memo, useEffect, useMemo, useState, useCallback, lazy, Suspense, useDeferredValue } from 'react'
import { Table, FileSpreadsheet } from 'lucide-react'

import type { ViewerProps } from '../formats/types'
import { useSetNavItems, useSetViewerStats } from './shared/useViewerContext'
import { useGridTheme } from './shared/useGridTheme'
import './__styles__/viewer-spreadsheet.css'

import type { GridCell, GridMouseEventArgs, GridColumn } from '@glideapps/glide-data-grid'

// Lazy load DataEditor and its CSS
const LazyDataEditor = lazy(async () => {
  const mod = await import('@glideapps/glide-data-grid')
  await import('@glideapps/glide-data-grid/dist/index.css')
  return { default: mod.DataEditor }
})

type SheetData = {
  name: string
  rows: string[][]
  colCount: number
}

function SpreadsheetViewerBase({ file }: ViewerProps) {
  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()
  const theme = useGridTheme()

  const [sheets, setSheets] = useState<SheetData[]>([])
  const [activeSheetName, setActiveSheetName] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)

  const [search, setSearch] = useState('')
  const deferredSearch = useDeferredValue(search)
  const [hoveredRow, setHoveredRow] = useState<number | undefined>()

  useEffect(() => {
    let cancelled = false

    if (file.kind !== 'binary') {
      setError('Expected binary file for spreadsheet viewer.')
      return
    }

    async function loadWorkbook() {
      const content = file.content
      try {
        const XLSX = await import('xlsx')
        const workbook = XLSX.read(content, { type: 'array', cellFormula: true, cellStyles: true, sheetStubs: true })
        if (cancelled) return

        const parsedSheets: SheetData[] = []
        for (const name of workbook.SheetNames) {
          const ws = workbook.Sheets[name]
          const json = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: '' })
          const maxCols = json.reduce((max, row) => Math.max(max, row.length), 0)
          parsedSheets.push({ name, rows: json, colCount: maxCols })
        }

        setSheets(parsedSheets)
        if (parsedSheets.length > 0) {
          setActiveSheetName(parsedSheets[0].name)
        }
      } catch (err) {
        if (!cancelled) {
          setError(err instanceof Error ? err.message : String(err))
        }
      }
    }

    loadWorkbook()

    return () => {
      cancelled = true
    }
  }, [file])

  const activeSheet = useMemo(
    () => sheets.find(s => s.name === activeSheetName) ?? sheets[0],
    [sheets, activeSheetName]
  )

  const filteredRows = useMemo(() => {
    if (!activeSheet) return []
    const q = deferredSearch.toLowerCase()
    if (!q) return activeSheet.rows
    return activeSheet.rows.filter(row => row.some(cell => String(cell).toLowerCase().includes(q)))
  }, [activeSheet, deferredSearch])

  const navItems = useMemo(() => {
    return sheets.map(sheet => ({
      id: sheet.name,
      label: sheet.name,
      icon: Table,
      onSelect: () => setActiveSheetName(sheet.name),
    }))
  }, [sheets])

  useEffect(() => {
    setNavItems(navItems)
  }, [setNavItems, navItems])

  useEffect(() => {
    if (activeSheet) {
      setStats({
        kind: 'spreadsheet',
        sheet: activeSheet.name,
        rows: filteredRows.length,
        cols: activeSheet.colCount,
      })
    } else {
      setStats(null)
    }
  }, [setStats, activeSheet, filteredRows.length])

  const getCellContent = useCallback(
    ([col, row]: readonly [number, number]): GridCell => {
      if (!activeSheet) {
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
    [activeSheet, filteredRows, theme, hoveredRow]
  )

  const [colWidths, setColWidths] = useState<Record<string, number>>({})

  // Clear col widths on sheet change
  useEffect(() => {
    setColWidths({})
  }, [activeSheetName])

  const columns = useMemo(() => {
    if (!activeSheet) return []
    return Array.from({ length: activeSheet.colCount }).map((_, i) => {
      let title = ''
      let n = i
      while (n >= 0) {
        title = String.fromCharCode(65 + (n % 26)) + title
        n = Math.floor(n / 26) - 1
      }
      
      // Sample first 50 non-empty rows
      let numericCount = 0
      let totalSampled = 0
      for (const row of activeSheet.rows) {
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
  }, [activeSheet, colWidths])

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
    return <div className="spreadsheet-viewer__error">{error}</div>
  }

  if (sheets.length === 0) {
    return <div className="spreadsheet-viewer" />
  }

  return (
    <div className="spreadsheet-viewer">
      <div className="spreadsheet-viewer__toolbar">
        <div className="spreadsheet-viewer__stats">
          {activeSheet ? `${filteredRows.length} rows × ${activeSheet.colCount} columns` : ''}
        </div>
        <div className="spreadsheet-viewer__search">
          <input
            type="search"
            placeholder="Search rows..."
            value={search}
            onChange={e => setSearch(e.target.value)}
          />
        </div>
      </div>
      <div className="spreadsheet-viewer__grid">
        {filteredRows.length === 0 ? (
          <div className="spreadsheet-viewer__empty">
            <FileSpreadsheet size={48} />
            <p>{search ? 'No rows match your search.' : 'This sheet is empty.'}</p>
          </div>
        ) : (
          <Suspense fallback={null}>
            {activeSheet && (
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
            )}
          </Suspense>
        )}
      </div>
      <div className="spreadsheet-viewer__tabs">
        {sheets.map(sheet => (
          <button
            key={sheet.name}
            className={`spreadsheet-viewer__tab ${sheet.name === activeSheetName ? 'spreadsheet-viewer__tab--active' : ''}`}
            onClick={() => {
              setActiveSheetName(sheet.name)
              setSearch('')
            }}
          >
            <Table size={14} />
            {sheet.name}
          </button>
        ))}
      </div>
    </div>
  )
}

export const SpreadsheetViewer = memo(SpreadsheetViewerBase)
