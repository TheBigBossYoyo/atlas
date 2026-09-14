import { memo, useEffect, useMemo, useState, lazy, Suspense, useDeferredValue } from 'react'
import { Table, FileSpreadsheet, Eye, EyeOff } from 'lucide-react'

import type { ViewerProps } from '../formats/types'
import { useSetNavItems, useSetViewerStats, useRegisterViewerFind } from './shared/useViewerContext'
import { useSpreadsheetWorkbook } from './shared/useSpreadsheetWorkbook'
import { useSpreadsheetGrid } from './shared/useSpreadsheetGrid'
import { useGridFind } from './shared/useGridFind'
import { ViewerLoading } from '../components/ViewerLoading'
import { SearchOverlay } from '../components/SearchOverlay'
import type { ParsedSheet } from './shared/spreadsheetGrid'
import './__styles__/viewer-spreadsheet.css'

// Lazy load DataEditor and its CSS
const LazyDataEditor = lazy(async () => {
  const mod = await import('@glideapps/glide-data-grid')
  await import('@glideapps/glide-data-grid/dist/index.css')
  return { default: mod.DataEditor }
})

function SpreadsheetViewerBase({ file }: ViewerProps) {
  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()
  const registerFind = useRegisterViewerFind()

  const [activeSheetName, setActiveSheetName] = useState<string | null>(null)
  const [showHiddenSheets, setShowHiddenSheets] = useState(false)
  const [search, setSearch] = useState('')
  const deferredSearch = useDeferredValue(search)

  const buffer = file.kind === 'binary' ? file.content : null
  const workbookState = useSpreadsheetWorkbook(buffer)
  const sheets: ParsedSheet[] = useMemo(
    () => (workbookState.status === 'ready' ? workbookState.sheets : []),
    [workbookState],
  )

  const hasHiddenSheets = useMemo(() => sheets.some((s) => s.hidden), [sheets])
  const visibleSheets = useMemo(
    () => (showHiddenSheets ? sheets : sheets.filter((s) => !s.hidden)),
    [sheets, showHiddenSheets],
  )

  // Pick an initial/fallback active sheet once the workbook is ready, or
  // when the current selection is no longer in the visible list (e.g. the
  // "show hidden sheets" toggle was switched off while a hidden one was
  // active). A pure derivation of `visibleSheets`/`activeSheetName`, so this
  // uses the render-time "adjust state" idiom (see ViewerContext.tsx) rather
  // than an effect.
  if (visibleSheets.length > 0 && !visibleSheets.some((s) => s.name === activeSheetName)) {
    setActiveSheetName(visibleSheets[0].name)
  }

  const activeSheet = useMemo(
    () => visibleSheets.find((s) => s.name === activeSheetName) ?? visibleSheets[0],
    [visibleSheets, activeSheetName],
  )

  const filteredRows = useMemo(() => {
    if (!activeSheet) return []
    const q = deferredSearch.toLowerCase()
    if (!q) return activeSheet.grid.rows
    return activeSheet.grid.rows.filter((row) => row.some((cell) => cell.toLowerCase().includes(q)))
  }, [activeSheet, deferredSearch])

  const navItems = useMemo(() => {
    return visibleSheets.map((sheet) => ({
      id: sheet.name,
      label: sheet.name,
      icon: Table,
      onSelect: () => setActiveSheetName(sheet.name),
    }))
  }, [visibleSheets])

  useEffect(() => {
    setNavItems(navItems)
  }, [setNavItems, navItems])

  useEffect(() => {
    if (activeSheet) {
      setStats({
        kind: 'spreadsheet',
        sheet: activeSheet.name,
        rows: filteredRows.length,
        cols: activeSheet.grid.colCount,
      })
    } else {
      setStats(null)
    }
  }, [setStats, activeSheet, filteredRows.length])

  const { columns, getCellContent, onColumnResize, onItemHovered, theme } = useSpreadsheetGrid({
    rows: filteredRows,
    colCount: activeSheet?.grid.colCount ?? 0,
    colWidthsPx: activeSheet?.grid.colWidthsPx,
    resetKey: activeSheetName ?? undefined,
  })

  const gridFind = useGridFind(filteredRows)
  useEffect(() => {
    registerFind(gridFind.open)
    return () => registerFind(null)
  }, [registerFind, gridFind.open])

  if (workbookState.status === 'error') {
    return <div className="spreadsheet-viewer__error">{workbookState.error}</div>
  }

  if (workbookState.status === 'loading') {
    return <ViewerLoading format={file.format} />
  }

  if (sheets.length === 0) {
    return <div className="spreadsheet-viewer" />
  }

  return (
    <div className="spreadsheet-viewer">
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
      <div className="spreadsheet-viewer__toolbar">
        <div className="spreadsheet-viewer__stats">
          {activeSheet ? `${filteredRows.length} rows × ${activeSheet.grid.colCount} columns` : ''}
        </div>
        <div className="spreadsheet-viewer__toolbar-actions">
          {hasHiddenSheets && (
            <button
              type="button"
              className="spreadsheet-viewer__hidden-toggle"
              onClick={() => setShowHiddenSheets((prev) => !prev)}
              title={showHiddenSheets ? 'Hide hidden sheets' : 'Show hidden sheets'}
              aria-pressed={showHiddenSheets}
            >
              {showHiddenSheets ? <EyeOff size={14} /> : <Eye size={14} />}
              {showHiddenSheets ? 'Hide hidden sheets' : 'Show hidden sheets'}
            </button>
          )}
          <div className="spreadsheet-viewer__search">
            <input
              type="search"
              placeholder="Search rows..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
          </div>
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
                getCellsForSelection={true}
                columns={columns}
                rows={filteredRows.length}
                theme={theme}
                width="100%"
                height="100%"
                smoothScrollX
                smoothScrollY
                rowMarkers="number"
                // Fixed UX default, not derived from the file's own freeze-pane
                // metadata — see the "Frozen panes" note in spreadsheetGrid.ts.
                freezeColumns={1}
                headerHeight={36}
                rowHeight={32}
                onItemHovered={onItemHovered}
                onColumnResize={onColumnResize}
                gridSelection={gridFind.gridSelection}
                onGridSelectionChange={gridFind.onGridSelectionChange}
              />
            )}
          </Suspense>
        )}
      </div>
      <div className="spreadsheet-viewer__tabs">
        {visibleSheets.map((sheet) => (
          <button
            key={sheet.name}
            className={`spreadsheet-viewer__tab ${sheet.name === activeSheetName ? 'spreadsheet-viewer__tab--active' : ''} ${sheet.hidden ? 'spreadsheet-viewer__tab--hidden' : ''}`}
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
