import { memo, useCallback, useEffect, useMemo, useState, useDeferredValue, lazy, Suspense } from 'react'
import { FileSpreadsheet } from 'lucide-react'
import type { Item } from '@glideapps/glide-data-grid'

import type { ViewerProps } from '../formats/types'
import { useSetNavItems, useSetViewerStats, useRegisterViewerFind } from './shared/useViewerContext'
import { useSpreadsheetGrid } from './shared/useSpreadsheetGrid'
import { useGridFind } from './shared/useGridFind'
import { useGridCellAnnouncement } from './shared/useGridCellAnnouncement'
import { commitGridEdit, useGridBlankMargin } from './shared/useGridBlankMargin'
import { parseCsv } from './shared/csvParse'
import { ViewerLoading } from '../components/ViewerLoading'
import { SearchOverlay } from '../components/SearchOverlay'
import { createDocumentFromRows } from './spreadsheet/spreadsheetDocument'
import { useSpreadsheetEditor, type SpreadsheetSaveTarget } from './spreadsheet/useSpreadsheetEditor'
import { SpreadsheetEditToolbar } from './spreadsheet/SpreadsheetEditToolbar'
import type { SpreadsheetDocument } from './spreadsheet/spreadsheetDocument'
import { useTranslate } from '../i18n'
import './__styles__/viewer-spreadsheet.css'

// Lazy load the grid (shared/SpreadsheetDataEditor pulls in glide-data-grid and its CSS)
const LazyDataEditor = lazy(async () => {
  const mod = await import('./shared/SpreadsheetDataEditor')
  return { default: mod.SpreadsheetDataEditor }
})

type CsvData = {
  rows: string[][]
  colCount: number
  /**
   * NIGHT/text-roundtrip — SHEET-6 fix. The delimiter Papa Parse actually
   * detected/used for THIS file (e.g. `;` for a semicolon-delimited CSV),
   * not necessarily the format's nominal default. Threaded through to the
   * save target below so a semicolon-delimited `.csv` is saved back with a
   * semicolon instead of silently being rewritten as comma-delimited.
   */
  delimiter: string
}

type LoadState =
  | { status: 'loading' }
  | { status: 'ready'; data: CsvData }
  | { status: 'error'; error: string }

const CSV_SAVE_FORMATS = [
  { id: 'csv', labelKey: 'csv.format.csv' },
  { id: 'tsv', labelKey: 'csv.format.tsv' },
] as const

const DEFAULT_DELIMITER_FOR_FORMAT: Readonly<Record<string, string>> = { tsv: '\t', csv: ',' }

function targetFor(formatId: string, delimiter?: string): SpreadsheetSaveTarget {
  const resolvedDelimiter = delimiter ?? DEFAULT_DELIMITER_FOR_FORMAT[formatId] ?? ','
  return formatId === 'tsv'
    ? { kind: 'delimited', delimiter: resolvedDelimiter, extension: 'tsv', filterName: 'Tab-Separated Values' }
    : { kind: 'delimited', delimiter: resolvedDelimiter, extension: 'csv', filterName: 'Comma-Separated Values' }
}

function CsvViewerBase({ file }: ViewerProps) {
  const t = useTranslate()
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
        // SHEET-6 — `result.meta.delimiter` is what Papa actually detected
        // (auto-detected for `.csv`, or the forced `'\t'` for `.tsv` above),
        // not necessarily the format's nominal default; falls back to the
        // format default only in the pathological case of a single-column/
        // single-row file where Papa reports an empty delimiter.
        const detectedDelimiter = result.meta.delimiter || (file.format === 'tsv' ? '\t' : ',')
        setState({ status: 'ready', data: { rows, colCount, delimiter: detectedDelimiter } })
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

  const initialDocument = useMemo<SpreadsheetDocument | null>(
    () => (data ? createDocumentFromRows(data.rows, data.colCount) : null),
    [data],
  )

  // SHEET-6 — use the delimiter actually detected in the loaded file (once
  // known) instead of the format's nominal default, so a semicolon-
  // delimited `.csv` saves back with a semicolon.
  const defaultTarget = useMemo<SpreadsheetSaveTarget>(
    () => targetFor(file.format === 'tsv' ? 'tsv' : 'csv', data?.delimiter),
    [file.format, data?.delimiter],
  )

  const editor = useSpreadsheetEditor(initialDocument, file.path, defaultTarget)
  const sheet = editor.document.sheets[0]

  // `filteredRowIndices[i]` is the actual (unfiltered) sheet row index grid
  // row `i` corresponds to — see SpreadsheetViewer's identical mapping for
  // why every grid-space row must be translated through this before being
  // used as a document row index (a confirmed data-corruption bug otherwise:
  // editing/pasting/inserting-around a row while "Search rows..." has
  // dropped non-matching rows would silently target the wrong sheet row).
  const filteredRowIndices = useMemo(() => {
    if (!sheet) return []
    const q = deferredSearch.toLowerCase()
    if (!q) return sheet.rows.map((_, i) => i)
    const indices: number[] = []
    sheet.rows.forEach((row, i) => {
      if (row.some((cell) => String(cell).toLowerCase().includes(q))) indices.push(i)
    })
    return indices
  }, [sheet, deferredSearch])

  const filteredRows = useMemo(
    () => filteredRowIndices.map((i) => sheet!.rows[i]),
    [filteredRowIndices, sheet],
  )
  // See SpreadsheetViewer's identical mapping and useSpreadsheetGrid's own
  // header: a CSV/TSV cell can carry a formula too, once the user types
  // "=..." into it, so this must reach the grid the same way.
  const filteredFormulas = useMemo(
    () => filteredRowIndices.map((i) => sheet!.formulas[i]),
    [filteredRowIndices, sheet],
  )

  useEffect(() => {
    if (sheet) {
      setStats({
        kind: 'spreadsheet',
        sheet: 'data',
        rows: filteredRows.length,
        cols: sheet.colCount,
      })
    } else {
      setStats(null)
    }
  }, [setStats, sheet, filteredRows.length])

  // Blank rows/columns past the data so empty cells can be typed into (USR-17).
  const { gridRowCount, gridColCount, sheetRowForGridRow } = useGridBlankMargin({
    bodyRowIndices: filteredRowIndices,
    colCount: sheet?.colCount ?? 0,
    isFiltering: deferredSearch !== '',
  })

  const handleCellEdited = useCallback(
    (row: number, col: number, rawText: string) => {
      const sheetRow = sheetRowForGridRow(row)
      if (!sheet || sheetRow === undefined) return
      commitGridEdit(editor, 0, sheet, sheetRow, col, rawText)
    },
    [editor, sheet, sheetRowForGridRow],
  )

  const { columns, getCellContent, onColumnResize, onItemHovered, theme, onCellEdited } = useSpreadsheetGrid({
    rows: filteredRows,
    colCount: gridColCount,
    onCellEdited: handleCellEdited,
    formulas: filteredFormulas,
  })

  const gridFind = useGridFind(filteredRows)
  useEffect(() => {
    registerFind(gridFind.open)
    return () => registerFind(null)
  }, [registerFind, gridFind.open])

  // A11Y pass 3 — see useGridCellAnnouncement's own doc comment: the canvas
  // grid has no cells for a screen reader to land on, so this is the
  // current-cell equivalent of aria-activedescendant for it.
  const cellAnnouncement = useGridCellAnnouncement(filteredRows, gridFind.gridSelection)

  const selection = useMemo(() => {
    const cell = gridFind.gridSelection?.current?.cell
    if (!cell) return null
    const sheetRow = sheetRowForGridRow(cell[1])
    return sheetRow === undefined ? null : { row: sheetRow, col: cell[0] }
  }, [gridFind.gridSelection, sheetRowForGridRow])

  const handleGridPaste = useCallback(
    (target: Item, values: readonly (readonly string[])[]): boolean => {
      const [col, row] = target
      const sheetRow = sheetRowForGridRow(row)
      if (sheetRow === undefined) return false
      // See SpreadsheetViewer's identical note: anchors correctly under an
      // active search filter, but a multi-row paste still targets contiguous
      // rows from that anchor.
      editor.pasteRange(0, sheetRow, col, values.map((r) => [...r]))
      return false
    },
    [editor, sheetRowForGridRow],
  )

  const saveFormats = useMemo(() => CSV_SAVE_FORMATS.map((f) => ({ id: f.id, label: t(f.labelKey) })), [t])

  const handleSaveAs = useCallback((formatId: string) => void editor.handleSaveAs(targetFor(formatId)), [editor])

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
          {t('spreadsheet.dimensions', {
            rows: t('spreadsheet.rowsCount', { count: filteredRows.length }),
            cols: t('spreadsheet.colsCount', { count: sheet?.colCount ?? 0 }),
          })}
        </div>
        <div className="csv-viewer__search">
          <input
            type="search"
            placeholder={t('spreadsheet.searchRowsPlaceholder')}
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
        </div>
      </div>
      <SpreadsheetEditToolbar
        canUndo={editor.canUndo}
        canRedo={editor.canRedo}
        onUndo={editor.undo}
        onRedo={editor.redo}
        selection={selection}
        onInsertRowAbove={(row) => editor.insertRowAt(0, row)}
        onDeleteRow={(row) => editor.deleteRowAt(0, row)}
        onInsertColumnLeft={(col) => editor.insertColumnAt(0, col)}
        onDeleteColumn={(col) => editor.deleteColumnAt(0, col)}
        onPaste={(row, col, values) => editor.pasteRange(0, row, col, values)}
        onSave={() => void editor.handleSave()}
        saveFormats={saveFormats}
        onSaveAs={handleSaveAs}
      />
      {editor.saveError && <div className="csv-viewer__error">{editor.saveError}</div>}
      {/* SHEET-4 — see SpreadsheetViewer for why this is a status, not an alert. */}
      {editor.saveWarning && (
        <div className="csv-viewer__save-warning" role="status">
          {editor.saveWarning}
        </div>
      )}
      <div
        className="csv-viewer__grid"
        role="group"
        aria-label={
          sheet
            ? t('csv.gridAria', {
                dimensions: t('spreadsheet.dimensions', {
                  rows: t('spreadsheet.rowsCount', { count: filteredRows.length }),
                  cols: t('spreadsheet.colsCount', { count: sheet.colCount }),
                }),
              })
            : undefined
        }
      >
        {!sheet ? null : search && filteredRows.length === 0 ? (
          <div className="csv-viewer__empty">
            <FileSpreadsheet size={48} />
            <p>{t('spreadsheet.noSearchResults')}</p>
          </div>
        ) : (
          <Suspense fallback={null}>
            <LazyDataEditor
              getCellContent={getCellContent}
              getCellsForSelection={true}
              columns={columns}
              rows={gridRowCount}
              theme={theme}
              width="100%"
              height="100%"
              smoothScrollX
              smoothScrollY
              rowMarkers="number"
              // CSV/TSV have no file-level freeze-pane concept at all (that's
              // an xlsx/ods-only feature — see shared/spreadsheetGrid.ts).
              freezeColumns={1}
              headerHeight={36}
              rowHeight={32}
              onItemHovered={onItemHovered}
              onColumnResize={onColumnResize}
              onCellEdited={onCellEdited}
              onPaste={handleGridPaste}
              gridSelection={gridFind.gridSelection}
              onGridSelectionChange={gridFind.onGridSelectionChange}
            />
          </Suspense>
        )}
        {/* A11Y pass 3 — the canvas grid has no cells of its own to announce;
            see useGridCellAnnouncement's doc comment. */}
        <span className="visually-hidden" role="status" aria-live="polite">
          {cellAnnouncement}
        </span>
      </div>
    </div>
  )
}

export const CsvViewer = memo(CsvViewerBase)
