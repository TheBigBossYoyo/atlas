import { memo, useCallback, useEffect, useMemo, useState, lazy, Suspense, useDeferredValue } from 'react'
import { Table, FileSpreadsheet, Eye, EyeOff, Plus, X } from 'lucide-react'
import type { GridColumn, Item, Rectangle } from '@glideapps/glide-data-grid'

import type { ViewerProps } from '../formats/types'
import { useSetNavItems, useSetViewerStats, useRegisterViewerFind } from './shared/useViewerContext'
import { useSpreadsheetWorkbook } from './shared/useSpreadsheetWorkbook'
import { useSpreadsheetGrid } from './shared/useSpreadsheetGrid'
import { useGridFind } from './shared/useGridFind'
import { commitGridEdit, useGridBlankMargin } from './shared/useGridBlankMargin'
import { ViewerLoading } from '../components/ViewerLoading'
import { SearchOverlay } from '../components/SearchOverlay'
import { createDocument } from './spreadsheet/spreadsheetDocument'
import { useSpreadsheetEditor, type SpreadsheetSaveTarget } from './spreadsheet/useSpreadsheetEditor'
import { bookTypeForExtension } from './spreadsheet/spreadsheetWrite'
import { SpreadsheetEditToolbar } from './spreadsheet/SpreadsheetEditToolbar'
import { FrozenRowsStrip } from './spreadsheet/FrozenRowsStrip'
import { isTableHeaderCell } from './spreadsheet/spreadsheetTables'
import { useTranslate } from '../i18n'
import './__styles__/viewer-spreadsheet.css'

// Lazy load the grid (shared/SpreadsheetDataEditor pulls in glide-data-grid and its CSS)
const LazyDataEditor = lazy(async () => {
  const mod = await import('./shared/SpreadsheetDataEditor')
  return { default: mod.SpreadsheetDataEditor }
})

/** Fixed so FrozenRowsStrip's spacer can line up exactly with the grid's own row-number column (T4/DAT-10 remainder). */
const ROW_MARKER_WIDTH_PX = 44

type WorkbookFormatOption = {
  readonly id: string
  readonly labelKey: string
  readonly extension: string
  readonly filterName: string
}

/**
 * Save As format choices offered for the xlsx/ods family (the plan's "Save
 * As with format choice"). `filterName` is the native OS Save dialog's
 * filter description (an Electron/Windows-level string, not one of this
 * array's own JSX) — left in English rather than plumbed through i18n.
 */
const WORKBOOK_SAVE_FORMATS: ReadonlyArray<WorkbookFormatOption> = [
  { id: 'xlsx', labelKey: 'spreadsheet.format.xlsx', extension: 'xlsx', filterName: 'Excel Workbook' },
  { id: 'xlsm', labelKey: 'spreadsheet.format.xlsm', extension: 'xlsm', filterName: 'Excel Macro-Enabled Workbook' },
  { id: 'xlsb', labelKey: 'spreadsheet.format.xlsb', extension: 'xlsb', filterName: 'Excel Binary Workbook' },
  { id: 'xls', labelKey: 'spreadsheet.format.xls', extension: 'xls', filterName: 'Excel 97-2003 Workbook' },
  { id: 'ods', labelKey: 'spreadsheet.format.ods', extension: 'ods', filterName: 'OpenDocument Spreadsheet' },
  { id: 'fods', labelKey: 'spreadsheet.format.fods', extension: 'fods', filterName: 'Flat OpenDocument Spreadsheet' },
]

/** `useSpreadsheetGrid`'s own columns always set an explicit `width` (see that hook), but `GridColumn`'s library type also allows a width-less `AutoGridColumn` — narrow defensively rather than asserting. */
function columnWidthOf(column: GridColumn): number {
  return 'width' in column && typeof column.width === 'number' ? column.width : 120
}

function extensionOf(path: string): string {
  const base = path.replace(/\\/g, '/').split('/').pop() ?? path
  const dot = base.lastIndexOf('.')
  return dot > 0 ? base.slice(dot + 1).toLowerCase() : ''
}

function workbookTargetFor(formatId: string): SpreadsheetSaveTarget {
  const option = WORKBOOK_SAVE_FORMATS.find((f) => f.id === formatId) ?? WORKBOOK_SAVE_FORMATS[0]
  return {
    kind: 'workbook',
    bookType: bookTypeForExtension(option.extension),
    extension: option.extension,
    filterName: option.filterName,
  }
}

/**
 * Legacy/flat formats scoped as "view + save-as-xlsx only" (owner's plan):
 * SheetJS can technically write `.xls`/`.xlsb`/`.fods` back out (verified
 * directly against the library — see `spreadsheetWrite.ts`'s header), but
 * the plan deliberately doesn't offer silently re-encoding the user's
 * original legacy file in place. Save always defaults to `.xlsx` and never
 * seeds the original path as an overwrite target for these three.
 */
const LEGACY_SAVE_AS_XLSX_ONLY: ReadonlySet<string> = new Set(['xls', 'xlsb', 'fods'])

function SpreadsheetViewerBase({ file }: ViewerProps) {
  const t = useTranslate()
  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()
  const registerFind = useRegisterViewerFind()

  const [activeSheetName, setActiveSheetName] = useState<string | null>(null)
  const [showHiddenSheets, setShowHiddenSheets] = useState(false)
  const [search, setSearch] = useState('')
  const deferredSearch = useDeferredValue(search)
  const [renamingSheet, setRenamingSheet] = useState<string | null>(null)
  const [gridTranslateX, setGridTranslateX] = useState(0)

  const buffer = file.kind === 'binary' ? file.content : null
  const workbookState = useSpreadsheetWorkbook(buffer)

  // Built once per successful parse — see useSpreadsheetEditor's own
  // "hydrate once" guard for why recomputing this on every render (a cheap
  // per-sheet wrap, not a deep clone — see spreadsheetDocument.ts) is safe.
  const initialDocument = useMemo(
    () => (workbookState.status === 'ready' ? createDocument(workbookState.sheets) : null),
    [workbookState],
  )

  const fileExtension = useMemo(() => extensionOf(file.path) || 'xlsx', [file.path])
  const isLegacySaveAsOnly = LEGACY_SAVE_AS_XLSX_ONLY.has(fileExtension)
  // Plain Save keeps the file's OWN original extension/format — UNLESS it's
  // one of the "view + save-as-xlsx only" legacy/flat formats above, which
  // default to .xlsx instead. Save As always lets the user pick a different
  // format either way (see the toolbar's format <select>).
  const defaultSaveTarget = useMemo<SpreadsheetSaveTarget>(
    () => workbookTargetFor(isLegacySaveAsOnly ? 'xlsx' : fileExtension),
    [fileExtension, isLegacySaveAsOnly],
  )

  const editor = useSpreadsheetEditor(initialDocument, file.path, defaultSaveTarget, !isLegacySaveAsOnly, buffer)
  const sheets = editor.document.sheets

  const hasHiddenSheets = useMemo(() => sheets.some((s) => s.hidden), [sheets])
  const visibleSheets = useMemo(
    () => (showHiddenSheets ? sheets : sheets.filter((s) => !s.hidden)),
    [sheets, showHiddenSheets],
  )

  // Pick an initial/fallback active sheet once the workbook is ready, or
  // when the current selection is no longer in the visible list. Pure
  // derivation of `visibleSheets`/`activeSheetName`, so this uses the
  // render-time "adjust state" idiom (see ViewerContext.tsx) rather than an effect.
  if (visibleSheets.length > 0 && !visibleSheets.some((s) => s.name === activeSheetName)) {
    setActiveSheetName(visibleSheets[0].name)
  }

  const activeSheet = useMemo(
    () => visibleSheets.find((s) => s.name === activeSheetName) ?? visibleSheets[0],
    [visibleSheets, activeSheetName],
  )
  const activeSheetIndex = useMemo(
    () => (activeSheet ? sheets.findIndex((s) => s.name === activeSheet.name) : -1),
    [sheets, activeSheet],
  )

  // `filteredRowIndices[i]` is the ACTUAL (unfiltered) sheet row index that
  // grid-space row `i` corresponds to. When the "Search rows..." filter is
  // active it drops non-matching rows, so grid row `i` and sheet row `i` are
  // no longer the same thing — every place below that turns a grid-space row
  // back into a document edit (cell edit, paste, and the toolbar's
  // insert/delete row/column, which read `selection.row`) MUST go through
  // this mapping rather than assuming identity. Getting this wrong doesn't
  // error — it silently edits a different, arbitrary row than the one the
  // user is looking at (a confirmed data-corruption bug this fixes: search
  // for a value, edit what looks like the matching row, and — before this
  // fix — the edit landed on whatever row shares that same position in the
  // FILTERED list instead).
  const filteredRowIndices = useMemo(() => {
    if (!activeSheet) return []
    const q = deferredSearch.toLowerCase()
    if (!q) return activeSheet.rows.map((_, i) => i)
    const indices: number[] = []
    activeSheet.rows.forEach((row, i) => {
      if (row.some((cell) => cell.toLowerCase().includes(q))) indices.push(i)
    })
    return indices
  }, [activeSheet, deferredSearch])

  const filteredRows = useMemo(
    () => filteredRowIndices.map((i) => activeSheet!.rows[i]),
    [filteredRowIndices, activeSheet],
  )
  // Formulas, mapped through the exact same filter as `filteredRows` (see
  // `useSpreadsheetGrid`'s own header for why this must reach the grid: a
  // formula cell's edit overlay needs its literal `=<formula>` text, not the
  // already-computed value `filteredRows` carries).
  const filteredFormulas = useMemo(
    () => filteredRowIndices.map((i) => activeSheet!.formulas[i]),
    [filteredRowIndices, activeSheet],
  )

  // Frozen ROWS (T4/DAT-10 remainder — see FrozenRowsStrip's header for why
  // this isn't a second DataEditor) are rendered as a fixed strip ABOVE the
  // main scrollable grid, which therefore only ever renders the REMAINING
  // rows — never both, which would show the frozen row twice. Disabled
  // while a search filter is active: the filter can drop/reorder rows
  // entirely, and "the top N rows of a reshuffled result set" is not a
  // meaningful thing to freeze.
  const frozenRowCount = !deferredSearch ? (activeSheet?.freeze?.rows ?? 0) : 0
  const bodyRows = frozenRowCount > 0 ? filteredRows.slice(frozenRowCount) : filteredRows
  // Sheet-absolute row index for each row of `bodyRows`, by the same
  // position — see `filteredRowIndices` above. Slicing this array in lockstep
  // with `bodyRows` (rather than re-deriving it from `frozenRowCount`
  // elsewhere) keeps the two arrays' indices aligned by construction.
  const bodyRowIndices = frozenRowCount > 0 ? filteredRowIndices.slice(frozenRowCount) : filteredRowIndices
  const bodyFormulas = frozenRowCount > 0 ? filteredFormulas.slice(frozenRowCount) : filteredFormulas
  const frozenRowsData = frozenRowCount > 0 ? activeSheet!.rows.slice(0, frozenRowCount) : []
  const frozenFormulasData = frozenRowCount > 0 ? activeSheet!.formulas.slice(0, frozenRowCount) : []

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
        cols: activeSheet.colCount,
      })
    } else {
      setStats(null)
    }
  }, [setStats, activeSheet, filteredRows.length])

  const isFiltering = deferredSearch !== ''
  // Blank rows/columns past the data (USR-17); rows past the data map one-to-one after it.
  const { gridRowCount, gridColCount, sheetRowForGridRow } = useGridBlankMargin({
    bodyRowIndices,
    colCount: activeSheet?.colCount ?? 0,
    isFiltering,
    rowOffset: frozenRowCount,
  })

  const handleCellEdited = useCallback(
    (row: number, col: number, rawText: string) => {
      const sheetRow = sheetRowForGridRow(row)
      if (activeSheetIndex < 0 || !activeSheet || sheetRow === undefined) return
      commitGridEdit(editor, activeSheetIndex, activeSheet, sheetRow, col, rawText)
    },
    [editor, activeSheetIndex, activeSheet, sheetRowForGridRow],
  )

  const isHeaderCell = useCallback(
    (row: number, col: number): boolean => {
      const sheetRow = sheetRowForGridRow(row)
      return sheetRow !== undefined && isTableHeaderCell(activeSheet?.tables, sheetRow, col)
    },
    [sheetRowForGridRow, activeSheet],
  )

  const { columns, getCellContent, onColumnResize, onItemHovered, theme, onCellEdited } = useSpreadsheetGrid({
    rows: bodyRows,
    colCount: gridColCount,
    colWidthsPx: activeSheet?.colWidthsPx,
    resetKey: activeSheetName ?? undefined,
    onCellEdited: handleCellEdited,
    formulas: bodyFormulas,
    isHeaderCell: activeSheet?.tables?.length ? isHeaderCell : undefined,
  })

  const gridFind = useGridFind(bodyRows)
  useEffect(() => {
    registerFind(gridFind.open)
    return () => registerFind(null)
  }, [registerFind, gridFind.open])

  // The currently-selected cell, sheet-absolute (mapped through
  // `bodyRowIndices`, which already accounts for both frozen rows AND an
  // active search filter — see that array's own comment) — read straight off
  // gridFind's own selection state rather than tracking a parallel one, since
  // it already reflects the live selection regardless of whether Find is open
  // (see that hook's `onGridSelectionChange`).
  const selection = useMemo(() => {
    const cell = gridFind.gridSelection?.current?.cell
    if (!cell) return null
    const sheetRow = sheetRowForGridRow(cell[1])
    return sheetRow === undefined ? null : { row: sheetRow, col: cell[0] }
  }, [gridFind.gridSelection, sheetRowForGridRow])

  const handleGridPaste = useCallback(
    (target: Item, values: readonly (readonly string[])[]): boolean => {
      if (activeSheetIndex < 0) return false
      const [col, row] = target
      const sheetRow = sheetRowForGridRow(row)
      if (sheetRow === undefined) return false
      // Anchors the paste at the correct sheet row even under an active
      // search filter (see `bodyRowIndices`). A multi-row paste while
      // filtered still writes to CONTIGUOUS rows from that anchor — matching
      // filtered rows are not, in general, contiguous in the sheet — since
      // `pasteRange` itself has no notion of a non-contiguous target range;
      // pasting a single row/cell (by far the common case) is unaffected.
      editor.pasteRange(
        activeSheetIndex,
        sheetRow,
        col,
        values.map((r) => [...r]),
      )
      return false // handled manually (can grow the sheet) — see DataEditor's onPaste docs.
    },
    [editor, activeSheetIndex, sheetRowForGridRow],
  )

  const handleVisibleRegionChanged = useCallback((_range: Rectangle, tx: number) => {
    setGridTranslateX(tx)
  }, [])

  const handleFrozenRowCommit = useCallback(
    (row: number, col: number, text: string) => {
      if (activeSheetIndex < 0) return
      editor.setCellValue(activeSheetIndex, row, col, text)
    },
    [editor, activeSheetIndex],
  )

  const commitRename = useCallback(
    (sheetName: string, newName: string) => {
      const idx = sheets.findIndex((s) => s.name === sheetName)
      const trimmed = newName.trim()
      if (idx < 0 || !trimmed || trimmed === sheetName) return
      // Mirrors `renameSheetOp`'s own duplicate-name guard (spreadsheetDocument.ts)
      // so this only follows the active-tab pointer when the rename will
      // actually take effect — see the comment below on why this pointer
      // must move at all.
      if (sheets.some((s, i) => i !== idx && s.name === trimmed)) return
      editor.renameSheet(idx, trimmed)
      // The active sheet is tracked by NAME (`activeSheetName`), and a
      // rename changes exactly that key out from under it. Without this, the
      // render-time "pick a fallback active sheet" adjustment above (which
      // only knows "the current activeSheetName no longer exists", not "it
      // was renamed") falls back to `visibleSheets[0]` — silently switching
      // the visible sheet to whichever tab happens to be first, unless the
      // renamed sheet already WAS the first one (the only case the original
      // single-sheet test for this happened to cover). Renaming the active
      // sheet must keep IT active, under its new name.
      if (sheetName === activeSheetName) {
        setActiveSheetName(trimmed)
      }
    },
    [sheets, editor, activeSheetName],
  )

  const saveFormats = useMemo(
    () => WORKBOOK_SAVE_FORMATS.map((f) => ({ id: f.id, label: t(f.labelKey) })),
    [t],
  )

  const handleSaveAs = useCallback(
    (formatId: string) => {
      void editor.handleSaveAs(workbookTargetFor(formatId))
    },
    [editor],
  )

  if (workbookState.status === 'error') {
    return <div className="spreadsheet-viewer__error">{workbookState.error}</div>
  }

  if (workbookState.status === 'loading') {
    return <ViewerLoading format={file.format} />
  }

  if (sheets.length === 0) {
    return <div className="spreadsheet-viewer" />
  }

  const rowMarkers: 'number' | { kind: 'number'; width: number; startIndex: number } =
    frozenRowCount > 0 ? { kind: 'number', width: ROW_MARKER_WIDTH_PX, startIndex: frozenRowCount + 1 } : 'number'

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
          {activeSheet
            ? t('spreadsheet.dimensions', {
                rows: t('spreadsheet.rowsCount', { count: filteredRows.length }),
                cols: t('spreadsheet.colsCount', { count: activeSheet.colCount }),
              })
            : ''}
        </div>
        <div className="spreadsheet-viewer__toolbar-actions">
          {hasHiddenSheets && (
            <button
              type="button"
              className="spreadsheet-viewer__hidden-toggle"
              onClick={() => setShowHiddenSheets((prev) => !prev)}
              title={showHiddenSheets ? t('spreadsheet.hideHiddenSheets') : t('spreadsheet.showHiddenSheets')}
              aria-pressed={showHiddenSheets}
            >
              {showHiddenSheets ? <EyeOff size={14} /> : <Eye size={14} />}
              {showHiddenSheets ? t('spreadsheet.hideHiddenSheets') : t('spreadsheet.showHiddenSheets')}
            </button>
          )}
          <div className="spreadsheet-viewer__search">
            <input
              type="search"
              placeholder={t('spreadsheet.searchRowsPlaceholder')}
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
          </div>
        </div>
      </div>
      <SpreadsheetEditToolbar
        canUndo={editor.canUndo}
        canRedo={editor.canRedo}
        onUndo={editor.undo}
        onRedo={editor.redo}
        selection={selection}
        onInsertRowAbove={(row) => activeSheetIndex >= 0 && editor.insertRowAt(activeSheetIndex, row)}
        onDeleteRow={(row) => activeSheetIndex >= 0 && editor.deleteRowAt(activeSheetIndex, row)}
        onInsertColumnLeft={(col) => activeSheetIndex >= 0 && editor.insertColumnAt(activeSheetIndex, col)}
        onDeleteColumn={(col) => activeSheetIndex >= 0 && editor.deleteColumnAt(activeSheetIndex, col)}
        onPaste={(row, col, values) => activeSheetIndex >= 0 && editor.pasteRange(activeSheetIndex, row, col, values)}
        onSave={() => void editor.handleSave()}
        saveFormats={saveFormats}
        onSaveAs={handleSaveAs}
      />
      {editor.saveError && <div className="spreadsheet-viewer__save-error">{editor.saveError}</div>}
      {frozenRowCount > 0 && (
        <FrozenRowsStrip
          rows={frozenRowsData}
          formulas={frozenFormulasData}
          columnWidthsPx={columns.map(columnWidthOf)}
          rowMarkerWidthPx={ROW_MARKER_WIDTH_PX}
          translateXPx={gridTranslateX}
          onCommit={handleFrozenRowCommit}
        />
      )}
      <div className="spreadsheet-viewer__grid">
        {isFiltering && filteredRows.length === 0 ? (
          <div className="spreadsheet-viewer__empty">
            <FileSpreadsheet size={48} />
            <p>{search ? t('spreadsheet.noSearchResults') : t('spreadsheet.emptySheet')}</p>
          </div>
        ) : (
          <Suspense fallback={null}>
            {activeSheet && (
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
                rowMarkers={rowMarkers}
                freezeColumns={activeSheet?.freeze?.cols ?? 1}
                headerHeight={36}
                rowHeight={32}
                onItemHovered={onItemHovered}
                onColumnResize={onColumnResize}
                onCellEdited={onCellEdited}
                onPaste={handleGridPaste}
                onVisibleRegionChanged={handleVisibleRegionChanged}
                gridSelection={gridFind.gridSelection}
                onGridSelectionChange={gridFind.onGridSelectionChange}
              />
            )}
          </Suspense>
        )}
      </div>
      <div className="spreadsheet-viewer__tabs">
        {visibleSheets.map((sheet) => (
          <div
            key={sheet.name}
            className={`spreadsheet-viewer__tab-wrap ${sheet.hidden ? 'spreadsheet-viewer__tab--hidden' : ''}`}
          >
            {renamingSheet === sheet.name ? (
              <input
                className="spreadsheet-viewer__tab-rename"
                defaultValue={sheet.name}
                autoFocus
                aria-label={t('spreadsheet.renameSheetAria', { name: sheet.name })}
                onBlur={(e) => {
                  commitRename(sheet.name, e.currentTarget.value)
                  setRenamingSheet(null)
                }}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') e.currentTarget.blur()
                  if (e.key === 'Escape') setRenamingSheet(null)
                }}
              />
            ) : (
              <button
                type="button"
                className={`spreadsheet-viewer__tab ${sheet.name === activeSheetName ? 'spreadsheet-viewer__tab--active' : ''}`}
                onClick={() => {
                  setActiveSheetName(sheet.name)
                  setSearch('')
                }}
                onDoubleClick={() => setRenamingSheet(sheet.name)}
                title={t('spreadsheet.renameSheetTitle')}
              >
                <Table size={14} />
                {sheet.name}
              </button>
            )}
            {sheets.length > 1 && renamingSheet !== sheet.name && (
              <button
                type="button"
                className="spreadsheet-viewer__tab-delete"
                aria-label={t('spreadsheet.deleteSheetAria', { name: sheet.name })}
                title={t('spreadsheet.deleteSheetTitle', { name: sheet.name })}
                onClick={() => {
                  const idx = sheets.findIndex((s) => s.name === sheet.name)
                  if (idx >= 0) editor.deleteSheet(idx)
                }}
              >
                <X size={12} />
              </button>
            )}
          </div>
        ))}
        <button
          type="button"
          className="spreadsheet-viewer__tab-add"
          onClick={() => editor.addSheet()}
          aria-label={t('spreadsheet.addSheetAria')}
          title={t('spreadsheet.addSheetTitle')}
        >
          <Plus size={14} />
        </button>
      </div>
    </div>
  )
}

export const SpreadsheetViewer = memo(SpreadsheetViewerBase)
