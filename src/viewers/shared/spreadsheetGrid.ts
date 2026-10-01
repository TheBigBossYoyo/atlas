/**
 * Shared spreadsheet-grid data model (T1/T3/T4/T5, DAT-04/08/09/10/11).
 *
 * Both SpreadsheetViewer (xlsx/ods) and CsvViewer (csv/tsv) render the same
 * shape of grid data through the same glide-data-grid instance. This module
 * is the single place that:
 *  - parses a `xlsx`/`ods` workbook buffer into a plain, structured-clone
 *    -friendly `SheetGrid[]` (safe to postMessage across a Worker boundary
 *    for T2), formatting every cell via `formatCellText` (T1) instead of
 *    XLSX's own `sheet_to_json` (whose `raw:false` still blanks out most
 *    formula-error cells — see the comment on `sheetToGrid` below);
 *  - backfills merged-cell ranges (T4/DAT-09) so a merged block's value
 *    doesn't appear to vanish outside its top-left cell;
 *  - carries column widths / row heights from `!cols`/`!rows` (T4/DAT-10);
 *  - reports each sheet's `hidden` flag from the workbook's own sheet
 *    visibility state (T4/DAT-11) so the UI can filter/toggle it.
 *
 * Frozen panes (T4/DAT-10's other sub-item): this `xlsx` build never parses a
 * sheet's `<pane>`/`xSplit`/`ySplit` XML at all (confirmed by grepping the
 * bundled `xlsx.js` — every freeze/split/pane-related case in its own
 * settings parser is a no-op `break`), so there is no `ws['!freeze']`
 * -equivalent field to read here. `spreadsheet/spreadsheetPanes.ts` reads it
 * separately, by re-unzipping the raw file buffer with JSZip and parsing the
 * relevant `xl/worksheets/sheetN.xml` with `DOMParser` — see that module's
 * header for why this is a distinct, buffer-driven step (not part of
 * `sheetToGrid`/`parseWorkbookBuffer` above) and how it stays off the main
 * thread for large files. `attachFrozenPanes` below merges its result back
 * onto this module's own `ParsedSheet[]` by sheet name.
 *
 * `XLSX.read`'s `cellStyles:true` is kept even though T3 asks to drop
 * "expensive, immediately-discarded" parse options: in this SheetJS build,
 * `!cols` (column widths) and `!rows` (row heights) are ONLY populated when
 * `cellStyles` is requested (see `parse_ws_xml_cols`/`parse_ws_xml_data` in
 * xlsx.js) — so it is no longer "immediately discarded", it is the only way
 * to get T4's column/row sizing. `sheetStubs` is dropped: nothing in Atlas
 * consumes stub (`t:'z'`) cells. `cellFormula` is NOT dropped (its default is
 * already `true`) — wave 3 editing reads each cell's `.f` into `formulas`
 * below so the editable-document model can tell "this cell is a formula"
 * apart from a plain value without re-walking the worksheet a second time.
 */
import * as XLSX from 'xlsx'
import { formatCellText } from './xlsxCellFormat'
import type { SheetTable } from '../spreadsheet/spreadsheetTables'
import type { SheetCellStyles, WorkbookCellStyles } from '../spreadsheet/xlsxCellStyles'
import { checkWorkbookZipBudgetSync, SpreadsheetZipBombError } from '../spreadsheet/spreadsheetZipBudget'
import { t } from '../../i18n'

export type MergeRange = {
  readonly r0: number
  readonly c0: number
  readonly r1: number
  readonly c1: number
}

export type SheetGrid = {
  readonly rows: ReadonlyArray<ReadonlyArray<string>>
  readonly colCount: number
  readonly merges: ReadonlyArray<MergeRange>
  /** Pixel width per column index, `undefined` where the file specifies none. */
  readonly colWidthsPx: ReadonlyArray<number | undefined>
  /** Pixel height per row index, `undefined` where the file specifies none. */
  readonly rowHeightsPx: ReadonlyArray<number | undefined>
  /**
   * The raw formula text (no leading `=`) behind each cell, `undefined` for
   * a plain-value cell. Carried alongside `rows`' already-formatted display
   * text (wave 3 editing) so the editable-document model
   * (`spreadsheet/spreadsheetDocument.ts`) can reconstruct "this cell is a
   * formula" without re-walking the worksheet a second time.
   */
  readonly formulas: ReadonlyArray<ReadonlyArray<string | undefined>>
  /**
   * The sheet's used range origin, in absolute sheet coordinates — `rows[0][0]`
   * IS this cell.
   *
   * SHEETFMT-1 needs it: cell formatting is read from the worksheet XML, where
   * every cell carries an absolute address (`B3`), while this grid is indexed
   * from the used range. A sheet whose data starts at B3 would otherwise have
   * every format landing a row and a column off.
   *
   * Optional, defaulting to 0 (= A1), so a hand-built grid — every spreadsheet
   * unit test constructs one — does not have to state an origin it has no
   * opinion about. `sheetToGrid` always sets it from the real `!ref`.
   */
  readonly originRow?: number
  readonly originCol?: number
}

/** Frozen-pane split, in leading column/row counts (T4/DAT-10 remainder). */
export type FrozenPanes = {
  readonly cols: number
  readonly rows: number
}

export type ParsedSheet = {
  readonly name: string
  readonly hidden: boolean
  readonly grid: SheetGrid
  /**
   * Present only when `spreadsheet/spreadsheetPanes.ts` found a `state="frozen"`
   * `<pane>` in this sheet's own XML (OOXML zip formats only — see that
   * module's header). `undefined` means "unknown/not applicable", NOT "no
   * freeze" — callers that want a UI default when this is absent should fall
   * back explicitly rather than treating `undefined` as `{cols:0,rows:0}`.
   */
  readonly freeze?: FrozenPanes
  /** Excel tables on this sheet (USR-17, OOXML only — see `spreadsheet/spreadsheetTables.ts`). */
  readonly tables?: ReadonlyArray<SheetTable>
  /** Worksheet part this sheet was parsed from (OOXML only) — lets a save write through the original file. */
  readonly sourcePath?: string
  /**
   * SHEETFMT-1 — the cell formatting this sheet carries (bold, fills, borders,
   * alignment, number formats), read from `xl/styles.xml` plus each cell's own
   * `s` index. `undefined` for any format that has no styles part to read
   * (`.csv`, `.xls`, `.ods`) and whenever that part could not be read, which is
   * the "render unstyled" case Atlas behaved as before this existed.
   */
  readonly cellStyles?: SheetCellStyles
}

/**
 * Attaches each sheet's cell formatting, by worksheet part path.
 *
 * Keyed by `sourcePath` rather than by index because that is what
 * `readWorkbookCellStyles` returns and what survives a sheet reorder — the same
 * reason `attachTables`/`attachFrozenPanes` key by sheet name.
 */
export function attachCellStyles(
  sheets: ReadonlyArray<ParsedSheet>,
  styles: WorkbookCellStyles | null,
): ParsedSheet[] {
  if (!styles) return [...sheets]
  return sheets.map(sheet => {
    const cellStyles = sheet.sourcePath ? styles.get(sheet.sourcePath) : undefined
    return cellStyles ? { ...sheet, cellStyles } : sheet
  })
}

/** The arguments `readWorkbookCellStyles` needs for each sheet, derived from sheets that already carry their `sourcePath`. */
export function cellStyleRequests(
  sheets: ReadonlyArray<ParsedSheet>,
): Array<{ sourcePath?: string; rowCount: number; colCount: number; offsetRow: number; offsetCol: number }> {
  return sheets.map(sheet => ({
    sourcePath: sheet.sourcePath,
    rowCount: sheet.grid.rows.length,
    colCount: sheet.grid.colCount,
    offsetRow: sheet.grid.originRow ?? 0,
    offsetCol: sheet.grid.originCol ?? 0,
  }))
}

/** Attaches each sheet's worksheet part path, by workbook order (USR-17 save-through-original). */
export function attachSheetSources(
  sheets: ReadonlyArray<ParsedSheet>,
  partPaths: ReadonlyArray<string>,
): ParsedSheet[] {
  return sheets.map((sheet, index) => {
    const sourcePath = partPaths[index]
    return sourcePath ? { ...sheet, sourcePath } : sheet
  })
}

/** Merges a sheet-name-keyed table map (from `readSheetTables`) onto already-parsed sheets. Pure/sync. */
export function attachTables(
  sheets: ReadonlyArray<ParsedSheet>,
  tableMap: Readonly<Record<string, ReadonlyArray<SheetTable>>>,
): ParsedSheet[] {
  return sheets.map((sheet) => {
    const tables = tableMap[sheet.name]
    return tables && tables.length > 0 ? { ...sheet, tables } : sheet
  })
}

/** Merges a sheet-name-keyed frozen-pane map (from `readFrozenPanes`) onto already-parsed sheets. Pure/sync. */
export function attachFrozenPanes(
  sheets: ReadonlyArray<ParsedSheet>,
  paneMap: Readonly<Record<string, FrozenPanes>>,
): ParsedSheet[] {
  return sheets.map((sheet) => {
    const freeze = paneMap[sheet.name]
    return freeze ? { ...sheet, freeze } : sheet
  })
}

const EMPTY_GRID: SheetGrid = {
  rows: [],
  colCount: 0,
  merges: [],
  colWidthsPx: [],
  rowHeightsPx: [],
  formulas: [], originRow: 0, originCol: 0,
}

/**
 * Converts one worksheet into a `SheetGrid`.
 *
 * Deliberately walks `!ref`'s range and reads each cell directly (via
 * `encode_cell`/`format_cell`) rather than calling `XLSX.utils.sheet_to_json`:
 * `sheet_to_json` special-cases `t:'e'` cells to `null`/`defval` *before*
 * formatting is ever considered (see its `make_json_row`), so even
 * `raw:false` still renders most formula errors as blank — the exact DAT-04
 * bug this rewrite fixes. A direct per-cell walk also gives us stable
 * (row, col) coordinates to backfill merges and index `!cols`/`!rows` by.
 */
export function sheetToGrid(ws: XLSX.WorkSheet): SheetGrid {
  const ref = ws['!ref']
  if (!ref) {
    return EMPTY_GRID
  }

  const range = XLSX.utils.decode_range(ref)
  const rowCount = range.e.r - range.s.r + 1
  const colCount = range.e.c - range.s.c + 1

  const rows: string[][] = []
  const formulas: (string | undefined)[][] = []
  for (let r = 0; r < rowCount; r++) {
    const row: string[] = new Array<string>(colCount).fill('')
    const formulaRow: (string | undefined)[] = new Array<string | undefined>(colCount).fill(undefined)
    for (let c = 0; c < colCount; c++) {
      const addr = XLSX.utils.encode_cell({ r: range.s.r + r, c: range.s.c + c })
      const cell = ws[addr] as XLSX.CellObject | undefined
      row[c] = formatCellText(cell)
      formulaRow[c] = cell?.f
    }
    rows.push(row)
    formulas.push(formulaRow)
  }

  const merges: MergeRange[] = []
  for (const merge of ws['!merges'] ?? []) {
    const r0 = merge.s.r - range.s.r
    const c0 = merge.s.c - range.s.c
    const r1 = merge.e.r - range.s.r
    const c1 = merge.e.c - range.s.c
    if (r0 < 0 || c0 < 0) continue
    merges.push({ r0, c0, r1, c1 })

    // Backfill (DAT-09): every cell in a merged range other than the
    // top-left one is genuinely empty in the file, so the merged value
    // would otherwise appear to vanish everywhere but that top-left cell.
    const text = rows[r0]?.[c0] ?? ''
    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) {
        if (r === r0 && c === c0) continue
        if (rows[r]) rows[r][c] = text
      }
    }
  }

  const colWidthsPx: (number | undefined)[] = []
  const cols = ws['!cols']
  if (cols) {
    for (let c = 0; c < colCount; c++) {
      colWidthsPx.push(cols[range.s.c + c]?.wpx)
    }
  }

  const rowHeightsPx: (number | undefined)[] = []
  const rowInfos = ws['!rows']
  if (rowInfos) {
    for (let r = 0; r < rowCount; r++) {
      rowHeightsPx.push(rowInfos[range.s.r + r]?.hpx)
    }
  }

  return { rows, colCount, merges, colWidthsPx, rowHeightsPx, formulas, originRow: range.s.r, originCol: range.s.c }
}

/** Sheet visibility state (0=visible, 1=hidden, 2=very hidden) → `hidden` (T4/DAT-11). */
function isSheetHidden(workbook: XLSX.WorkBook, index: number): boolean {
  const hidden = workbook.Workbook?.Sheets?.[index]?.Hidden
  return hidden === 1 || hidden === 2
}

/**
 * Parses an xlsx/ods/xls workbook buffer into one `ParsedSheet` per sheet,
 * in file order. Safe to call from a Worker (T2) or the main thread.
 *
 * SHEET-1 (security review follow-up): before handing `buffer` to
 * `XLSX.read` — which unzips and inflates the entire archive with no size
 * check of its own — this runs `checkWorkbookZipBudgetSync`, which refuses a
 * workbook whose zip central directory already declares an implausible
 * uncompressed size, before a single byte is inflated (see that function's
 * header on why it's a separate, synchronous sibling of `loadWorkbookZip`
 * rather than that same async check reused here). `parseWorkbookBuffer` is
 * the ONE path the viewer (`useSpreadsheetWorkbook.ts`), the parsing Worker
 * (`spreadsheetWorker.worker.ts`) and PDF/CSV export
 * (`utils/export/spreadsheetPdf.ts`) all go through — guarding it here, once,
 * is what makes every one of those inherit the guard, including the ordering
 * trap where this function used to run BEFORE the budget-guarded passthrough
 * helpers in the viewer/Worker code paths.
 */
export function parseWorkbookBuffer(buffer: ArrayBuffer): ParsedSheet[] {
  try {
    checkWorkbookZipBudgetSync(buffer)
  } catch (err) {
    if (err instanceof SpreadsheetZipBombError) {
      // A translated, plain-language message — `checkWorkbookZipBudgetSync`'s
      // own message is precise but developer-facing (byte counts, "zip
      // bomb"); `errors.library.tooLarge` is the existing catalogue entry
      // already used elsewhere in the app for "this file is too big to
      // safely process" (see `friendlyLibraryError.ts`), so this reuses it
      // rather than adding a new key.
      throw new Error(t('errors.library.tooLarge'))
    }
    throw err
  }

  // Wrapped in a Uint8Array rather than handed the raw ArrayBuffer directly:
  // XLSX.read's own `instanceof ArrayBuffer` auto-wrapping has been observed
  // to misdetect a real-file-read buffer as plain text in some hosts (Vitest's
  // jsdom test environment can end up with more than one ArrayBuffer realm in
  // play), producing one giant text cell instead of a parsed workbook. A
  // Uint8Array view sidesteps that detection path entirely and is otherwise
  // equivalent — same bytes, same offset/length.
  const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array', cellStyles: true })
  return workbook.SheetNames.map((name, index) => ({
    name,
    hidden: isSheetHidden(workbook, index),
    grid: sheetToGrid(workbook.Sheets[name]),
  }))
}
