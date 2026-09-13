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
 * `XLSX.read`'s `cellStyles:true` is kept even though T3 asks to drop
 * "expensive, immediately-discarded" parse options: in this SheetJS build,
 * `!cols` (column widths) and `!rows` (row heights) are ONLY populated when
 * `cellStyles` is requested (see `parse_ws_xml_cols`/`parse_ws_xml_data` in
 * xlsx.js) — so it is no longer "immediately discarded", it is the only way
 * to get T4's column/row sizing. `cellFormula` and `sheetStubs` are dropped:
 * nothing in Atlas reads `cell.f` or consumes stub (`t:'z'`) cells.
 */
import * as XLSX from 'xlsx'
import { formatCellText } from './xlsxCellFormat'

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
}

export type ParsedSheet = {
  readonly name: string
  readonly hidden: boolean
  readonly grid: SheetGrid
}

const EMPTY_GRID: SheetGrid = {
  rows: [],
  colCount: 0,
  merges: [],
  colWidthsPx: [],
  rowHeightsPx: [],
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
  for (let r = 0; r < rowCount; r++) {
    const row: string[] = new Array<string>(colCount).fill('')
    for (let c = 0; c < colCount; c++) {
      const addr = XLSX.utils.encode_cell({ r: range.s.r + r, c: range.s.c + c })
      const cell = ws[addr] as XLSX.CellObject | undefined
      row[c] = formatCellText(cell)
    }
    rows.push(row)
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

  return { rows, colCount, merges, colWidthsPx, rowHeightsPx }
}

/** Sheet visibility state (0=visible, 1=hidden, 2=very hidden) → `hidden` (T4/DAT-11). */
function isSheetHidden(workbook: XLSX.WorkBook, index: number): boolean {
  const hidden = workbook.Workbook?.Sheets?.[index]?.Hidden
  return hidden === 1 || hidden === 2
}

/**
 * Parses an xlsx/ods/xls workbook buffer into one `ParsedSheet` per sheet,
 * in file order. Safe to call from a Worker (T2) or the main thread.
 */
export function parseWorkbookBuffer(buffer: ArrayBuffer): ParsedSheet[] {
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
