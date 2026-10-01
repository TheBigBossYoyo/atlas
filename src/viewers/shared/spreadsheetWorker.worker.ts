/**
 * Worker entry for off-main-thread xlsx/ods parsing (T2/DAT-07).
 *
 * `SpreadsheetViewer` posts the file's raw ArrayBuffer here (as a
 * transferable) once its size crosses `XLSX_WORKER_BYTE_THRESHOLD`; parsing
 * a 100k-row workbook happens entirely off the main thread, so the app stays
 * responsive while it runs. All the actual parsing logic lives in
 * `spreadsheetGrid.ts`, which is plain, side-effect-free TS — this file is
 * only the postMessage plumbing, so the parsing logic itself stays testable
 * (and used) directly from the main-thread synchronous fallback path too.
 *
 * Also runs the frozen-pane reader (T4/DAT-10 remainder, `spreadsheetPanes.ts`)
 * on the same buffer, for the same off-main-thread reason: unzipping and
 * DOMParser-walking a large sheet's raw XML is exactly the kind of
 * main-thread-blocking work T2 exists to move off the renderer's thread.
 * Chromium exposes `DOMParser` on the Worker global scope, so this works
 * unmodified here.
 */
import {
  attachCellStyles,
  attachFrozenPanes,
  attachSheetSources,
  attachTables,
  cellStyleRequests,
  parseWorkbookBuffer,
  type ParsedSheet,
} from './spreadsheetGrid'
import { readWorkbookCellStyles } from '../spreadsheet/xlsxCellStyles'
import { readFrozenPanes } from '../spreadsheet/spreadsheetPanes'
import { readSheetPartPaths, readSheetTables } from '../spreadsheet/spreadsheetTables'

export type SpreadsheetWorkerRequest = {
  readonly buffer: ArrayBuffer
}

export type SpreadsheetWorkerResponse =
  | { readonly ok: true; readonly sheets: ParsedSheet[] }
  | { readonly ok: false; readonly error: string }

self.onmessage = async (event: MessageEvent<SpreadsheetWorkerRequest>) => {
  try {
    const sheets = parseWorkbookBuffer(event.data.buffer)
    const [paneMap, tableMap, partPaths] = await Promise.all([
      readFrozenPanes(event.data.buffer),
      readSheetTables(event.data.buffer),
      readSheetPartPaths(event.data.buffer),
    ])
    const withSources = attachSheetSources(attachTables(attachFrozenPanes(sheets, paneMap), tableMap), partPaths)
    // SHEETFMT-1 — cell formatting needs each sheet's `sourcePath` to know
    // which worksheet part to scan, so it is read after `attachSheetSources`.
    // This is the large-file path, which is exactly where doing it off the main
    // thread matters: a 100k-row sheet's style scan is the same size as its
    // value parse.
    const styles = await readWorkbookCellStyles(event.data.buffer, cellStyleRequests(withSources))
    const response: SpreadsheetWorkerResponse = {
      ok: true,
      sheets: attachCellStyles(withSources, styles?.styles ?? null),
    }
    self.postMessage(response)
  } catch (err) {
    const response: SpreadsheetWorkerResponse = {
      ok: false,
      error: err instanceof Error ? err.message : String(err),
    }
    self.postMessage(response)
  }
}
