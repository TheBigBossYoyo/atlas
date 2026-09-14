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
 */
import { parseWorkbookBuffer, type ParsedSheet } from './spreadsheetGrid'

export type SpreadsheetWorkerRequest = {
  readonly buffer: ArrayBuffer
}

export type SpreadsheetWorkerResponse =
  | { readonly ok: true; readonly sheets: ParsedSheet[] }
  | { readonly ok: false; readonly error: string }

self.onmessage = (event: MessageEvent<SpreadsheetWorkerRequest>) => {
  try {
    const sheets = parseWorkbookBuffer(event.data.buffer)
    const response: SpreadsheetWorkerResponse = { ok: true, sheets }
    self.postMessage(response)
  } catch (err) {
    const response: SpreadsheetWorkerResponse = {
      ok: false,
      error: err instanceof Error ? err.message : String(err),
    }
    self.postMessage(response)
  }
}
