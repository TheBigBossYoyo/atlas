/**
 * Loads an xlsx/ods workbook buffer into `ParsedSheet[]`, off the main
 * thread via a Worker once the file is large enough to risk a noticeable
 * freeze (T2/DAT-07).
 *
 * Small files are parsed synchronously and inline — spinning up a Worker has
 * its own (small but real) latency, and jsdom / non-browser test
 * environments have no `Worker` global at all, so the synchronous path
 * doubles as the automatic fallback there. Either way the caller only ever
 * sees `{status, sheets, error}`; which path ran is an implementation detail.
 */
import { useEffect, useRef, useState } from 'react'

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
import { XLSX_WORKER_BYTE_THRESHOLD } from './sizeThresholds'
import type { SpreadsheetWorkerRequest, SpreadsheetWorkerResponse } from './spreadsheetWorker.worker'
import { readFrozenPanes } from '../spreadsheet/spreadsheetPanes'
import { readSheetPartPaths, readSheetTables } from '../spreadsheet/spreadsheetTables'

export type SpreadsheetWorkbookState =
  | { readonly status: 'loading' }
  | { readonly status: 'ready'; readonly sheets: ParsedSheet[] }
  | { readonly status: 'error'; readonly error: string }

function shouldUseWorker(byteLength: number): boolean {
  return byteLength >= XLSX_WORKER_BYTE_THRESHOLD && typeof Worker !== 'undefined'
}

export function useSpreadsheetWorkbook(buffer: ArrayBuffer | null): SpreadsheetWorkbookState {
  const [state, setState] = useState<SpreadsheetWorkbookState>({ status: 'loading' })
  const workerRef = useRef<Worker | null>(null)

  useEffect(() => {
    workerRef.current?.terminate()
    workerRef.current = null

    let cancelled = false
    let worker: Worker | undefined

    // The WORKER branch below runs synchronously up to the point `worker` is
    // assigned, so it is already set by the time this IIFE yields and the
    // effect's own cleanup runs. (The small-file branch does await — frozen
    // panes, tables, part paths and, since SHEETFMT-1, cell styles are all
    // read from the zip — but it returns before ever reaching the worker.)
    // Nested in a callback, rather than at the effect's top level, so it
    // synchronizes with the external parse/Worker call instead of reading as
    // derivable state (mirrors DocxViewer's load effect).
    void (async () => {
      if (!buffer) {
        setState({ status: 'loading' })
        return
      }

      setState({ status: 'loading' })

      if (!shouldUseWorker(buffer.byteLength)) {
        try {
          const sheets = parseWorkbookBuffer(buffer)
          // Small-file-only (guarded by the `shouldUseWorker` branch above):
          // safe to also read frozen-pane metadata inline here (see
          // spreadsheetPanes.ts's header on why this step is skipped on the
          // main thread for large files instead).
          const [paneMap, tableMap, partPaths] = await Promise.all([
            readFrozenPanes(buffer),
            readSheetTables(buffer),
            readSheetPartPaths(buffer),
          ])
          const withSources = attachSheetSources(
            attachTables(attachFrozenPanes(sheets, paneMap), tableMap),
            partPaths,
          )
      // SHEETFMT-1 — cell formatting is read last, because it needs each
      // sheet's `sourcePath` (from `attachSheetSources` above) to know which
      // worksheet part to scan. It returns `null` for any format with no
      // `styles.xml`, which is the "render unstyled" path Atlas had before.
          const styles = await readWorkbookCellStyles(buffer, cellStyleRequests(withSources))
          if (!cancelled) {
            setState({ status: 'ready', sheets: attachCellStyles(withSources, styles?.styles ?? null) })
          }
        } catch (err) {
          if (!cancelled) setState({ status: 'error', error: err instanceof Error ? err.message : String(err) })
        }
        return
      }

      worker = new Worker(new URL('./spreadsheetWorker.worker.ts', import.meta.url), { type: 'module' })
      workerRef.current = worker

      worker.onmessage = (event: MessageEvent<SpreadsheetWorkerResponse>) => {
        if (cancelled) return
        const response = event.data
        if (response.ok) {
          setState({ status: 'ready', sheets: response.sheets })
        } else {
          setState({ status: 'error', error: response.error })
        }
      }
      worker.onerror = (event: ErrorEvent) => {
        if (cancelled) return
        setState({ status: 'error', error: event.message || 'Failed to parse the workbook.' })
      }

      // Structured-cloned (not transferred): `buffer` is `file.content`,
      // owned by the app shell's LoadedFile, not by this hook — transferring
      // it would detach it globally, breaking anything else that later reads
      // the same file object (e.g. re-opening this viewer, or a binary-export
      // path). The clone is a memcpy, not a CPU-bound parse, so it does not
      // reintroduce the long-main-thread-task problem this hook exists to avoid.
      const request: SpreadsheetWorkerRequest = { buffer }
      worker.postMessage(request)
    })()

    return () => {
      cancelled = true
      worker?.terminate()
      workerRef.current = null
    }
  }, [buffer])

  return state
}
