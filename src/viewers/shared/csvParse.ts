/**
 * Promise-wrapped Papa Parse, switching on Papa's own `worker:true` above a
 * size threshold so a large CSV/TSV doesn't block the main thread while
 * parsing (T2/DAT-07).
 *
 * Papa Parse's worker mode re-serializes its own module source into a
 * `blob:` Worker (see `papaparse.js`'s `getWorkerBlob`) rather than fetching
 * an external script file, so it works unmodified under a bundled build and
 * Atlas's `worker-src 'self' blob:` CSP (P1.2) — no separate worker file
 * needed here, unlike the xlsx path in `useSpreadsheetWorkbook.ts`. Without
 * `worker:true` (or when `Papa.WORKERS_SUPPORTED` is false, e.g. in a
 * non-browser test runner) Papa still calls `complete` — synchronously, in
 * the same tick — so callers can always just `await` this.
 */
import Papa from 'papaparse'
import type { ParseResult } from 'papaparse'

import { CSV_WORKER_CHAR_THRESHOLD } from './sizeThresholds'

export type CsvParseOptions = {
  readonly delimiter?: string
}

export function parseCsv(content: string, options: CsvParseOptions = {}): Promise<ParseResult<string[]>> {
  return new Promise((resolve) => {
    const complete = (result: ParseResult<string[]>): void => resolve(result)

    // Papa's TS overloads key off a literal `worker: true`/absent-`worker`
    // rather than a plain `boolean`, so the threshold check has to pick an
    // overload rather than being passed as a field.
    if (content.length >= CSV_WORKER_CHAR_THRESHOLD) {
      Papa.parse<string[]>(content, {
        worker: true,
        header: false,
        skipEmptyLines: true,
        delimiter: options.delimiter,
        complete,
      })
    } else {
      Papa.parse<string[]>(content, {
        header: false,
        skipEmptyLines: true,
        delimiter: options.delimiter,
        complete,
      })
    }
  })
}
