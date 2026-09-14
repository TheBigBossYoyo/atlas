/**
 * Named size thresholds above which a viewer switches to an
 * asynchronous/virtualized code path instead of parsing or rendering
 * synchronously on the main thread (T2/DAT-07/DAT-13).
 */

/** Above this many bytes, SpreadsheetViewer parses xlsx/ods off the main thread via a Worker. */
export const XLSX_WORKER_BYTE_THRESHOLD = 1_000_000

/** Above this many characters, CsvViewer enables Papa Parse's own `worker:true`. */
export const CSV_WORKER_CHAR_THRESHOLD = 1_000_000

/** Above this many lines, CodeViewer skips shiki's synchronous tokenizer and falls back to virtualized plain text. */
export const CODE_VIRTUALIZE_LINE_THRESHOLD = 5_000

/** Above this many bytes, CodeViewer falls back to virtualized plain text even if the line count is low (e.g. minified files). */
export const CODE_VIRTUALIZE_BYTE_THRESHOLD = 500_000
