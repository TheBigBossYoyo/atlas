/**
 * TSV clipboard serialization for copy/paste ranges (wave 3 editing).
 *
 * Real spreadsheet applications put a tab-separated (rows by `\n`, cells by
 * `\t`) plain-text representation on the clipboard alongside their own
 * richer native format, specifically so pasting into a *different*
 * application (or, here, back into Atlas) still works. This module is
 * deliberately the plain-TSV subset only: a cell value containing a literal
 * tab or newline is not quote-escaped/unescaped (unlike full CSV/RFC 4180
 * quoting) — a real spreadsheet's own TSV export has the same limitation for
 * the same reason (TSV has no standardized quoting convention the way CSV
 * does), so this matches real-world behavior rather than under-delivering
 * against it.
 */

/** Serializes a rectangular block of cell values as TSV text for the clipboard. */
export function rangeToTsv(rows: ReadonlyArray<ReadonlyArray<string>>): string {
  return rows.map((row) => row.join('\t')).join('\n')
}

/**
 * Parses clipboard text (TSV, or a single value with no tabs/newlines at
 * all) back into a rectangular block of rows. Trailing `\r` from
 * Windows-style line endings is stripped per line.
 */
export function tsvToRows(text: string): string[][] {
  const normalized = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n')
  // A trailing newline (common when copying from a real spreadsheet) should
  // not produce a spurious trailing empty row.
  const withoutTrailingNewline = normalized.endsWith('\n') ? normalized.slice(0, -1) : normalized
  if (withoutTrailingNewline === '') return [['']]
  return withoutTrailingNewline.split('\n').map((line) => line.split('\t'))
}
