/**
 * Shared "rows → print-ready `<table>` HTML" builder (X1). Used by both
 * `spreadsheetPdf.ts` (XLSX/ODS, parsed via the shared `spreadsheetGrid.ts`)
 * and `csvPdf.ts` (CSV/TSV, parsed via the shared `csvParse.ts`) so a
 * spreadsheet-shaped export always looks the same regardless of source
 * format.
 *
 * Pagination: an HTML `<table>`'s `<thead>` repeats on every printed page
 * automatically (standard, widely-supported browser/Chromium print
 * behavior) — this is what "header row repeated" actually means here; there
 * is no special-case pagination logic needed for row count, only for merged
 * cells (via `rowspan`/`colspan`, skipping cells a merge already covers).
 *
 * Known limitation (documented for the merger): a sheet with more columns
 * than fit across the printed page width still simply clips at the page
 * edge — HTML/CSS tables have no native "repeat these columns across
 * additional page-strips" pagination the way real spreadsheet apps offer.
 * Landscape orientation (see each caller's `@page` rule) mitigates this for
 * the common case but does not eliminate it for very wide sheets.
 */
import { escapeHtml } from './printDocument';

export interface TabularMerge {
  readonly r0: number;
  readonly c0: number;
  readonly r1: number;
  readonly c1: number;
}

export interface TabularSheet {
  readonly name: string;
  readonly rows: ReadonlyArray<ReadonlyArray<string>>;
  readonly merges?: ReadonlyArray<TabularMerge>;
}

function buildSpanIndex(merges: ReadonlyArray<TabularMerge>): {
  readonly spanByTopLeft: ReadonlyMap<string, { rowSpan: number; colSpan: number }>;
  readonly covered: ReadonlySet<string>;
} {
  const spanByTopLeft = new Map<string, { rowSpan: number; colSpan: number }>();
  const covered = new Set<string>();

  for (const merge of merges) {
    spanByTopLeft.set(`${merge.r0},${merge.c0}`, {
      rowSpan: merge.r1 - merge.r0 + 1,
      colSpan: merge.c1 - merge.c0 + 1,
    });
    for (let r = merge.r0; r <= merge.r1; r++) {
      for (let c = merge.c0; c <= merge.c1; c++) {
        if (r === merge.r0 && c === merge.c0) continue;
        covered.add(`${r},${c}`);
      }
    }
  }

  return { spanByTopLeft, covered };
}

function renderRow(
  row: ReadonlyArray<string>,
  rowIndex: number,
  tag: 'th' | 'td',
  spanIndex: ReturnType<typeof buildSpanIndex>,
): string {
  const cells = row
    .map((value, colIndex) => {
      const key = `${rowIndex},${colIndex}`;
      if (spanIndex.covered.has(key)) {
        return '';
      }
      const span = spanIndex.spanByTopLeft.get(key);
      const attrs = span ? ` rowspan="${span.rowSpan}" colspan="${span.colSpan}"` : '';
      return `<${tag}${attrs}>${escapeHtml(value)}</${tag}>`;
    })
    .join('');
  return `<tr>${cells}</tr>`;
}

function buildSheetSection(sheet: TabularSheet): string {
  const heading = `<h2>${escapeHtml(sheet.name)}</h2>`;

  if (sheet.rows.length === 0) {
    return `<section class="export-sheet">${heading}<p class="export-sheet__empty">(empty sheet)</p></section>`;
  }

  const spanIndex = buildSpanIndex(sheet.merges ?? []);
  const [headerRow, ...bodyRows] = sheet.rows;
  const thead = headerRow ? `<thead>${renderRow(headerRow, 0, 'th', spanIndex)}</thead>` : '';
  const tbody = `<tbody>${bodyRows.map((row, i) => renderRow(row, i + 1, 'td', spanIndex)).join('')}</tbody>`;

  return `<section class="export-sheet">${heading}<table>${thead}${tbody}</table></section>`;
}

/** Builds the `<section>` markup for every sheet, ready to sanitize + embed. */
export function buildTabularBodyHtml(sheets: ReadonlyArray<TabularSheet>): string {
  return sheets.map(buildSheetSection).join('\n');
}

/** Shared table styling + one-sheet-per-page pagination. Landscape Letter by default — mitigates (does not eliminate) wide-sheet column clipping. */
export const TABULAR_PRINT_CSS = `
@page { size: 11in 8.5in; margin: 0.5in; }
* { box-sizing: border-box; }
body {
  margin: 0;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif;
  font-size: 9pt;
  color: #1f2328;
  background: #ffffff;
}
.export-sheet { page-break-before: always; break-before: page; }
.export-sheet:first-child { page-break-before: avoid; break-before: avoid; }
.export-sheet h2 { font-size: 13pt; margin: 0 0 10px; }
.export-sheet__empty { opacity: 0.6; font-style: italic; }
table { border-collapse: collapse; width: 100%; table-layout: auto; }
th, td {
  border: 1px solid #999999;
  padding: 3px 6px;
  text-align: left;
  vertical-align: top;
  word-break: break-word;
}
thead th {
  background: #eeeeee;
  font-weight: 600;
}
tr { page-break-inside: avoid; break-inside: avoid; }
`;
