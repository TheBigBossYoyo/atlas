/**
 * XLSX/ODS → PDF export (X1 / RUN-07 / SHELL-13). Builds a table PDF from
 * the PARSED workbook data — not the live `SpreadsheetViewer` DOM, which
 * renders through `@glideapps/glide-data-grid` (a `<canvas>`-backed
 * virtualized grid: only the currently-scrolled-into-view cells ever exist
 * as real DOM/canvas content). Reuses `parseWorkbookBuffer` — the exact
 * same parser `useSpreadsheetWorkbook` uses — so every row, every sheet
 * (including ones scrolled out of view) is present in the export, and
 * export never drifts from what the viewer would show for the same cell.
 */
import { parseWorkbookBuffer, type ParsedSheet } from '../../viewers/shared/spreadsheetGrid';
import { buildTabularBodyHtml, TABULAR_PRINT_CSS, type TabularSheet } from './tabularHtml';
import { renderHtmlToPdfFile } from './printDocument';
import { sanitizeFileName, saveBinaryOutput, saveTextOutput, type SaveFilter } from './download';
import { toFriendlyError } from '../friendlyLibraryError';

function toTabularSheets(sheets: ReadonlyArray<ParsedSheet>): TabularSheet[] {
  // S14-equivalent for spreadsheets: a hidden sheet isn't part of the
  // document the user actually sees, so (matching PptxViewer/OdpViewer's
  // own "hidden slides are excluded" convention) it's excluded here too.
  return sheets
    .filter(sheet => !sheet.hidden)
    .map(sheet => ({ name: sheet.name, rows: sheet.grid.rows, merges: sheet.grid.merges }));
}

/** Parses `buffer` (xlsx/ods) and exports every visible sheet as one PDF, one page-group per sheet, all rows, header row repeated per page. */
export async function exportSpreadsheetPdf(buffer: ArrayBuffer, fileName: string): Promise<void> {
  try {
    const sheets = toTabularSheets(parseWorkbookBuffer(buffer));
    if (sheets.length === 0) {
      throw new Error('this workbook has no visible sheets to export');
    }

    const bodyHtml = buildTabularBodyHtml(sheets);
    await renderHtmlToPdfFile(bodyHtml, TABULAR_PRINT_CSS, fileName, fileName);
  } catch (err) {
    console.error('[export] exportSpreadsheetPdf failed:', err);
    throw toFriendlyError(err, 'PDF export failed');
  }
}

/** "Save a copy" for XLSX/ODS — copies the original bytes verbatim through the native save dialog (mirrors the PDF passthrough; UX-11 requires the dialog, not a raw browser download). */
export async function exportWorkbookCopy(
  buffer: ArrayBuffer,
  fileName: string,
  format: 'xlsx' | 'ods',
): Promise<void> {
  try {
    const mime =
      format === 'xlsx'
        ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        : 'application/vnd.oasis.opendocument.spreadsheet';
    const filters: readonly SaveFilter[] = [
      { name: format === 'xlsx' ? 'Excel Workbook' : 'OpenDocument Spreadsheet', extensions: [format] },
    ];
    await saveBinaryOutput(buffer, sanitizeFileName(fileName, format), mime, filters);
  } catch (err) {
    console.error('[export] exportWorkbookCopy failed:', err);
    throw toFriendlyError(err, `${format.toUpperCase()} export failed`);
  }
}

function quoteCsvField(field: string): string {
  return /[",\r\n]/.test(field) ? `"${field.replace(/"/g, '""')}"` : field;
}

function rowsToCsv(rows: ReadonlyArray<ReadonlyArray<string>>): string {
  return rows.map(row => row.map(quoteCsvField).join(',')).join('\r\n');
}

/**
 * UX-12 — exports every VISIBLE sheet as its own `.csv` file. A workbook has
 * no native single-file multi-sheet CSV representation (CSV is inherently
 * one table), so a multi-sheet workbook opens one native save dialog per
 * sheet in turn, each suggested as `"<name> - <sheet>.csv"`; a single-sheet
 * workbook is exactly one dialog. Stops (without erroring) if the user
 * cancels a dialog — matches the existing "cancel is not a failure"
 * convention (`saveTextOutput`/`saveBinaryOutput`).
 */
export async function exportSpreadsheetCsvPerSheet(buffer: ArrayBuffer, fileName: string): Promise<void> {
  try {
    const sheets = parseWorkbookBuffer(buffer).filter(sheet => !sheet.hidden);
    if (sheets.length === 0) {
      throw new Error('this workbook has no visible sheets to export');
    }

    const baseName = fileName.replace(/\.[^./\\]+$/, '');
    const single = sheets.length === 1;

    const csvFilters: readonly SaveFilter[] = [{ name: 'CSV', extensions: ['csv'] }];
    for (const sheet of sheets) {
      const csv = rowsToCsv(sheet.grid.rows);
      const suggested = single ? `${baseName}.csv` : `${baseName} - ${sheet.name}.csv`;
      // Deliberately sequential (not Promise.all): each sheet needs its own
      // native save dialog, shown one at a time.
      await saveTextOutput(csv, suggested, 'text/csv;charset=utf-8', csvFilters);
    }
  } catch (err) {
    console.error('[export] exportSpreadsheetCsvPerSheet failed:', err);
    throw toFriendlyError(err, 'CSV export failed');
  }
}
