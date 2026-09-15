/**
 * CSV/TSV → PDF export (X1 / RUN-07 / SHELL-13). Same rationale as
 * `spreadsheetPdf.ts`: `CsvViewer` renders through the same virtualized
 * `@glideapps/glide-data-grid` as `SpreadsheetViewer`, so exporting from the
 * live DOM would only ever capture the rows currently scrolled into view.
 * Parses the file's own raw text (already in hand — csv/tsv are `kind:
 * 'text'` `LoadedFile`s) via the exact same `parseCsv` `CsvViewer` uses, so
 * every row is present and the export can never disagree with the viewer
 * about how a row was parsed.
 */
import { parseCsv } from '../../viewers/shared/csvParse';
import { buildTabularBodyHtml, TABULAR_PRINT_CSS } from './tabularHtml';
import { renderHtmlToPdfFile } from './printDocument';
import { toFriendlyError } from '../friendlyLibraryError';

/** Parses `content` (csv or tab-delimited tsv) and exports every row as one paginated table PDF, header row repeated per page. */
export async function exportDelimitedTablePdf(
  content: string,
  fileName: string,
  sourceFormat: 'csv' | 'tsv',
): Promise<void> {
  try {
    const delimiter = sourceFormat === 'tsv' ? '\t' : undefined;
    const result = await parseCsv(content, { delimiter });
    const rows = result.data;

    if (rows.length === 0) {
      throw new Error('this file has no rows to export');
    }

    const bodyHtml = buildTabularBodyHtml([{ name: fileName, rows }]);
    await renderHtmlToPdfFile(bodyHtml, TABULAR_PRINT_CSS, fileName, fileName);
  } catch (err) {
    console.error('[export] exportDelimitedTablePdf failed:', err);
    throw toFriendlyError(err, 'PDF export failed');
  }
}
