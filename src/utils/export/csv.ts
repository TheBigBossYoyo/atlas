/**
 * CSV export for csv/tsv (and, once a viewer registers `getExportableContent`
 * with `format: 'csv'`, spreadsheet) documents — UX-12.
 */
import { sanitizeFileName, saveTextOutput, type SaveFilter } from './download';
import { toFriendlyError } from '../friendlyLibraryError';

const CSV_MIME = 'text/csv;charset=utf-8';
const CSV_FILTERS: readonly SaveFilter[] = [{ name: 'CSV', extensions: ['csv'] }];

function quoteCsvField(field: string): string {
  return /[",\r\n]/.test(field) ? `"${field.replace(/"/g, '""')}"` : field;
}

/**
 * Converts tab-delimited text to comma-delimited CSV, quoting any field that
 * needs it (contains a comma, quote, or line break). TSV files essentially
 * never quote fields (the delimiter itself is unambiguous), so a plain
 * split-on-tab per line is a safe, simple conversion for the format.
 */
export function tsvToCsv(tsv: string): string {
  return tsv
    .split(/\r\n|\r|\n/)
    .map(line => line.split('\t').map(quoteCsvField).join(','))
    .join('\r\n');
}

/** Exports CSV or TSV text content as a `.csv` file, converting delimiters for TSV. */
export async function exportCsv(content: string, fileName: string, sourceFormat: 'csv' | 'tsv'): Promise<void> {
  try {
    const name = sanitizeFileName(fileName, 'csv');
    const csv = sourceFormat === 'tsv' ? tsvToCsv(content) : content;
    await saveTextOutput(csv, name, CSV_MIME, CSV_FILTERS);
  } catch (err) {
    console.error('[export] exportCsv failed:', err);
    throw toFriendlyError(err, 'CSV export failed');
  }
}
