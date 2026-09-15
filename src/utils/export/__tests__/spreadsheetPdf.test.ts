/**
 * spreadsheetPdf.ts — X1/RUN-07/SHELL-13. Builds workbooks programmatically
 * with `xlsx` (mirrors `SpreadsheetViewer.test.tsx`'s own `buildWorkbookFile`
 * helper) so the export can be checked against a workbook whose exact row
 * count, sheet visibility, and merges are known, without depending on a
 * binary fixture file's exact contents.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest';
import * as XLSX from 'xlsx';
import {
  exportSpreadsheetPdf,
  exportWorkbookCopy,
  exportSpreadsheetCsvPerSheet,
} from '../spreadsheetPdf';

function buildWorkbookBuffer(build: (wb: XLSX.WorkBook) => void): ArrayBuffer {
  const wb = XLSX.utils.book_new();
  build(wb);
  return XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
}

describe('exportSpreadsheetPdf', () => {
  let printToPdfMock: ReturnType<typeof vi.fn>;
  let saveBinaryFileMock: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    printToPdfMock = vi.fn().mockResolvedValue({ ok: true, bytes: new Uint8Array([1]) });
    saveBinaryFileMock = vi.fn().mockResolvedValue({ saved: true });
    window.electronAPI = {
      printToPdf: printToPdfMock,
      saveBinaryFile: saveBinaryFileMock,
    } as unknown as typeof window.electronAPI;
  });

  it('renders every row of a large sheet (not just what would fit on screen — RUN-07)', async () => {
    // 500 data rows + 1 header — a virtualized on-screen grid only ever
    // mounts the rows scrolled into view; export must have all of them.
    const rows = [['Name', 'Score'], ...Array.from({ length: 500 }, (_, i) => [`Row ${i}`, String(i)])];
    const buffer = buildWorkbookBuffer(wb => {
      const ws = XLSX.utils.aoa_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, 'Data');
    });

    await exportSpreadsheetPdf(buffer, 'big.xlsx');

    const html = printToPdfMock.mock.calls[0]![0] as string;
    expect(html).toContain('Row 0');
    expect(html).toContain('Row 499');
    expect((html.match(/<tr>/g) ?? []).length).toBe(501); // header + 500 data rows
  });

  it('excludes hidden sheets (mirrors S14\'s "hidden slides excluded" convention)', async () => {
    const buffer = buildWorkbookBuffer(wb => {
      const visible = XLSX.utils.aoa_to_sheet([['Visible Sheet Content']]);
      const hidden = XLSX.utils.aoa_to_sheet([['Hidden Sheet Content']]);
      XLSX.utils.book_append_sheet(wb, visible, 'Visible');
      XLSX.utils.book_append_sheet(wb, hidden, 'Hidden');
      wb.Workbook = { Sheets: [{ Hidden: 0 }, { Hidden: 1 }] };
    });

    await exportSpreadsheetPdf(buffer, 'wb.xlsx');

    const html = printToPdfMock.mock.calls[0]![0] as string;
    expect(html).toContain('Visible Sheet Content');
    expect(html).not.toContain('Hidden Sheet Content');
  });

  it('renders a merged cell as one spanning <th>/<td>, not repeated across every covered cell', async () => {
    const buffer = buildWorkbookBuffer(wb => {
      const ws = XLSX.utils.aoa_to_sheet([
        ['Title', '', 'Score'],
        ['A', 'B', '1'],
      ]);
      ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    });

    await exportSpreadsheetPdf(buffer, 'merged.xlsx');

    const html = printToPdfMock.mock.calls[0]![0] as string;
    expect(html).toContain('colspan="2"');
  });

  it('gives each sheet its own printed page group (page-break-before)', async () => {
    const buffer = buildWorkbookBuffer(wb => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['One']]), 'First');
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Two']]), 'Second');
    });

    await exportSpreadsheetPdf(buffer, 'multi.xlsx');

    const html = printToPdfMock.mock.calls[0]![0] as string;
    expect((html.match(/class="export-sheet"/g) ?? []).length).toBe(2);
    expect(html).toContain('page-break-before: always');
  });

  it('saves the returned bytes as a .pdf through the native save dialog', async () => {
    const buffer = buildWorkbookBuffer(wb => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['x']]), 'Sheet1');
    });

    await exportSpreadsheetPdf(buffer, 'report.xlsx');

    expect(saveBinaryFileMock).toHaveBeenCalledWith(
      expect.objectContaining({ suggestedName: 'report.pdf' }),
    );
  });

  it('rejects with a friendly error when the workbook has no visible sheets', async () => {
    const buffer = buildWorkbookBuffer(wb => {
      const ws = XLSX.utils.aoa_to_sheet([['x']]);
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
      wb.Workbook = { Sheets: [{ Hidden: 1 }] };
    });

    await expect(exportSpreadsheetPdf(buffer, 'allhidden.xlsx')).rejects.toThrow('PDF export failed');
  });
});

describe('exportWorkbookCopy', () => {
  it('saves the original bytes verbatim through the native save dialog with the right extension/MIME (UX-11 — this used to bypass the dialog for PDF; xlsx/ods never had a passthrough at all)', async () => {
    const saveBinaryFileMock = vi.fn().mockResolvedValue({ saved: true });
    window.electronAPI = { saveBinaryFile: saveBinaryFileMock } as unknown as typeof window.electronAPI;
    const bytes = new Uint8Array([9, 9, 9]).buffer;

    await exportWorkbookCopy(bytes, 'report.xlsx', 'xlsx');

    expect(saveBinaryFileMock).toHaveBeenCalledWith(
      expect.objectContaining({
        suggestedName: 'report.xlsx',
        filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }],
      }),
    );
  });
});

describe('exportSpreadsheetCsvPerSheet', () => {
  it('opens one save dialog per visible sheet, suggesting "<name> - <sheet>.csv"', async () => {
    const saveFileMock = vi.fn().mockResolvedValue({ saved: true });
    window.electronAPI = { saveFile: saveFileMock } as unknown as typeof window.electronAPI;
    const buffer = buildWorkbookBuffer(wb => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a', 'b'], ['1', '2']]), 'First');
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['c']]), 'Second');
    });

    await exportSpreadsheetCsvPerSheet(buffer, 'wb.xlsx');

    expect(saveFileMock).toHaveBeenCalledTimes(2);
    expect(saveFileMock.mock.calls[0]![0]).toMatchObject({ suggestedName: 'wb - First.csv', content: 'a,b\r\n1,2' });
    expect(saveFileMock.mock.calls[1]![0]).toMatchObject({ suggestedName: 'wb - Second.csv', content: 'c' });
  });

  it('uses the plain file name (no sheet suffix) for a single-sheet workbook', async () => {
    const saveFileMock = vi.fn().mockResolvedValue({ saved: true });
    window.electronAPI = { saveFile: saveFileMock } as unknown as typeof window.electronAPI;
    const buffer = buildWorkbookBuffer(wb => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a']]), 'Sheet1');
    });

    await exportSpreadsheetCsvPerSheet(buffer, 'single.xlsx');

    expect(saveFileMock).toHaveBeenCalledTimes(1);
    expect(saveFileMock.mock.calls[0]![0]).toMatchObject({ suggestedName: 'single.csv' });
  });
});
