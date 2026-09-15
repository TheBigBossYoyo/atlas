/**
 * csvPdf.ts — X1/RUN-07. `CsvViewer` renders through the same virtualized
 * grid as `SpreadsheetViewer`, so this proves the export parses the file's
 * OWN raw text (every row) rather than reading anything from the viewer.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { exportDelimitedTablePdf } from '../csvPdf';

describe('exportDelimitedTablePdf', () => {
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

  it('renders every row of a large CSV file (RUN-07 — a 394,501-row live grid only ever painted the on-screen canvas region)', async () => {
    const header = 'name,age';
    const rows = Array.from({ length: 300 }, (_, i) => `person-${i},${i}`);
    const content = [header, ...rows].join('\n');

    await exportDelimitedTablePdf(content, 'people.csv', 'csv');

    const html = printToPdfMock.mock.calls[0]![0] as string;
    expect(html).toContain('person-0');
    expect(html).toContain('person-299');
    expect((html.match(/<tr>/g) ?? []).length).toBe(301);
  });

  it('converts tab-delimited tsv content the same way CsvViewer parses it', async () => {
    await exportDelimitedTablePdf('name\tage\nAda\t36', 'data.tsv', 'tsv');

    const html = printToPdfMock.mock.calls[0]![0] as string;
    expect(html).toContain('<th>name</th><th>age</th>');
    expect(html).toContain('<td>Ada</td><td>36</td>');
  });

  it('escapes HTML-significant characters in cell content', async () => {
    await exportDelimitedTablePdf('name\n<script>alert(1)</script>', 'data.csv', 'csv');

    const html = printToPdfMock.mock.calls[0]![0] as string;
    expect(html).not.toContain('<script>alert(1)</script>');
    expect(html).toContain('&lt;script&gt;');
  });

  it('saves the returned bytes as a .pdf through the native save dialog', async () => {
    await exportDelimitedTablePdf('a,b\n1,2', 'report.csv', 'csv');

    expect(saveBinaryFileMock).toHaveBeenCalledWith(
      expect.objectContaining({
        suggestedName: 'report.pdf',
        filters: [{ name: 'PDF Document', extensions: ['pdf'] }],
      }),
    );
  });

  it('rejects with a friendly error for an empty file', async () => {
    await expect(exportDelimitedTablePdf('', 'empty.csv', 'csv')).rejects.toThrow('PDF export failed');
  });
});
