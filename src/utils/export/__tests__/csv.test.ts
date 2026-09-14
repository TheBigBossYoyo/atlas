import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { exportCsv, tsvToCsv } from '../csv';

describe('tsvToCsv', () => {
  it('converts tab-delimited rows to comma-delimited rows', () => {
    expect(tsvToCsv('a\tb\tc\n1\t2\t3')).toBe('a,b,c\r\n1,2,3');
  });

  it('quotes a field that itself contains a comma', () => {
    expect(tsvToCsv('name\tnote\nAda\tHello, world')).toBe('name,note\r\nAda,"Hello, world"');
  });

  it('quotes a field containing a double quote, doubling the inner quote', () => {
    expect(tsvToCsv('5" screen')).toBe('"5"" screen"');
  });

  it('normalizes CRLF/CR/LF line endings to CRLF', () => {
    expect(tsvToCsv('a\tb\r\n1\t2\r3\t4')).toBe('a,b\r\n1,2\r\n3,4');
  });
});

describe('exportCsv', () => {
  let createObjectURLMock: ReturnType<typeof vi.spyOn>;
  let anchorClicks: Array<{ download: string; href: string }>;

  beforeEach(() => {
    anchorClicks = [];
    createObjectURLMock = vi.spyOn(URL, 'createObjectURL').mockReturnValue('blob:mock-url');
    vi.spyOn(URL, 'revokeObjectURL').mockImplementation(() => {});
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (this: HTMLAnchorElement) {
      anchorClicks.push({ download: this.download, href: this.href });
    });
  });

  afterEach(() => {
    vi.restoreAllMocks();
    delete (window as { electronAPI?: unknown }).electronAPI;
  });

  it('exports csv content as-is', async () => {
    await exportCsv('a,b\n1,2', 'sheet.csv', 'csv');

    expect(anchorClicks).toEqual([{ download: 'sheet.csv', href: 'blob:mock-url' }]);
    const blob = createObjectURLMock.mock.calls[0]![0] as Blob;
    expect(blob.type).toBe('text/csv;charset=utf-8');
    await expect(blob.text()).resolves.toBe('a,b\n1,2');
  });

  it('converts tsv content to csv before exporting', async () => {
    await exportCsv('a\tb\n1\t2', 'sheet.tsv', 'tsv');

    const blob = createObjectURLMock.mock.calls[0]![0] as Blob;
    await expect(blob.text()).resolves.toBe('a,b\r\n1,2');
  });

  it('sanitizes the filename to a .csv extension regardless of source format', async () => {
    await exportCsv('a,b', 'notes.tsv', 'tsv');
    expect(anchorClicks[0]!.download).toBe('notes.csv');
  });

  it('routes through the Electron save dialog with a .csv filter when window.electronAPI is present', async () => {
    const saveFileMock = vi.fn().mockResolvedValue({ saved: true });
    window.electronAPI = { saveFile: saveFileMock } as unknown as typeof window.electronAPI;

    await exportCsv('a,b\n1,2', 'sheet.csv', 'csv');

    expect(saveFileMock).toHaveBeenCalledWith({
      content: 'a,b\n1,2',
      suggestedName: 'sheet.csv',
      filters: [{ name: 'CSV', extensions: ['csv'] }],
    });
    expect(anchorClicks).toHaveLength(0);
  });

  it('wraps a save failure in a friendly, format-specific message', async () => {
    window.electronAPI = {
      saveFile: vi.fn().mockResolvedValue({ saved: false, error: 'disk full' }),
    } as unknown as typeof window.electronAPI;

    await expect(exportCsv('a,b', 'sheet.csv', 'csv')).rejects.toThrow('CSV export failed: disk full');
  });
});
