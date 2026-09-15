import { afterEach, describe, expect, it, vi } from 'vitest';
import { printHtmlToPdf, PdfExportUnavailableError } from '../printToPdf';

afterEach(() => {
  delete (window as { electronAPI?: unknown }).electronAPI;
});

describe('printHtmlToPdf', () => {
  it('returns the bytes from a successful window.electronAPI.printToPdf call', async () => {
    const bytes = new Uint8Array([1, 2, 3]);
    window.electronAPI = {
      printToPdf: vi.fn().mockResolvedValue({ ok: true, bytes }),
    } as unknown as typeof window.electronAPI;

    await expect(printHtmlToPdf('<html></html>')).resolves.toBe(bytes);
  });

  it('throws the main-process-reported error when printToPdf reports failure', async () => {
    window.electronAPI = {
      printToPdf: vi.fn().mockResolvedValue({ ok: false, error: 'boom' }),
    } as unknown as typeof window.electronAPI;

    await expect(printHtmlToPdf('<html></html>')).rejects.toThrow('boom');
  });

  it('throws PdfExportUnavailableError when window.electronAPI.printToPdf does not exist', async () => {
    await expect(printHtmlToPdf('<html></html>')).rejects.toThrow(PdfExportUnavailableError);
    await expect(printHtmlToPdf('<html></html>')).rejects.toThrow('PDF export requires the desktop app.');
  });
});
