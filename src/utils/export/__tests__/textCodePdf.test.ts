import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { exportTextPdf, exportTextHtml } from '../textCodePdf';

afterEach(() => {
  delete (window as { electronAPI?: unknown }).electronAPI;
  vi.restoreAllMocks();
});

describe('exportTextPdf', () => {
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

  it('renders the full raw content preformatted, regardless of file size (TextViewer/CodeViewer virtualize the live DOM)', async () => {
    const lines = Array.from({ length: 5000 }, (_, i) => `line ${i}`).join('\n');

    await exportTextPdf(lines, 'huge.txt');

    const html = printToPdfMock.mock.calls[0]![0] as string;
    expect(html).toContain('<pre>');
    expect(html).toContain('line 0');
    expect(html).toContain('line 4999');
  });

  it('HTML-escapes content and strips executable markup (a code file can contain literal "<script>" text)', async () => {
    await exportTextPdf('<script>alert(1)</script>', 'snippet.js');

    const html = printToPdfMock.mock.calls[0]![0] as string;
    expect(html).not.toContain('<script>alert(1)</script>');
    expect(html).toContain('&lt;script&gt;');
  });

  it('saves the returned bytes as a .pdf through the native save dialog', async () => {
    await exportTextPdf('hello', 'notes.txt');
    expect(saveBinaryFileMock).toHaveBeenCalledWith(expect.objectContaining({ suggestedName: 'notes.pdf' }));
  });
});

describe('exportTextHtml', () => {
  it('saves a self-contained, sanitized .html file with the raw content preformatted', async () => {
    const saveFileMock = vi.fn().mockResolvedValue({ saved: true });
    window.electronAPI = { saveFile: saveFileMock } as unknown as typeof window.electronAPI;

    await exportTextHtml('function greet() {}\n<script>alert(1)</script>', 'code.ts');

    expect(saveFileMock).toHaveBeenCalledTimes(1);
    const call = saveFileMock.mock.calls[0]![0] as { content: string; suggestedName: string };
    expect(call.suggestedName).toBe('code.html');
    expect(call.content).toContain('<pre>');
    expect(call.content).toContain('function greet');
    expect(call.content).not.toContain('<script>alert(1)</script>');
    expect(call.content).toContain('Content-Security-Policy');
  });
});
