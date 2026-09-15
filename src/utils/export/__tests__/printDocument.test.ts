import { afterEach, describe, expect, it, vi } from 'vitest';
import { buildPrintDocument, escapeHtml, renderHtmlToPdfFile } from '../printDocument';

describe('escapeHtml', () => {
  it('escapes the five HTML-significant characters', () => {
    expect(escapeHtml(`<a href="x">'&'</a>`)).toBe('&lt;a href=&quot;x&quot;&gt;&#39;&amp;&#39;&lt;/a&gt;');
  });
});

describe('buildPrintDocument', () => {
  it('embeds a restrictive CSP meta tag, the escaped title, the given CSS, and the given body verbatim', () => {
    const html = buildPrintDocument({ title: 'My <Doc>', css: 'body{color:red}', bodyHtml: '<p>hi</p>' });
    expect(html).toContain('<meta http-equiv="Content-Security-Policy"');
    expect(html).toContain("default-src 'none'");
    expect(html).toContain('<title>My &lt;Doc&gt;</title>');
    expect(html).toContain('<style>body{color:red}</style>');
    expect(html).toContain('<body><p>hi</p></body>');
  });
});

describe('renderHtmlToPdfFile', () => {
  afterEach(() => {
    delete (window as { electronAPI?: unknown }).electronAPI;
  });

  it('sanitizes the body, sends the built document to printToPdf, and saves the bytes as a .pdf', async () => {
    const printToPdfMock = vi.fn().mockResolvedValue({ ok: true, bytes: new Uint8Array([9]) });
    const saveBinaryFileMock = vi.fn().mockResolvedValue({ saved: true });
    window.electronAPI = {
      printToPdf: printToPdfMock,
      saveBinaryFile: saveBinaryFileMock,
    } as unknown as typeof window.electronAPI;

    await renderHtmlToPdfFile('<p>hi</p><script>pwn()</script>', 'body{}', 'My Doc', 'my-doc');

    const sentHtml = printToPdfMock.mock.calls[0]![0] as string;
    expect(sentHtml).toContain('<p>hi</p>');
    expect(sentHtml).not.toContain('<script');

    expect(saveBinaryFileMock).toHaveBeenCalledWith(
      expect.objectContaining({
        content: expect.any(Uint8Array),
        suggestedName: 'my-doc.pdf',
        filters: [{ name: 'PDF Document', extensions: ['pdf'] }],
      }),
    );
  });
});
