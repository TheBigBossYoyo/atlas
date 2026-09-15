/**
 * pdf.ts — the DOCX/RTF/ODT/PDF-passthrough exporters not already covered by
 * `src/utils/__tests__/export.markdown.test.tsx` (which owns
 * `exportMarkdownPdf`, since it shares that file's markdown-fixture
 * infrastructure).
 */
import { afterEach, describe, expect, it, vi } from 'vitest';
import { exportDocxPdf, exportRtfPdf, exportOdtPdf, exportPdfCopy } from '../pdf';

function mountViewerContent(innerHtml: string): void {
  const el = document.createElement('div');
  el.id = 'viewer-content';
  el.innerHTML = innerHtml;
  document.body.appendChild(el);
}

function mockElectron(): { printToPdfMock: ReturnType<typeof vi.fn>; saveBinaryFileMock: ReturnType<typeof vi.fn> } {
  const printToPdfMock = vi.fn().mockResolvedValue({ ok: true, bytes: new Uint8Array([1]) });
  const saveBinaryFileMock = vi.fn().mockResolvedValue({ saved: true });
  window.electronAPI = {
    printToPdf: printToPdfMock,
    saveBinaryFile: saveBinaryFileMock,
  } as unknown as typeof window.electronAPI;
  return { printToPdfMock, saveBinaryFileMock };
}

afterEach(() => {
  document.body.innerHTML = '';
  delete (window as { electronAPI?: unknown }).electronAPI;
  vi.restoreAllMocks();
});

describe('exportDocxPdf', () => {
  it('sizes the @page rule from the rendered page\'s own pixel size (zoom is always 1 — see PageView)', async () => {
    const { printToPdfMock } = mockElectron();
    mountViewerContent(
      '<div class="docx-page-stack">' +
        '<div class="docx-page" style="width: 816px; height: 1056px;">Page 1 content</div>' +
        '</div>',
    );

    await exportDocxPdf('viewer-content', 'letter.docx');

    const html = printToPdfMock.mock.calls[0]![0] as string;
    // 816px / 96 = 8.5in, 1056px / 96 = 11in — US Letter.
    expect(html).toMatch(/@page\s*\{\s*size:\s*8\.5in\s+11in;\s*margin:\s*0;\s*\}/);
    expect(html).toContain('Page 1 content');
  });

  it('uses one named @page per distinct section size for a multi-size document', async () => {
    const { printToPdfMock } = mockElectron();
    mountViewerContent(
      '<div class="docx-page-stack">' +
        '<div class="docx-page" style="width: 816px; height: 1056px;">Portrait</div>' +
        '<div class="docx-page" style="width: 1056px; height: 816px;">Landscape</div>' +
        '</div>',
    );

    await exportDocxPdf('viewer-content', 'mixed.docx');

    const html = printToPdfMock.mock.calls[0]![0] as string;
    expect(html).toMatch(/@page docx-size-0 \{ size: 8\.5in 11in; margin: 0; \}/);
    expect(html).toMatch(/@page docx-size-1 \{ size: 11in 8\.5in; margin: 0; \}/);
    expect(html).toContain('data-size-key="816x1056"');
    expect(html).toContain('data-size-key="1056x816"');
  });

  it('forces content-visibility back to visible (defense in depth against a partial export)', async () => {
    const { printToPdfMock } = mockElectron();
    mountViewerContent('<div class="docx-page-stack"><div class="docx-page" style="width: 816px; height: 1056px;">x</div></div>');

    await exportDocxPdf('viewer-content', 'doc.docx');

    const html = printToPdfMock.mock.calls[0]![0] as string;
    expect(html).toContain('content-visibility: visible !important');
  });

  it('rejects with a friendly error when the document has not finished loading', async () => {
    mockElectron();
    mountViewerContent('<div class="docx-viewer__loading">Loading…</div>');

    await expect(exportDocxPdf('viewer-content', 'doc.docx')).rejects.toThrow(
      'PDF export failed: the document has not finished loading yet',
    );
  });

  it('rejects with a friendly, id-specific error when the target element does not exist', async () => {
    mockElectron();
    await expect(exportDocxPdf('does-not-exist', 'doc.docx')).rejects.toThrow(
      'PDF export failed: element #does-not-exist not found',
    );
  });
});

describe('exportRtfPdf / exportOdtPdf', () => {
  it('exportRtfPdf sanitizes and sends the full converted document to printToPdf', async () => {
    const { printToPdfMock, saveBinaryFileMock } = mockElectron();
    mountViewerContent('');
    document.getElementById('viewer-content')!.outerHTML =
      '<div id="viewer-content" class="rtf-viewer"><p>Hello <script>alert(1)</script>world</p></div>';

    await exportRtfPdf('viewer-content', 'notes.rtf');

    const html = printToPdfMock.mock.calls[0]![0] as string;
    expect(html).toContain('Hello');
    expect(html).not.toContain('<script');
    expect(saveBinaryFileMock).toHaveBeenCalledWith(expect.objectContaining({ suggestedName: 'notes.pdf' }));
  });

  it('exportOdtPdf sends the full converted document to printToPdf', async () => {
    const { printToPdfMock } = mockElectron();
    mountViewerContent('<div class="odt-viewer__body">ODT body text</div>');

    await exportOdtPdf('viewer-content', 'doc.odt');

    const html = printToPdfMock.mock.calls[0]![0] as string;
    expect(html).toContain('ODT body text');
  });
});

describe('exportPdfCopy', () => {
  it('saves the original PDF bytes verbatim through the native save dialog (UX-11 — this used to bypass the dialog entirely)', async () => {
    const saveBinaryFileMock = vi.fn().mockResolvedValue({ saved: true });
    window.electronAPI = { saveBinaryFile: saveBinaryFileMock } as unknown as typeof window.electronAPI;
    const bytes = new Uint8Array([0x25, 0x50, 0x44, 0x46]).buffer;

    await exportPdfCopy(bytes, 'report.pdf');

    expect(saveBinaryFileMock).toHaveBeenCalledWith(
      expect.objectContaining({
        content: expect.any(Uint8Array),
        suggestedName: 'report.pdf',
        filters: [{ name: 'PDF Document', extensions: ['pdf'] }],
      }),
    );
  });
});
