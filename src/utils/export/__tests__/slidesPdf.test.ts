/**
 * slidesPdf.ts — X1/SLD-01/UX-02. Uses the same in-memory PPTX fixture
 * builder (`buildPptxFixtureZip`) `PptxViewer.test.tsx` uses, so this proves
 * the export produces one PDF page per VISIBLE slide from the exact same
 * parser/model the live viewer renders — not a screenshot of whatever
 * happened to be on screen (the old, `SLD-01`-broken behavior).
 */
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { buildPptxFixtureZip } from '../../../viewers/slides/pptx/__tests__/pptxFixture';
import { parsePptxSlides } from '../../../viewers/slides/pptx/parser';
import type { CancelSignal, ZipArchive } from '../../../viewers/slides/shared/xmlUtils';
import { DocxParseError } from '../../../docx/parser/unzip';
import { loadOfficePackage } from '../../../office/officePackage';
import { exportSlidesPdf } from '../slidesPdf';

// Review fix — exportSlidesPdf now reads the archive through
// `loadOfficePackage` (the same size-budgeted `unzipDocx` the live
// viewer/editor use) instead of a bare `JSZip.loadAsync`. Wrapping the real
// implementation (rather than replacing it outright) keeps every other test
// in this file exercising the genuine read path; only the dedicated test
// below swaps in a rejection to prove the guard's failure propagates.
vi.mock('../../../office/officePackage', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../../../office/officePackage')>();
  return { ...actual, loadOfficePackage: vi.fn(actual.loadOfficePackage) };
});

async function buildPptxBuffer(): Promise<ArrayBuffer> {
  const zip = buildPptxFixtureZip();
  return zip.generateAsync({ type: 'arraybuffer' });
}

async function countVisibleSlides(): Promise<number> {
  const zip = buildPptxFixtureZip();
  const signal: CancelSignal = { cancelled: false };
  const slides = await parsePptxSlides(zip as unknown as ZipArchive, signal);
  return slides.filter(s => !s.hidden).length;
}

describe('exportSlidesPdf', () => {
  let printToPdfMock: ReturnType<typeof vi.fn>;
  let saveBinaryFileMock: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    printToPdfMock = vi.fn().mockResolvedValue({ ok: true, bytes: new Uint8Array([1, 2, 3]) });
    saveBinaryFileMock = vi.fn().mockResolvedValue({ saved: true });
    window.electronAPI = {
      printToPdf: printToPdfMock,
      saveBinaryFile: saveBinaryFileMock,
    } as unknown as typeof window.electronAPI;
  });

  it('renders one .export-slide page per VISIBLE slide (hidden slides excluded, S14)', async () => {
    const buffer = await buildPptxBuffer();
    const expectedCount = await countVisibleSlides();
    expect(expectedCount).toBeGreaterThan(0);

    await exportSlidesPdf(buffer, 'deck.pptx', 'pptx');

    expect(printToPdfMock).toHaveBeenCalledTimes(1);
    const html = printToPdfMock.mock.calls[0]![0] as string;
    const slideCount = (html.match(/class="export-slide"/g) ?? []).length;
    expect(slideCount).toBe(expectedCount);
    // Slide 2 is hidden in the fixture and must never appear.
    expect(html).not.toContain('Hidden slide');
  });

  it('sizes the @page rule to the deck\'s own aspect ratio (12192000x6858000 EMU = 1280x720px = 13.33x7.5in)', async () => {
    const buffer = await buildPptxBuffer();
    await exportSlidesPdf(buffer, 'deck.pptx', 'pptx');

    const html = printToPdfMock.mock.calls[0]![0] as string;
    expect(html).toMatch(/@page\s*\{\s*size:\s*13\.3333\d*in\s+7\.5in/);
  });

  it('saves the returned bytes as a .pdf through the native save dialog', async () => {
    const buffer = await buildPptxBuffer();
    await exportSlidesPdf(buffer, 'deck.pptx', 'pptx');

    expect(saveBinaryFileMock).toHaveBeenCalledWith(
      expect.objectContaining({
        suggestedName: 'deck.pdf',
        filters: [{ name: 'PDF Document', extensions: ['pdf'] }],
      }),
    );
  });

  it('rejects with a friendly error when no slides are found', async () => {
    const { default: JSZip } = await import('jszip');
    const emptyZip = new JSZip();
    emptyZip.file('ppt/presentation.xml', '<p:presentation xmlns:p="x"><p:sldIdLst/></p:presentation>');
    const buffer = await emptyZip.generateAsync({ type: 'arraybuffer' });

    await expect(exportSlidesPdf(buffer, 'empty.pptx', 'pptx')).rejects.toThrow('PDF export failed');
  });

  it('reads the archive through the size-budgeted office-package reader, and surfaces its zip-bomb guard as a friendly error', async () => {
    const buffer = await buildPptxBuffer();
    vi.mocked(loadOfficePackage).mockRejectedValueOnce(
      new DocxParseError(
        'Archive\'s combined uncompressed size is 999,999,999 bytes, over the 500 MiB total limit; ' +
          'refusing to extract (possible zip bomb).',
      ),
    );

    await expect(exportSlidesPdf(buffer, 'deck.pptx', 'pptx')).rejects.toThrow(/possible zip bomb/);
    expect(loadOfficePackage).toHaveBeenCalledWith(buffer);
    // The parser never even runs once the size guard rejects the archive.
    expect(printToPdfMock).not.toHaveBeenCalled();
  });
});
