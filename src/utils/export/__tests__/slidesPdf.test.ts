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
import { exportSlidesPdf } from '../slidesPdf';

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
});
