/**
 * src/utils/export.ts — characterization tests (P0.4 / QA-07).
 *
 * Freezes today's exact export output for the same fixture set used by
 * MarkdownRenderer.characterization.test.tsx and useToc.test.ts. Per the
 * improvement plan's Section 7 guardrail and its X2 (export parser parity)
 * task note: these *export*-output snapshots are expected to change once
 * X2 lands (they already diverge from the live-render snapshots — that
 * divergence is a real, pre-existing bug this suite documents, not
 * something to fix here), while MarkdownRenderer's *live-render* snapshots
 * must stay byte-identical. Two known divergences this file locks in:
 *
 *   1. `exportHtml` renders markdown through `marked.parse` (not the live
 *      remark/rehype pipeline), so KaTeX math delimiters are NOT rendered —
 *      `marked` treats `$...$`/`$$...$$` as literal text and additionally
 *      strips escape backslashes inside them (`\int` -> `int`).
 *   2. `exportDocx`'s link handling (`inlineTokensToRuns`'s 'link' case)
 *      spreads an already-constructed `TextRun` instance into a new
 *      `TextRun({...})` to apply the Hyperlink style; `TextRun`'s own
 *      properties aren't the plain `{ text }` shape the constructor expects,
 *      so the link's visible text is silently dropped from every exported
 *      hyperlink run (the `w:hyperlink` and its style survive; the `w:t`
 *      inside it does not).
 *
 * exportDocx's hyperlink relationship ids (`r:id="..."`) are generated
 * fresh (effectively random) by the `docx` package on every `Packer.toBlob`
 * call, so they're sanitized to a fixed placeholder before snapshotting —
 * see `stabilizeDocxXml`. Rasterized PDF pixels are out of scope (jsdom has
 * no canvas backend); `exportPdf`'s tests assert its html2canvas/jsPDF
 * pipeline is driven correctly instead of inspecting pixels.
 */

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';

// ---------------------------------------------------------------------------
// html2canvas-pro / jsPDF mocks — jsdom has no canvas backend (getContext('2d')
// returns null, toDataURL() returns null), so exportPdf's real rasterization
// pipeline cannot run in this environment. Mocked deterministically per the
// task brief ("assert exportPdf calls its pipeline with the right element id").
// ---------------------------------------------------------------------------

const { html2canvasMock, jsPdfCtorMock, jsPdfAddImageMock, jsPdfAddPageMock, jsPdfOutputMock } = vi.hoisted(() => {
  const addImageMock = vi.fn();
  const addPageMock = vi.fn();
  const outputMock = vi.fn(() => new Blob(['%PDF-mock'], { type: 'application/pdf' }));
  return {
    // Typed with html2canvas's real (element, options) signature so the
    // assertions below can read `.mock.calls[0]` as [HTMLElement, object];
    // the values themselves are unused in the return, hence the `void`s.
    html2canvasMock: vi.fn(async (element: HTMLElement, options?: Record<string, unknown>) => {
      void element;
      void options;
      return {
        width: 800,
        height: 600,
        toDataURL: () => 'data:image/png;base64,MOCK',
      };
    }),
    // A plain function (not an arrow) — jsPDF is invoked with `new`, and an
    // arrow-returning mockImplementation fails with "is not a constructor".
    jsPdfCtorMock: vi.fn().mockImplementation(function jsPDF() {
      return {
        addPage: addPageMock,
        addImage: addImageMock,
        output: outputMock,
      };
    }),
    jsPdfAddImageMock: addImageMock,
    jsPdfAddPageMock: addPageMock,
    jsPdfOutputMock: outputMock,
  };
});

vi.mock('html2canvas-pro', () => ({ default: html2canvasMock }));
vi.mock('jspdf', () => ({ default: jsPdfCtorMock }));

import { exportMarkdown, exportHtml, exportPdf, exportDocx } from '../export';

import headingsFixture from '../../__tests__/fixtures/markdown/headings.md?raw';
import gfmTableFixture from '../../__tests__/fixtures/markdown/gfm-table.md?raw';
import taskListFixture from '../../__tests__/fixtures/markdown/task-list.md?raw';
import strikethroughAutolinksFixture from '../../__tests__/fixtures/markdown/strikethrough-autolinks.md?raw';
import codeFenceFixture from '../../__tests__/fixtures/markdown/code-fence.md?raw';
import mathFixture from '../../__tests__/fixtures/markdown/math.md?raw';
import mermaidFixture from '../../__tests__/fixtures/markdown/mermaid.md?raw';
import nestedListsFixture from '../../__tests__/fixtures/markdown/nested-lists.md?raw';
import blockquoteFixture from '../../__tests__/fixtures/markdown/blockquote.md?raw';
import rawHtmlFixture from '../../__tests__/fixtures/markdown/raw-html.md?raw';
import linksImagesFixture from '../../__tests__/fixtures/markdown/links-images.md?raw';
import kitchenSinkFixture from '../../__tests__/fixtures/markdown/kitchen-sink.md?raw';

const FIXTURES: ReadonlyArray<readonly [name: string, markdown: string]> = [
  ['headings', headingsFixture],
  ['gfm-table', gfmTableFixture],
  ['task-list', taskListFixture],
  ['strikethrough-autolinks', strikethroughAutolinksFixture],
  ['code-fence', codeFenceFixture],
  ['math', mathFixture],
  ['mermaid', mermaidFixture],
  ['nested-lists', nestedListsFixture],
  ['blockquote', blockquoteFixture],
  ['raw-html', rawHtmlFixture],
  ['links-images', linksImagesFixture],
  ['kitchen-sink', kitchenSinkFixture],
];

// ---------------------------------------------------------------------------
// Download side-effect mocks (URL.createObjectURL + anchor click)
// ---------------------------------------------------------------------------

interface AnchorClick {
  readonly download: string;
  readonly href: string;
}

let createObjectURLMock: ReturnType<typeof vi.spyOn>;
let anchorClicks: AnchorClick[];

beforeEach(() => {
  anchorClicks = [];
  createObjectURLMock = vi.spyOn(URL, 'createObjectURL').mockReturnValue('blob:mock-url');
  vi.spyOn(URL, 'revokeObjectURL').mockImplementation(() => {});
  vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (this: HTMLAnchorElement) {
    anchorClicks.push({ download: this.download, href: this.href });
  });
  // jsdom has no canvas backend (no `canvas` npm package installed), so
  // `HTMLCanvasElement.prototype.toDataURL` returns null instead of a real
  // data URI. exportPdf's page-slicing loop creates its own <canvas> per
  // page internally, so this stub keeps that codepath's output realistic.
  vi.spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockReturnValue('data:image/png;base64,SLICE');
});

afterEach(() => {
  vi.restoreAllMocks();
  html2canvasMock.mockClear();
  jsPdfCtorMock.mockClear();
  jsPdfAddImageMock.mockClear();
  jsPdfAddPageMock.mockClear();
  jsPdfOutputMock.mockClear();
  delete (window as { electronAPI?: unknown }).electronAPI;
  document.body.innerHTML = '';
});

function getDownloadedBlob(callIndex = 0): Blob {
  const call = createObjectURLMock.mock.calls[callIndex] as [Blob] | undefined;
  if (!call) throw new Error(`no URL.createObjectURL call recorded at index ${callIndex}`);
  return call[0];
}

function stabilizeDocxXml(xml: string): string {
  return xml.replace(/r:id="[^"]+"/g, 'r:id="RID"');
}

async function getDocumentXml(blob: Blob): Promise<string> {
  const buffer = await blob.arrayBuffer();
  const zip = await JSZip.loadAsync(buffer);
  const entry = zip.file('word/document.xml');
  if (!entry) throw new Error('word/document.xml missing from exported .docx');
  return entry.async('string');
}

function extractMarkdownBodyHtml(fullHtml: string): string {
  const match = /<div class="markdown-body">\n([\s\S]*?)\n {2}<\/div>/.exec(fullHtml);
  if (!match) throw new Error('could not locate .markdown-body in exported HTML document');
  return match[1]!;
}

// ---------------------------------------------------------------------------
// exportMarkdown
// ---------------------------------------------------------------------------

describe('exportMarkdown', () => {
  it.each(FIXTURES)('round-trips the "%s" fixture byte-for-byte and sanitizes the filename', async (name, markdown) => {
    await exportMarkdown(markdown, `${name}.txt`);

    expect(anchorClicks).toEqual([{ download: `${name}.md`, href: 'blob:mock-url' }]);
    const blob = getDownloadedBlob();
    expect(blob.type).toBe('text/markdown;charset=utf-8');
    await expect(blob.text()).resolves.toBe(markdown);
  });

  it(
    'when window.electronAPI is present, calls saveFile(content, name) POSITIONALLY and skips the browser download ' +
      '(this does not match ElectronAPI.saveFile\'s real single-request-object signature — a pre-existing export/IPC ' +
      'mismatch this suite documents rather than fixes)',
    async () => {
      const saveFileMock = vi.fn().mockResolvedValue(undefined);
      window.electronAPI = { saveFile: saveFileMock } as unknown as typeof window.electronAPI;

      await exportMarkdown(headingsFixture, 'notes.md');

      expect(saveFileMock).toHaveBeenCalledWith(headingsFixture, 'notes.md');
      expect(anchorClicks).toHaveLength(0);
      expect(createObjectURLMock).not.toHaveBeenCalled();
    },
  );
});

// ---------------------------------------------------------------------------
// exportHtml
// ---------------------------------------------------------------------------

describe('exportHtml', () => {
  it.each(FIXTURES)('snapshots the rendered .markdown-body HTML for the "%s" fixture', async (name, markdown) => {
    await exportHtml(markdown, `${name}.md`, 'light');

    const blob = getDownloadedBlob();
    expect(blob.type).toBe('text/html;charset=utf-8');
    const html = await blob.text();
    expect(extractMarkdownBodyHtml(html)).toMatchSnapshot(`${name} body`);
  });

  it('snapshots the full generated HTML document (theme attribute, title escaping, CSS/CDN boilerplate)', async () => {
    await exportHtml(headingsFixture, 'My <Notes> & Ideas.md', 'dracula');

    expect(anchorClicks).toEqual([{ download: 'My <Notes> & Ideas.html', href: 'blob:mock-url' }]);
    const blob = getDownloadedBlob();
    const html = await blob.text();
    expect(html).toMatchSnapshot();
  });
});

// ---------------------------------------------------------------------------
// exportDocx
// ---------------------------------------------------------------------------

describe('exportDocx', () => {
  it.each(FIXTURES)('snapshots the generated word/document.xml structure for the "%s" fixture', async (name, markdown) => {
    await exportDocx(markdown, `${name}.md`);

    expect(anchorClicks).toEqual([{ download: `${name}.docx`, href: 'blob:mock-url' }]);
    const xml = await getDocumentXml(getDownloadedBlob());
    expect(stabilizeDocxXml(xml)).toMatchSnapshot(`${name} document.xml`);
  });
});

// ---------------------------------------------------------------------------
// exportPdf
// ---------------------------------------------------------------------------

describe('exportPdf', () => {
  function mountTarget(id: string): HTMLElement {
    const el = document.createElement('div');
    el.id = id;
    el.textContent = 'content to rasterize';
    document.body.appendChild(el);
    return el;
  }

  it('drives html2canvas with the target element and jsPDF with the resulting image, then downloads a .pdf', async () => {
    const target = mountTarget('markdown-content');

    await exportPdf('markdown-content', 'report.md');

    expect(html2canvasMock).toHaveBeenCalledTimes(1);
    const [calledElement, options] = html2canvasMock.mock.calls[0]!;
    expect(calledElement).toBe(target);
    expect(options).toMatchObject({ scale: 2, useCORS: true, logging: false });

    expect(jsPdfCtorMock).toHaveBeenCalledWith({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    expect(jsPdfAddImageMock).toHaveBeenCalledTimes(1);
    expect(jsPdfAddImageMock).toHaveBeenCalledWith(expect.any(String), 'PNG', 0, 0, 210, expect.any(Number));
    expect(jsPdfAddPageMock).not.toHaveBeenCalled();
    expect(jsPdfOutputMock).toHaveBeenCalledWith('blob');

    expect(anchorClicks).toEqual([{ download: 'report.pdf', href: 'blob:mock-url' }]);
  });

  it('slices a tall canvas across multiple A4 pages', async () => {
    mountTarget('markdown-content');
    html2canvasMock.mockResolvedValueOnce({
      width: 800,
      height: 3000,
      toDataURL: () => 'data:image/png;base64,MOCK',
    });

    await exportPdf('markdown-content', 'long-doc.md');

    // canvasHeightMm = (3000/800)*210 = 787.5mm over three A4 (297mm) pages.
    expect(jsPdfAddImageMock).toHaveBeenCalledTimes(3);
    expect(jsPdfAddPageMock).toHaveBeenCalledTimes(2);
  });

  it('rejects with a friendly, id-specific error when the target element does not exist', async () => {
    await expect(exportPdf('does-not-exist', 'doc.md')).rejects.toThrow(
      '[export] exportPdf: element #does-not-exist not found',
    );
    expect(anchorClicks).toHaveLength(0);
  });
});
