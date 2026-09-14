/**
 * src/utils/export/* — characterization tests (P0.4 / QA-07, updated for X2/UX-05).
 *
 * X2 replaced the old single `export.ts` (which re-parsed raw markdown with a
 * separate `marked` pipeline for HTML export) with the small `export/*`
 * modules: `exportHtml` now serializes the *already-rendered* `#markdown-content`
 * DOM (the live preview's own KaTeX/Mermaid/highlight.js output) instead of
 * re-parsing; `exportDocx` still uses `marked`'s token lexer (Word has no
 * remark/rehype-equivalent renderer to reuse) but now rasterizes math/Mermaid
 * to images instead of leaking raw source text, fixes the dropped-hyperlink-text
 * bug, and renders task-list checkboxes as real glyphs. Per the improvement
 * plan's Section 7 guardrail, these snapshots are intentionally different from
 * the pre-X2 baseline — this file documents the *new* frozen behavior, not a
 * regression. `MarkdownRenderer`/`useToc`'s own *live-render* snapshots are a
 * separate suite and are untouched by this change.
 *
 * exportDocx's docx-package-generated relationship ids (`r:id="..."` for
 * hyperlinks, `r:embed="..."` for embedded images) are regenerated fresh on
 * every `Packer.toBlob` call, so both are sanitized to fixed placeholders
 * before snapshotting — see `stabilizeDocxXml`. Embedded image *bytes* come
 * from a mocked `html2canvas-pro` (jsdom has no canvas backend), so they are
 * deterministic and safe to snapshot as-is.
 */

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { render, waitFor, cleanup } from '@testing-library/react';
import { MarkdownRenderer } from '../../components/MarkdownRenderer';

// ---------------------------------------------------------------------------
// mermaid mock — jsdom cannot render real Mermaid diagrams (same rationale as
// MarkdownRenderer.characterization.test.tsx). Both `exportHtml` (via
// rendering the fixture through the real MarkdownRenderer first) and
// `exportDocx` (via `docxMedia.ts`'s dynamic `import('mermaid')`) go through
// this mock.
// ---------------------------------------------------------------------------

const { mermaidRenderMock, mermaidInitializeMock } = vi.hoisted(() => ({
  mermaidRenderMock: vi.fn(async (id: string, code: string) => {
    void id;
    void code;
    return { svg: '<svg data-mock-mermaid="true" role="img"><text>MOCKED MERMAID SVG</text></svg>' };
  }),
  mermaidInitializeMock: vi.fn(),
}));

vi.mock('mermaid', () => ({
  default: {
    initialize: mermaidInitializeMock,
    render: mermaidRenderMock,
  },
}));

// ---------------------------------------------------------------------------
// html2canvas-pro / jsPDF mocks — jsdom has no canvas backend, so neither
// exportPdf's screenshot pipeline nor docxMedia.ts's math/Mermaid
// rasterization can run for real here. `html2canvasMock` backs both.
// ---------------------------------------------------------------------------

const { html2canvasMock, jsPdfCtorMock, jsPdfAddImageMock, jsPdfAddPageMock, jsPdfOutputMock } = vi.hoisted(() => {
  const addImageMock = vi.fn();
  const addPageMock = vi.fn();
  const outputMock = vi.fn(() => new Uint8Array([1, 2, 3]).buffer);
  return {
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
});

afterEach(() => {
  vi.restoreAllMocks();
  html2canvasMock.mockClear();
  jsPdfCtorMock.mockClear();
  jsPdfAddImageMock.mockClear();
  jsPdfAddPageMock.mockClear();
  jsPdfOutputMock.mockClear();
  mermaidRenderMock.mockClear();
  mermaidInitializeMock.mockClear();
  delete (window as { electronAPI?: unknown }).electronAPI;
  cleanup();
  document.body.innerHTML = '';
});

function getDownloadedBlob(callIndex = 0): Blob {
  const call = createObjectURLMock.mock.calls[callIndex] as [Blob] | undefined;
  if (!call) throw new Error(`no URL.createObjectURL call recorded at index ${callIndex}`);
  return call[0];
}

/** Normalizes docx-package-generated relationship ids (hyperlinks use
 * `r:id`, embedded images use `r:embed`) so byte-identical fixture content
 * snapshots the same way across runs despite fresh ids every `Packer.toBlob`. */
function stabilizeDocxXml(xml: string): string {
  return xml.replace(/r:id="[^"]+"/g, 'r:id="RID"').replace(/r:embed="[^"]+"/g, 'r:embed="REMBED"');
}

async function getDocumentXml(blob: Blob): Promise<string> {
  const buffer = await blob.arrayBuffer();
  const zip = await JSZip.loadAsync(buffer);
  const entry = zip.file('word/document.xml');
  if (!entry) throw new Error('word/document.xml missing from exported .docx');
  return entry.async('string');
}

/** Mirrors MarkdownRenderer.characterization.test.tsx's normalization of the
 * non-deterministic `useId()`-derived tokens (React's `:rN:` ids and the
 * Mermaid-diagram id built from one) that would otherwise make an
 * `outerHTML` snapshot flaky across runs/positions in the file. */
function stabilizeHtml(html: string): string {
  return html
    .replace(/:r[0-9a-z]+:/gi, ':rSTABLE:')
    .replace(/mermaid-[a-zA-Z0-9]+-\d+/g, 'mermaid-STABLE-ID');
}

function extractMarkdownBodyHtml(fullHtml: string): string {
  const match = /<body data-theme="[^"]*">\n([\s\S]*?)\n {2}<footer/.exec(fullHtml);
  if (!match) throw new Error('could not locate <body>...<footer> in exported HTML document');
  return match[1]!;
}

/** Renders `markdown` through the real (unmodified) `MarkdownRenderer` — the
 * same live component the app renders — so `exportHtml` has real
 * `#markdown-content` DOM (KaTeX HTML, mocked Mermaid SVG, rehype-highlight
 * code, GFM tables, ...) to serialize, exactly like the running app. */
async function renderMarkdownContent(markdown: string) {
  const utils = render(<MarkdownRenderer markdown={markdown} />);
  if (markdown.includes('```mermaid')) {
    await waitFor(() => {
      expect(utils.container.querySelector('[data-mock-mermaid]')).toBeTruthy();
    });
  }
  return utils;
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
    'when window.electronAPI is present, calls saveFile with the real single-request-object signature ' +
      '(wave1 follow-up: this used to call saveFile(content, name) POSITIONALLY, silently skipping the ' +
      "IPC call's actual { content, suggestedName, filters? } shape)",
    async () => {
      const saveFileMock = vi.fn().mockResolvedValue({ saved: true });
      window.electronAPI = { saveFile: saveFileMock } as unknown as typeof window.electronAPI;

      await exportMarkdown(headingsFixture, 'notes.md');

      expect(saveFileMock).toHaveBeenCalledWith({
        content: headingsFixture,
        suggestedName: 'notes.md',
        filters: [{ name: 'Markdown', extensions: ['md'] }],
      });
      expect(anchorClicks).toHaveLength(0);
      expect(createObjectURLMock).not.toHaveBeenCalled();
    },
  );

  it('throws a friendly error when the Electron save dialog reports a failure', async () => {
    window.electronAPI = {
      saveFile: vi.fn().mockResolvedValue({ saved: false, error: 'Disk is full' }),
    } as unknown as typeof window.electronAPI;

    await expect(exportMarkdown(headingsFixture, 'notes.md')).rejects.toThrow(
      'Markdown export failed: Disk is full',
    );
  });

  it('resolves quietly (no error, no browser fallback) when the user cancels the Electron save dialog', async () => {
    window.electronAPI = {
      saveFile: vi.fn().mockResolvedValue({ saved: false }),
    } as unknown as typeof window.electronAPI;

    await expect(exportMarkdown(headingsFixture, 'notes.md')).resolves.toBeUndefined();
    expect(anchorClicks).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// exportHtml
// ---------------------------------------------------------------------------

describe('exportHtml', () => {
  it.each(FIXTURES)('snapshots the serialized live-DOM body for the "%s" fixture', async (name, markdown) => {
    await renderMarkdownContent(markdown);

    await exportHtml('markdown-content', `${name}.md`, 'light');

    const blob = getDownloadedBlob();
    expect(blob.type).toBe('text/html;charset=utf-8');
    const html = await blob.text();
    expect(stabilizeHtml(extractMarkdownBodyHtml(html))).toMatchSnapshot(`${name} body`);
  });

  it('serializes the exact live #markdown-content element, not a re-parse of the source', async () => {
    await renderMarkdownContent('# Hello\n\nSome **bold** text.');

    const live = document.getElementById('markdown-content')!;
    // Mutate the live DOM after render (simulating whatever KaTeX/Mermaid/
    // highlight.js already did to it) — exportHtml must reflect this exact
    // markup, proving it reads the rendered DOM rather than re-deriving HTML
    // from the raw markdown string a second time.
    const marker = document.createElement('span');
    marker.className = 'export-dom-identity-probe';
    marker.textContent = 'proof';
    live.appendChild(marker);

    await exportHtml('markdown-content', 'doc.md', 'light');

    const html = await getDownloadedBlob().text();
    expect(html).toContain('export-dom-identity-probe');
  });

  it('snapshots the full generated HTML document (theme attribute, title escaping, CSS/CDN boilerplate)', async () => {
    await renderMarkdownContent(headingsFixture);

    await exportHtml('markdown-content', 'My <Notes> & Ideas.md', 'dracula');

    expect(anchorClicks).toEqual([{ download: 'My <Notes> & Ideas.html', href: 'blob:mock-url' }]);
    const blob = getDownloadedBlob();
    const html = await blob.text();
    expect(stabilizeHtml(html)).toMatchSnapshot();
  });

  it('rejects with a friendly, id-specific error when the target element does not exist', async () => {
    await expect(exportHtml('does-not-exist', 'doc.md', 'light')).rejects.toThrow(
      'HTML export failed: element #does-not-exist not found',
    );
    expect(anchorClicks).toHaveLength(0);
  });

  it('routes through the Electron save dialog (UX-11) with an HTML filter when window.electronAPI is present', async () => {
    await renderMarkdownContent('# Hi');
    const saveFileMock = vi.fn().mockResolvedValue({ saved: true });
    window.electronAPI = { saveFile: saveFileMock } as unknown as typeof window.electronAPI;

    await exportHtml('markdown-content', 'doc.md', 'light');

    expect(saveFileMock).toHaveBeenCalledWith(
      expect.objectContaining({
        suggestedName: 'doc.html',
        filters: [{ name: 'HTML Document', extensions: ['html'] }],
      }),
    );
    expect(createObjectURLMock).not.toHaveBeenCalled();
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

  it('preserves hyperlink visible text (wave1 follow-up fix — used to spread a built TextRun, silently dropping it)', async () => {
    await exportDocx(linksImagesFixture, 'doc.md');
    const xml = await getDocumentXml(getDownloadedBlob());

    expect(xml).toContain('Atlas repository');
    // The old, buggy shape had an empty run inside every hyperlink:
    // <w:hyperlink ...><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr></w:r></w:hyperlink>
    expect(xml).not.toMatch(/<w:hyperlink[^>]*><w:r><w:rPr><w:rStyle w:val="Hyperlink"\/><\/w:rPr><\/w:r><\/w:hyperlink>/);
  });

  it('renders task-list checkboxes as ☑/☐ glyphs, not raw "[x] "/"[ ] " syntax (UX-06)', async () => {
    await exportDocx(taskListFixture, 'doc.md');
    const xml = await getDocumentXml(getDownloadedBlob());

    expect(xml).toContain('☑');
    expect(xml).toContain('☐');
    expect(xml).not.toContain('[x] ');
    expect(xml).not.toContain('[ ] ');
  });

  it('rasterizes inline and block math to embedded images instead of raw "$..$"/"$$..$$" text', async () => {
    await exportDocx(mathFixture, 'doc.md');
    const xml = await getDocumentXml(getDownloadedBlob());

    expect(xml).toContain('<w:drawing>');
    expect(xml).not.toContain('$E = mc^2$');
    expect(html2canvasMock).toHaveBeenCalled();
  });

  it('rasterizes a ```mermaid fence to an embedded image instead of raw diagram source', async () => {
    await exportDocx(mermaidFixture, 'doc.md');
    const xml = await getDocumentXml(getDownloadedBlob());

    expect(xml).toContain('<w:drawing>');
    expect(xml).not.toContain('graph TD;');
    expect(mermaidRenderMock).toHaveBeenCalledTimes(1);
  });

  it('falls back to a visible literal instead of silently dropping math when rasterization fails', async () => {
    html2canvasMock.mockRejectedValueOnce(new Error('canvas failure'));
    await exportDocx('Inline math: $E = mc^2$ done.', 'doc.md');
    const xml = await getDocumentXml(getDownloadedBlob());

    expect(xml).toContain('$E = mc^2$');
  });

  it('routes through the Electron save dialog (UX-11) with a .docx filter when window.electronAPI is present', async () => {
    const saveBinaryFileMock = vi.fn().mockResolvedValue({ saved: true });
    window.electronAPI = { saveBinaryFile: saveBinaryFileMock } as unknown as typeof window.electronAPI;

    await exportDocx(headingsFixture, 'doc.md');

    expect(saveBinaryFileMock).toHaveBeenCalledWith(
      expect.objectContaining({
        suggestedName: 'doc.docx',
        filters: [{ name: 'Word Document', extensions: ['docx'] }],
      }),
    );
    expect(createObjectURLMock).not.toHaveBeenCalled();
  });

  it('wraps a thrown library error in a friendly, format-specific message (RUN-14)', async () => {
    window.electronAPI = {
      saveBinaryFile: vi.fn().mockRejectedValue(new Error("Can't find end of central directory")),
    } as unknown as typeof window.electronAPI;

    await expect(exportDocx(headingsFixture, 'doc.md')).rejects.toThrow(
      'DOCX export failed: the file could not be read as a valid Office document (it may be corrupted or not a real Office file)',
    );
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

  beforeEach(() => {
    // jsdom has no canvas backend (no `canvas` npm package installed), so
    // `HTMLCanvasElement.prototype.toDataURL` returns null instead of a real
    // data URI. exportPdf's page-slicing loop creates its own <canvas> per
    // page internally, so this stub keeps that codepath's output realistic.
    vi.spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockReturnValue('data:image/png;base64,SLICE');
  });

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
    expect(jsPdfOutputMock).toHaveBeenCalledWith('arraybuffer');

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

  it('routes through the Electron save dialog (UX-11) with a .pdf filter when window.electronAPI is present', async () => {
    mountTarget('markdown-content');
    const saveBinaryFileMock = vi.fn().mockResolvedValue({ saved: true });
    window.electronAPI = { saveBinaryFile: saveBinaryFileMock } as unknown as typeof window.electronAPI;

    await exportPdf('markdown-content', 'report.md');

    expect(saveBinaryFileMock).toHaveBeenCalledWith(
      expect.objectContaining({
        suggestedName: 'report.pdf',
        filters: [{ name: 'PDF Document', extensions: ['pdf'] }],
      }),
    );
    expect(createObjectURLMock).not.toHaveBeenCalled();
  });
});
