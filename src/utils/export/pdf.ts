/**
 * Live-DOM → PDF export (X1).
 *
 * Every exporter here builds a self-contained, sanitized HTML document and
 * hands it to `printHtmlToPdf` (Electron's `webContents.printToPDF`, via the
 * `export:printToPdf` IPC) instead of rasterizing the on-screen viewport with
 * `html2canvas-pro` (the pre-X1 behavior — see git history). This produces a
 * real vector PDF with selectable text, and — the actual point of X1 —
 * captures the document's FULL content instead of whatever happened to be
 * scrolled into view.
 *
 * P4.4 note: `html2canvas-pro` itself is NOT removed as a dependency even
 * though this file no longer uses it for PDF export. `docxMedia.ts` (X2, DOCX
 * export's math/Mermaid rasterization — Word has no live LaTeX/diagram
 * renderer, so a picture of the already-rendered output is the only way to
 * embed one) still genuinely needs it. The plan's original P4.4 wording
 * assumed X1 shipping would make the package fully dead; that assumption
 * doesn't hold once `docxMedia.ts`'s independent, ongoing use is accounted
 * for — see notesForMerger.
 *
 * - `exportMarkdownPdf` replaces the old generic `exportPdf('markdown-content', ...)`
 *   with a printToPDF-based render of the same live `#markdown-content` DOM
 *   `exportHtml` already serializes, using the same embedded theme CSS.
 * - `exportDocxPdf` reuses DocxViewer's own rendering (`zoom` is always `1`)
 *   and its own `@media print` CSS (`viewer-docx.css`/`page-view.css`,
 *   harvested verbatim via `?raw` imports so this can never drift from what
 *   the app actually renders). D23-PERF-2 — `PageStack`/`PageView` no longer
 *   render every page eagerly: only pages near the viewport (plus a small
 *   buffer, plus whichever page holds the caret) actually mount, the rest
 *   are lightweight placeholders (see `pageVirtualization.ts`). Before
 *   reading the live DOM below, `exportDocxPdf` dispatches an
 *   `atlas:docx-force-full-render` window event, which `DocxViewer.tsx`
 *   listens for and answers with a SYNCHRONOUS (`flushSync`) full-page
 *   render, so every page is real by the time `dispatchEvent` returns — and
 *   releases it again afterwards with `atlas:docx-release-full-render`.
 * - `exportFlowDocumentPdf` is the RTF/ODT case: both already render their
 *   FULL converted document into the live DOM (no virtualization), so the
 *   same "snapshot the live DOM" strategy applies with each viewer's own CSS.
 *
 * Spreadsheet/CSV/TSV and PPTX/ODP export from their PARSED data instead of
 * the live DOM (both virtualize what's actually mounted) — see
 * `spreadsheetPdf.ts`/`csvPdf.ts`/`slidesPdf.ts`.
 */
import { sanitizeFileName, saveBinaryOutput, type SaveFilter } from './download';
import { buildMarkdownExportCss, MARKDOWN_PRINT_PAGE_CSS } from './markdownExportCss';
import { renderHtmlToPdfFile } from './printDocument';
import { toFriendlyError } from '../friendlyLibraryError';

import viewerDocxCss from '../../viewers/__styles__/viewer-docx.css?raw';
import pageViewCss from '../../docx/render/__styles__/page-view.css?raw';
import viewerRtfCss from '../../viewers/__styles__/viewer-rtf.css?raw';
import viewerOdtCss from '../../viewers/__styles__/viewer-odt.css?raw';

// Markdown-review fix — `buildMarkdownExportCss()` only carries structural
// typography + theme colors; it was never meant to cover KaTeX math or
// rehype-highlight's `.hljs-*` code coloring (those are the LIVE app's own
// globally-loaded `katex/dist/katex.min.css` — see `main.tsx` — and
// `markdownHtml.ts`'s CDN-linked `highlight.js` theme, neither of which
// travels with a detached, serialized DOM string). Without these, a math- or
// code-heavy document exported to PDF renders with unstyled/garbled math and
// plain black code text — a real regression from the pre-X1 html2canvas-pro
// path, which screenshotted the already-CSS-styled live DOM. Harvested the
// same way `viewerDocxCss`/etc. are above so this can never silently drift
// from the real KaTeX/highlight.js versions this app actually bundles.
// Known limitation: KaTeX's own `@font-face` rules point at relative
// `fonts/*.woff2` paths that don't resolve from the print window's temp-file
// location, so math falls back to a substitute font — layout/spacing is
// still correct (that's what most of this CSS governs), only the exact
// glyph shapes differ; embedding the actual font files as `data:` URIs would
// fix this fully but is a larger follow-up, not done here.
import katexCss from 'katex/dist/katex.min.css?raw';
import hljsCss from 'highlight.js/styles/github.min.css?raw';

export const PDF_MIME = 'application/pdf';
export const PDF_FILTERS: readonly SaveFilter[] = [{ name: 'PDF Document', extensions: ['pdf'] }];

const PX_PER_INCH = 96;

// ---------------------------------------------------------------------------
// Markdown
// ---------------------------------------------------------------------------

/**
 * Renders the live `#markdown-content` DOM to a vector PDF using the same
 * theme CSS `exportHtml` embeds in the standalone `.html` export, plus the
 * real KaTeX/highlight.js CSS (see the import comment above) so math and
 * code blocks aren't left unstyled the way a plain CSS-less `<style>` from
 * `buildMarkdownExportCss()` alone would leave them. Replaces the pre-X1
 * `html2canvas-pro` raster path.
 */
export async function exportMarkdownPdf(elementId: string, fileName: string, theme: string): Promise<void> {
  try {
    const target = document.getElementById(elementId);
    if (!target) {
      throw new Error(`element #${elementId} not found`);
    }

    const css = `${buildMarkdownExportCss()}\n${katexCss}\n${hljsCss}\n${MARKDOWN_PRINT_PAGE_CSS}`;
    const bodyHtml = `<div data-theme="${theme.replace(/"/g, '')}">${target.outerHTML}</div>`;
    await renderHtmlToPdfFile(bodyHtml, css, fileName, fileName);
  } catch (err) {
    console.error('[export] exportMarkdownPdf failed:', err);
    throw toFriendlyError(err, 'PDF export failed');
  }
}

// ---------------------------------------------------------------------------
// DOCX
// ---------------------------------------------------------------------------

function parsePagePx(value: string): number | null {
  const match = /^([\d.]+)px$/.exec(value.trim());
  return match ? Number.parseFloat(match[1]) : null;
}

interface PageSizePx {
  readonly width: number;
  readonly height: number;
}

function readPageSizePx(pageEl: HTMLElement): PageSizePx | null {
  const width = parsePagePx(pageEl.style.width);
  const height = parsePagePx(pageEl.style.height);
  return width !== null && height !== null ? { width, height } : null;
}

function pageSizeKey(size: PageSizePx): string {
  return `${Math.round(size.width)}x${Math.round(size.height)}`;
}

/**
 * Builds the `@page` CSS for a DOCX export from each page's own rendered
 * pixel size (`PageView` always renders at `zoom={1}`, so a page's inline
 * `width`/`height` style — set from `page.sizePt` at 96dpi — already IS its
 * true physical size; no unit conversion beyond px→in is needed). A DOCX
 * with more than one distinct section page size uses one named `@page` rule
 * per distinct size (Chromium's CSS Paged Media support); Chromium's
 * `printToPDF` only honors ONE page box's `size` for the whole document in
 * practice, so a multi-size DOCX still prints every page at the FIRST page's
 * size — a known limitation (documented for the merger), not a silent bug.
 */
function buildDocxPageCss(pageEls: readonly HTMLElement[]): string {
  const bySize = new Map<string, PageSizePx>();
  for (const el of pageEls) {
    const size = readPageSizePx(el);
    if (size) bySize.set(pageSizeKey(size), size);
  }

  if (bySize.size <= 1) {
    const [size] = bySize.values();
    const widthIn = (size?.width ?? 8.5 * PX_PER_INCH) / PX_PER_INCH;
    const heightIn = (size?.height ?? 11 * PX_PER_INCH) / PX_PER_INCH;
    return `@page { size: ${widthIn}in ${heightIn}in; margin: 0; }`;
  }

  const rules: string[] = [];
  let index = 0;
  for (const [key, size] of bySize) {
    const name = `docx-size-${index++}`;
    const widthIn = size.width / PX_PER_INCH;
    const heightIn = size.height / PX_PER_INCH;
    rules.push(`@page ${name} { size: ${widthIn}in ${heightIn}in; margin: 0; }`);
    rules.push(`.docx-page[data-size-key="${key}"] { page: ${name}; }`);
  }
  return rules.join('\n');
}

/**
 * Print-only overrides layered on top of the harvested `viewer-docx.css`/
 * `page-view.css`.
 *
 * - `content-visibility: auto` (a perf optimization for the on-screen scroll
 *   view — see `page-view.css`) is forced back to `visible` here as a
 *   defense-in-depth guarantee: the whole point of X1 is that export must
 *   never depend on what happens to be laid out, and this removes any doubt
 *   about how a given Chromium version treats `content-visibility` under
 *   `printToPDF` rather than trusting it to already do the right thing.
 * - `.docx-page-stack`'s on-screen `padding`/`gap`/`flex` (chrome for the
 *   scrollable "sheets of paper on a desk" view — see `page-view.css`) is
 *   zeroed and switched to `display: block` for print. Left in place, that
 *   padding/gap accumulates extra height ABOVE and BETWEEN each already
 *   exactly-one-physical-page-tall `.docx-page` box; Chromium's print
 *   pagination flows the whole document by physical page height, so that
 *   extra height pushes each page's own trailing content onto the NEXT
 *   printed page — compounding page over page until a 2-page document
 *   prints as 3+ pages (the concrete regression this fixes). An explicit
 *   `page-break-after` on each `.docx-page` (mirroring `slidesPdf.ts`'s
 *   `.export-slide`) makes the one-page-per-`.docx-page` pagination
 *   independent of exact height-matching entirely, rather than relying on
 *   it as the pre-fix code implicitly did.
 */
const DOCX_PRINT_OVERRIDES = `
.docx-page-stack {
  display: block !important;
  padding: 0 !important;
  gap: 0 !important;
  background: none !important;
}
.docx-page {
  content-visibility: visible !important;
  contain: none !important;
  margin: 0 !important;
  box-shadow: none !important;
  page-break-after: always;
  break-after: page;
}
.docx-page:last-child {
  page-break-after: auto;
  break-after: auto;
}
`;

/**
 * Exports a DOCX document to a real, paginated vector PDF by cloning the
 * live `.docx-page-stack` (every page already rendered — see module header)
 * and reusing DocxViewer's own print CSS instead of a screenshot.
 */
/** D23-PERF-2 — see the module header: asks the live `DocxViewer` to mount
 * every page before/after this export reads the DOM. A no-op (no listener
 * ever answers) for any other viewer or if DocxViewer isn't mounted at all,
 * which is fine — `exportDocxPdf` still checks for `.docx-page-stack`/
 * `.docx-page` below and fails with its existing, already-user-facing error
 * message rather than silently exporting an empty document either way. */
function dispatchDocxFullRender(eventName: 'atlas:docx-force-full-render' | 'atlas:docx-release-full-render'): void {
  window.dispatchEvent(new Event(eventName));
}

export async function exportDocxPdf(elementId: string, fileName: string): Promise<void> {
  dispatchDocxFullRender('atlas:docx-force-full-render');
  try {
    const root = document.getElementById(elementId);
    if (!root) {
      throw new Error(`element #${elementId} not found`);
    }

    const stack = root.querySelector('.docx-page-stack');
    if (!stack) {
      throw new Error('the document has not finished loading yet — try exporting again in a moment');
    }

    const clone = stack.cloneNode(true) as HTMLElement;
    const pageEls = Array.from(clone.querySelectorAll<HTMLElement>('.docx-page'));
    if (pageEls.length === 0) {
      throw new Error('no pages to export');
    }

    for (const el of pageEls) {
      const size = readPageSizePx(el);
      if (size) el.setAttribute('data-size-key', pageSizeKey(size));
    }

    const pageCss = buildDocxPageCss(pageEls);
    const css = `${viewerDocxCss}\n${pageViewCss}\n${pageCss}\n${DOCX_PRINT_OVERRIDES}`;
    await renderHtmlToPdfFile(clone.outerHTML, css, fileName, fileName);
  } catch (err) {
    console.error('[export] exportDocxPdf failed:', err);
    throw toFriendlyError(err, 'PDF export failed');
  } finally {
    dispatchDocxFullRender('atlas:docx-release-full-render');
  }
}

// ---------------------------------------------------------------------------
// RTF / ODT — both render their full converted document into the live DOM
// with no virtualization, so (like DOCX) the live DOM already has everything
// needed; unlike DOCX they have no inherent "page" geometry (odf-kit/rtf.js
// produce a flowing document, not discrete pages), so a standard Letter
// sheet with real margins is used.
// ---------------------------------------------------------------------------

const FLOW_PRINT_PAGE_CSS = `
@page { size: 8.5in 11in; margin: 0.75in; }
`;

/**
 * Print-only override: the on-screen viewer owns its own scrollbar via a
 * fixed `height: 100%; overflow: auto` (see `viewer-rtf.css`/`viewer-odt.css`
 * — needed because the shared `.content--viewer .preview-panel` sets
 * `overflow: hidden`). That has nothing to clip against once cloned into a
 * detached, unstyled print document, but forcing it explicitly removes any
 * doubt that a 100%-of-nothing height could truncate the printed output.
 */
const FLOW_PRINT_OVERRIDES = `
.rtf-viewer, .odt-viewer, .odt-viewer__body {
  height: auto !important;
  overflow: visible !important;
}
`;

async function exportFlowDocumentPdf(elementId: string, fileName: string, viewerCss: string): Promise<void> {
  const root = document.getElementById(elementId);
  if (!root) {
    throw new Error(`element #${elementId} not found`);
  }

  const css = `${viewerCss}\n${FLOW_PRINT_PAGE_CSS}\n${FLOW_PRINT_OVERRIDES}`;
  await renderHtmlToPdfFile(root.outerHTML, css, fileName, fileName);
}

export async function exportRtfPdf(elementId: string, fileName: string): Promise<void> {
  try {
    await exportFlowDocumentPdf(elementId, fileName, viewerRtfCss);
  } catch (err) {
    console.error('[export] exportRtfPdf failed:', err);
    throw toFriendlyError(err, 'PDF export failed');
  }
}

export async function exportOdtPdf(elementId: string, fileName: string): Promise<void> {
  try {
    await exportFlowDocumentPdf(elementId, fileName, viewerOdtCss);
  } catch (err) {
    console.error('[export] exportOdtPdf failed:', err);
    throw toFriendlyError(err, 'PDF export failed');
  }
}

// ---------------------------------------------------------------------------
// PDF passthrough (verified/fixed — UX-11: this used to bypass the native
// save dialog with a raw browser download even inside Electron).
// ---------------------------------------------------------------------------

/** "Save a copy" for an already-PDF document: copies the original bytes verbatim. */
export async function exportPdfCopy(content: ArrayBuffer, fileName: string): Promise<void> {
  try {
    await saveBinaryOutput(content, sanitizeFileName(fileName, 'pdf'), PDF_MIME, PDF_FILTERS);
  } catch (err) {
    console.error('[export] exportPdfCopy failed:', err);
    throw toFriendlyError(err, 'PDF export failed');
  }
}
