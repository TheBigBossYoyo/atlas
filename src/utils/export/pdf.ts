/**
 * Live-DOM → PDF export (X1).
 *
 * Every exporter here builds a self-contained, sanitized HTML document and
 * hands it to `printHtmlToPdf` (Electron's `webContents.printToPDF`, via the
 * `export:printToPdf` IPC) instead of rasterizing the on-screen viewport with
 * `html2canvas-pro` (the pre-X1 behavior — see git history / P4.4 for its
 * removal). This produces a real vector PDF with selectable text, and — the
 * actual point of X1 — captures the document's FULL content instead of
 * whatever happened to be scrolled into view.
 *
 * - `exportMarkdownPdf` replaces the old generic `exportPdf('markdown-content', ...)`
 *   with a printToPDF-based render of the same live `#markdown-content` DOM
 *   `exportHtml` already serializes, using the same embedded theme CSS.
 * - `exportDocxPdf` reuses DocxViewer's own rendering (all pages already
 *   exist in the live DOM — `PageStack`/`PageView` render every page eagerly,
 *   `zoom` is always `1`, and there is no virtualization to work around) and
 *   its own `@media print` CSS (`viewer-docx.css`/`page-view.css`, harvested
 *   verbatim via `?raw` imports so this can never drift from what the app
 *   actually renders).
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

export const PDF_MIME = 'application/pdf';
export const PDF_FILTERS: readonly SaveFilter[] = [{ name: 'PDF Document', extensions: ['pdf'] }];

const PX_PER_INCH = 96;

// ---------------------------------------------------------------------------
// Markdown
// ---------------------------------------------------------------------------

/**
 * Renders the live `#markdown-content` DOM to a vector PDF using the same
 * theme CSS `exportHtml` embeds in the standalone `.html` export, so the two
 * exports look identical. Replaces the pre-X1 `html2canvas-pro` raster path.
 */
export async function exportMarkdownPdf(elementId: string, fileName: string, theme: string): Promise<void> {
  try {
    const target = document.getElementById(elementId);
    if (!target) {
      throw new Error(`element #${elementId} not found`);
    }

    const css = `${buildMarkdownExportCss()}\n${MARKDOWN_PRINT_PAGE_CSS}`;
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
 * `page-view.css`. `content-visibility: auto` (a perf optimization for the
 * on-screen scroll view — see `page-view.css`) is forced back to `visible`
 * here as a defense-in-depth guarantee: the whole point of X1 is that export
 * must never depend on what happens to be laid out, and this removes any
 * doubt about how a given Chromium version treats `content-visibility` under
 * `printToPDF` rather than trusting it to already do the right thing.
 */
const DOCX_PRINT_OVERRIDES = `
.docx-page {
  content-visibility: visible !important;
  contain: none !important;
}
`;

/**
 * Exports a DOCX document to a real, paginated vector PDF by cloning the
 * live `.docx-page-stack` (every page already rendered — see module header)
 * and reusing DocxViewer's own print CSS instead of a screenshot.
 */
export async function exportDocxPdf(elementId: string, fileName: string): Promise<void> {
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
