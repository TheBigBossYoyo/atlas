/**
 * Builds the self-contained HTML5 document string sent to
 * `export:printToPdf`. One shared builder so every per-format exporter
 * produces a document with the same restrictive CSP meta tag (X5) instead of
 * each hand-rolling its own `<html>` wrapper.
 *
 * The CSP here is deliberately far stricter than the live app's own
 * (`electron/lib/csp.cjs`): this document never needs to run script (the
 * hidden print window also has `javascript: false` — see
 * `electron/lib/printToPdf.cjs` — so this is defense in depth, not the only
 * guard) or fetch anything remote. `data:` stays allowed for `img-src`/
 * `font-src` because DOCX/slide exports embed images and Atlas's own
 * bundled substitute fonts as `data:` URIs — nothing this codebase produces
 * ever references a `file:` URI (found during wave3/export review: the
 * original policy allowed it too, which would have let a sanitizer bypass
 * load an arbitrary local file into the hidden print window; DOMPurify
 * already strips `file:` `src`/`href` values by default, so this is
 * defense-in-depth tightening, not a fix for a reachable bug).
 */
import { sanitizeExportHtml } from './sanitizeExportHtml';
import { printHtmlToPdf } from './printToPdf';
import { sanitizeFileName, saveBinaryOutput } from './download';

/** Escape HTML special characters for safe embedding in text/attribute positions. */
export function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

const EXPORT_CSP = "default-src 'none'; img-src data:; font-src data:; style-src 'unsafe-inline';";

export interface PrintDocumentOptions {
  /** Used for the document `<title>` only — never rendered as page content. */
  readonly title: string;
  /** Trusted CSS (from this codebase's own `?raw` stylesheet imports or hand-authored rules) — NOT sanitized, must not come from file content. */
  readonly css: string;
  /** Body markup — callers are responsible for sanitizing anything derived from untrusted file content via `sanitizeExportHtml` before passing it here. */
  readonly bodyHtml: string;
}

/** Builds a complete, standalone HTML5 document ready for `printHtmlToPdf`. */
export function buildPrintDocument({ title, css, bodyHtml }: PrintDocumentOptions): string {
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta http-equiv="Content-Security-Policy" content="${EXPORT_CSP}" />
<title>${escapeHtml(title)}</title>
<style>${css}</style>
</head>
<body>${bodyHtml}</body>
</html>`;
}

/**
 * Sanitizes `bodyHtmlRaw` (untrusted — derived from file content), wraps it
 * in a print document with `css` (trusted — from this codebase's own CSS
 * files/hand-authored rules, never from file content), sends it through
 * `printHtmlToPdf`, and saves the resulting bytes via the native save
 * dialog. The one shared "finish the export" step every per-format PDF
 * exporter in `pdf.ts`/`spreadsheetPdf.ts`/`csvPdf.ts`/`slidesPdf.ts`/
 * `textCodeExport.ts` ends with.
 */
export async function renderHtmlToPdfFile(
  bodyHtmlRaw: string,
  css: string,
  title: string,
  fileName: string,
): Promise<void> {
  const bodyHtml = sanitizeExportHtml(bodyHtmlRaw);
  const html = buildPrintDocument({ title, css, bodyHtml });
  const bytes = await printHtmlToPdf(html);
  await saveBinaryOutput(bytes, sanitizeFileName(fileName, 'pdf'), 'application/pdf', [
    { name: 'PDF Document', extensions: ['pdf'] },
  ]);
}
