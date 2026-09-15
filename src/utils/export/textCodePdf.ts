/**
 * Text/Code → PDF and → HTML export (X1). `TextViewer` always virtualizes
 * via `react-window` and `CodeViewer` virtualizes above a size threshold
 * (see `shared/sizeThresholds.ts`), so neither viewer's live DOM ever holds
 * the WHOLE file at once. Both exporters here work directly off the file's
 * own raw text content (already in hand as `LoadedFile.content` for a
 * `kind: 'text'` file — no viewer DOM involved at all), so export always has
 * every line regardless of file size.
 *
 * Deliberately plain preformatted text, not syntax-highlighted: reproducing
 * CodeViewer's shiki highlighting off-thread here would mean duplicating a
 * meaningful slice of its logic for a purely cosmetic gain — every line's
 * actual content is what matters for "the export has the full file", which
 * this already guarantees.
 */
import { sanitizeFileName, saveBinaryOutput, saveTextOutput, type SaveFilter } from './download';
import { escapeHtml, buildPrintDocument } from './printDocument';
import { sanitizeExportHtml } from './sanitizeExportHtml';
import { printHtmlToPdf } from './printToPdf';
import { toFriendlyError } from '../friendlyLibraryError';

const PDF_MIME = 'application/pdf';
const PDF_FILTERS: readonly SaveFilter[] = [{ name: 'PDF Document', extensions: ['pdf'] }];
const HTML_MIME = 'text/html;charset=utf-8';
const HTML_FILTERS: readonly SaveFilter[] = [{ name: 'HTML Document', extensions: ['html'] }];

const TEXT_PRINT_CSS = `
@page { size: 8.5in 11in; margin: 0.6in; }
body { margin: 0; background: #ffffff; color: #1f2328; }
pre {
  font-family: 'SFMono-Regular', Consolas, 'Liberation Mono', Menlo, monospace;
  font-size: 9pt;
  line-height: 1.5;
  white-space: pre-wrap;
  word-break: break-word;
  margin: 0;
}
`;

function preformatted(content: string): string {
  return `<pre>${escapeHtml(content)}</pre>`;
}

/** Exports the full raw text/code content as a paginated PDF. */
export async function exportTextPdf(content: string, fileName: string): Promise<void> {
  try {
    const bodyHtml = sanitizeExportHtml(preformatted(content));
    const html = buildPrintDocument({ title: fileName, css: TEXT_PRINT_CSS, bodyHtml });
    const bytes = await printHtmlToPdf(html);
    await saveBinaryOutput(bytes, sanitizeFileName(fileName, 'pdf'), PDF_MIME, PDF_FILTERS);
  } catch (err) {
    console.error('[export] exportTextPdf failed:', err);
    throw toFriendlyError(err, 'PDF export failed');
  }
}

/** Exports the full raw text/code content as a standalone, self-contained `.html` file. */
export async function exportTextHtml(content: string, fileName: string): Promise<void> {
  try {
    const name = sanitizeFileName(fileName, 'html');
    const title = fileName.replace(/\.[^./\\]+$/, '');
    const bodyHtml = sanitizeExportHtml(preformatted(content));
    const html = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline';" />
<title>${escapeHtml(title)}</title>
<style>body{margin:0;padding:1.5rem;background:#fff;color:#1f2328;} pre{font-family:'SFMono-Regular',Consolas,'Liberation Mono',Menlo,monospace;font-size:13px;line-height:1.6;white-space:pre-wrap;word-break:break-word;}</style>
</head>
<body>${bodyHtml}</body>
</html>`;
    await saveTextOutput(html, name, HTML_MIME, HTML_FILTERS);
  } catch (err) {
    console.error('[export] exportTextHtml failed:', err);
    throw toFriendlyError(err, 'HTML export failed');
  }
}
