/**
 * Markdown → standalone HTML export (X2 / UX-05).
 *
 * Serializes the *already-rendered* `#markdown-content` DOM (the live
 * preview's own output — KaTeX HTML, Mermaid SVG, rehype-highlight code,
 * GFM tables, everything) instead of re-parsing the raw markdown source with
 * a separate `marked` pipeline. The old `marked`-based approach didn't
 * understand `remark-math` delimiters at all (math exported as literal `$..$`
 * text) and could drift from whatever the live preview actually renders;
 * serializing the live DOM makes those two paths structurally the same by
 * construction. This is an export-only change — `MarkdownRenderer.tsx`
 * itself is untouched (on the plan's markdown no-touch list) and still owns
 * all rendering; this module only ever *reads* its output.
 */
import { sanitizeFileName, saveTextOutput, type SaveFilter } from './download';
import { buildMarkdownExportCss } from './markdownExportCss';
import { sanitizeExportHtml } from './sanitizeExportHtml';
import { escapeHtml } from './printDocument';
import { toFriendlyError } from '../friendlyLibraryError';

const HTML_MIME = 'text/html;charset=utf-8';
const HTML_FILTERS: readonly SaveFilter[] = [{ name: 'HTML Document', extensions: ['html'] }];

/**
 * X5 — this exported document is meant to be opened directly in a real
 * browser, outside the app's own CSP/sandbox. `MarkdownRenderer` renders raw
 * HTML embedded in the markdown source via `rehype-raw` with no
 * sanitization step of its own; that is inert in the live React-rendered
 * preview (React never attaches a string as an `onerror`/`onclick`
 * listener, and never executes a `<script>` it creates via
 * `document.createElement`), but a fresh HTML parser — which is exactly
 * what happens when this string is written to a `.html` file and opened —
 * treats those as live markup. `style-src`/`font-src` allow the two CDN
 * origins the head itself links; nothing else may load.
 */
const HTML_EXPORT_CSP =
  "default-src 'none'; img-src data: https:; style-src 'unsafe-inline' https://cdn.jsdelivr.net; font-src data: https://cdn.jsdelivr.net;";

/**
 * Serializes the live-rendered markdown element identified by `elementId`
 * (the same DOM `exportPdf` rasterizes) into a self-contained HTML5 document
 * and downloads it as `.html`. Embeds typography CSS matching the live
 * preview's five themes and links KaTeX + highlight.js CSS from CDN for the
 * math/code markup the live DOM already carries.
 */
export async function exportHtml(elementId: string, fileName: string, theme: string): Promise<void> {
  try {
    const target = document.getElementById(elementId);
    if (!target) {
      throw new Error(`element #${elementId} not found`);
    }

    const name = sanitizeFileName(fileName, 'html');
    const title = fileName.replace(/\.[^./\\]+$/, '');
    const bodyHtml = sanitizeExportHtml(target.outerHTML);

    const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta http-equiv="Content-Security-Policy" content="${HTML_EXPORT_CSP}" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>${escapeHtml(title)}</title>
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.10/dist/katex.min.css" crossorigin="anonymous" />
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/highlight.js@11.9.0/styles/github.min.css" crossorigin="anonymous" />
  <style>
${buildMarkdownExportCss()}
  </style>
</head>
<body data-theme="${escapeHtml(theme)}">
${bodyHtml}
  <footer class="export-footer">Exported from Atlas</footer>
</body>
</html>`;

    await saveTextOutput(html, name, HTML_MIME, HTML_FILTERS);
  } catch (err) {
    console.error('[export] exportHtml failed:', err);
    throw toFriendlyError(err, 'HTML export failed');
  }
}
