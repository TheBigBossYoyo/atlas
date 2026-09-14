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
import { EXPORT_THEME_ORDER, EXPORT_THEME_TOKENS } from './exportThemeTokens';
import { toFriendlyError } from '../friendlyLibraryError';

const HTML_MIME = 'text/html;charset=utf-8';
const HTML_FILTERS: readonly SaveFilter[] = [{ name: 'HTML Document', extensions: ['html'] }];

/** Escape HTML special characters for safe embedding in attributes/text. */
function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function themeBlock(selector: string, tokens: (typeof EXPORT_THEME_TOKENS)[keyof typeof EXPORT_THEME_TOKENS]): string {
  return `${selector} {
  --bg:       ${tokens.bg};
  --text:     ${tokens.text};
  --code-bg:  ${tokens.codeBg};
  --border:   ${tokens.border};
  --accent:   ${tokens.accent};
}`;
}

function buildThemeCss(): string {
  const blocks = EXPORT_THEME_ORDER.map(id =>
    id === 'light' ? themeBlock(':root', EXPORT_THEME_TOKENS[id]) : themeBlock(`[data-theme="${id}"]`, EXPORT_THEME_TOKENS[id]),
  );
  return blocks.join('\n\n');
}

const STRUCTURAL_CSS = `
/* ── Reset & base ─────────────────────────────────────────────────────── */
*, *::before, *::after { box-sizing: border-box; }

body {
  margin: 0;
  background-color: var(--bg);
  color: var(--text);
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif;
  font-size: 16px;
  line-height: 1.7;
  padding: 2rem 1rem;
}

.markdown-body {
  max-width: 860px;
  margin: 0 auto;
  word-wrap: break-word;
}

/* ── Typography ───────────────────────────────────────────────────────── */
.markdown-body h1, .markdown-body h2, .markdown-body h3,
.markdown-body h4, .markdown-body h5, .markdown-body h6 {
  margin-top: 1.5em;
  margin-bottom: 0.5em;
  font-weight: 600;
  line-height: 1.25;
  color: var(--text);
}
.markdown-body h1 { font-size: 2em;    border-bottom: 1px solid var(--border); padding-bottom: 0.3em; }
.markdown-body h2 { font-size: 1.5em;  border-bottom: 1px solid var(--border); padding-bottom: 0.3em; }
.markdown-body h3 { font-size: 1.25em; }
.markdown-body h4 { font-size: 1em;    }
.markdown-body h5 { font-size: 0.875em;}
.markdown-body h6 { font-size: 0.85em; opacity: 0.7; }

.markdown-body p { margin-bottom: 1em; }
.markdown-body a { color: var(--accent); text-decoration: none; }
.markdown-body a:hover { text-decoration: underline; }
.markdown-body strong { font-weight: 700; }
.markdown-body em     { font-style: italic; }
.markdown-body del    { text-decoration: line-through; }

/* ── Code ─────────────────────────────────────────────────────────────── */
.markdown-body code {
  font-family: 'SFMono-Regular', Consolas, 'Liberation Mono', Menlo, monospace;
  font-size: 0.875em;
  background: var(--code-bg);
  border-radius: 4px;
  padding: 0.2em 0.4em;
}
.markdown-body pre {
  background: var(--code-bg);
  border: 1px solid var(--border);
  border-radius: 6px;
  padding: 1em 1.25em;
  overflow-x: auto;
  margin-bottom: 1em;
}
.markdown-body pre code { background: none; padding: 0; font-size: 0.875em; }

/* ── Blockquote ───────────────────────────────────────────────────────── */
.markdown-body blockquote {
  border-left: 4px solid var(--accent);
  padding: 0.5em 1em;
  margin: 0 0 1em 0;
  opacity: 0.85;
  font-style: italic;
}

/* ── Lists ────────────────────────────────────────────────────────────── */
.markdown-body ul, .markdown-body ol { padding-left: 2em; margin-bottom: 1em; }
.markdown-body li { margin-bottom: 0.25em; }
.markdown-body li > ul, .markdown-body li > ol { margin-bottom: 0; }

/* ── Tables ───────────────────────────────────────────────────────────── */
.markdown-body table {
  border-collapse: collapse;
  width: 100%;
  margin-bottom: 1em;
  font-size: 0.9em;
}
.markdown-body th, .markdown-body td {
  border: 1px solid var(--border);
  padding: 0.5em 0.75em;
  text-align: left;
}
.markdown-body th { background: var(--code-bg); font-weight: 600; }

/* ── HR & images ──────────────────────────────────────────────────────── */
.markdown-body hr { border: none; border-top: 1px solid var(--border); margin: 2em 0; }
.markdown-body img { max-width: 100%; height: auto; border-radius: 4px; }

/* ── Mermaid ──────────────────────────────────────────────────────────── */
.markdown-body .mermaid svg { max-width: 100%; height: auto; }

/* ── Footer ───────────────────────────────────────────────────────────── */
.export-footer {
  margin-top: 3rem;
  padding-top: 1rem;
  border-top: 1px solid var(--border);
  font-size: 0.75em;
  opacity: 0.6;
  text-align: center;
}
`.trim();

function buildEmbeddedCss(): string {
  return `${buildThemeCss()}\n\n${STRUCTURAL_CSS}`;
}

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
    const bodyHtml = target.outerHTML;

    const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>${escapeHtml(title)}</title>
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.10/dist/katex.min.css" crossorigin="anonymous" />
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/highlight.js@11.9.0/styles/github.min.css" crossorigin="anonymous" />
  <style>
${buildEmbeddedCss()}
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
