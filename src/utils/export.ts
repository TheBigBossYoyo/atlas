import { marked } from 'marked';
import type { Tokens } from 'marked';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  ExternalHyperlink,
} from 'docx';
import type { IPropertiesOptions } from 'docx';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas-pro';

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/** Trigger a browser file download from a Blob, then revoke the object URL. */
function triggerDownload(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  anchor.style.display = 'none';
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  // Revoke after a short delay so the browser has time to start the download
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/** Strip any existing file extension from `name` and append `.{ext}`. */
function sanitizeFileName(name: string, ext: string): string {
  const stripped = name.replace(/\.[^./\\]+$/, '');
  return `${stripped}.${ext}`;
}

// ---------------------------------------------------------------------------
// Embedded CSS for HTML export
// ---------------------------------------------------------------------------

const EMBEDDED_CSS = `
/* ── Reset & base ─────────────────────────────────────────────────────── */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  --bg:       #ffffff;
  --text:     #1f2328;
  --code-bg:  #eff1f3;
  --border:   #d0d7de;
  --accent:   #0969da;
}

[data-theme="dark"] {
  --bg:       #0d1117;
  --text:     #e6edf3;
  --code-bg:  #262c36;
  --border:   #30363d;
  --accent:   #58a6ff;
}

[data-theme="sepia"] {
  --bg:       #f4ecd8;
  --text:     #5b4636;
  --code-bg:  #ebe2c8;
  --border:   #d4c9a8;
  --accent:   #8b5a2b;
}

[data-theme="nord"] {
  --bg:       #2e3440;
  --text:     #eceff4;
  --code-bg:  #3b4252;
  --border:   #434c5e;
  --accent:   #88c0d0;
}

[data-theme="dracula"] {
  --bg:       #282a36;
  --text:     #f8f8f2;
  --code-bg:  #44475a;
  --border:   #44475a;
  --accent:   #bd93f9;
}

/* ── Layout ───────────────────────────────────────────────────────────── */
body {
  background-color: var(--bg);
  color: var(--text);
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif;
  font-size: 16px;
  line-height: 1.7;
  padding: 2rem 1rem;
}

.markdown-body {
  max-width: 800px;
  margin: 0 auto;
}

/* ── Typography ───────────────────────────────────────────────────────── */
h1, h2, h3, h4, h5, h6 {
  margin-top: 1.5em;
  margin-bottom: 0.5em;
  font-weight: 600;
  line-height: 1.25;
  color: var(--text);
}
h1 { font-size: 2em;    border-bottom: 1px solid var(--border); padding-bottom: 0.3em; }
h2 { font-size: 1.5em;  border-bottom: 1px solid var(--border); padding-bottom: 0.3em; }
h3 { font-size: 1.25em; }
h4 { font-size: 1em;    }
h5 { font-size: 0.875em;}
h6 { font-size: 0.85em; color: color-mix(in srgb, var(--text) 70%, transparent); }

p { margin-bottom: 1em; }

a { color: var(--accent); text-decoration: none; }
a:hover { text-decoration: underline; }

strong { font-weight: 700; }
em     { font-style: italic; }
del    { text-decoration: line-through; }

/* ── Code ─────────────────────────────────────────────────────────────── */
code {
  font-family: 'SFMono-Regular', Consolas, 'Liberation Mono', Menlo, monospace;
  font-size: 0.875em;
  background: var(--code-bg);
  border-radius: 4px;
  padding: 0.2em 0.4em;
}

pre {
  background: var(--code-bg);
  border: 1px solid var(--border);
  border-radius: 6px;
  padding: 1em 1.25em;
  overflow-x: auto;
  margin-bottom: 1em;
}
pre code { background: none; padding: 0; font-size: 0.875em; }

/* ── Blockquote ───────────────────────────────────────────────────────── */
blockquote {
  border-left: 4px solid var(--accent);
  padding: 0.5em 1em;
  margin: 0 0 1em 0;
  color: color-mix(in srgb, var(--text) 75%, transparent);
  font-style: italic;
}

/* ── Lists ────────────────────────────────────────────────────────────── */
ul, ol { padding-left: 2em; margin-bottom: 1em; }
li { margin-bottom: 0.25em; }
li > ul, li > ol { margin-bottom: 0; }

/* ── Tables ───────────────────────────────────────────────────────────── */
table {
  border-collapse: collapse;
  width: 100%;
  margin-bottom: 1em;
  font-size: 0.9em;
}
th, td {
  border: 1px solid var(--border);
  padding: 0.5em 0.75em;
  text-align: left;
}
th { background: var(--code-bg); font-weight: 600; }
tr:nth-child(even) td { background: color-mix(in srgb, var(--code-bg) 40%, transparent); }

/* ── HR ───────────────────────────────────────────────────────────────── */
hr {
  border: none;
  border-top: 1px solid var(--border);
  margin: 2em 0;
}

/* ── Images ───────────────────────────────────────────────────────────── */
img { max-width: 100%; height: auto; border-radius: 4px; }

/* ── Footer ───────────────────────────────────────────────────────────── */
.export-footer {
  margin-top: 3rem;
  padding-top: 1rem;
  border-top: 1px solid var(--border);
  font-size: 0.75em;
  color: color-mix(in srgb, var(--text) 50%, transparent);
  text-align: center;
}
`.trim();

// ---------------------------------------------------------------------------
// Exported functions
// ---------------------------------------------------------------------------

/**
 * Download the raw markdown source as a `.md` file.
 * Uses Electron's saveFile API when available.
 */
export async function exportMarkdown(markdown: string, fileName: string): Promise<void> {
  try {
    const name = sanitizeFileName(fileName, 'md');
    const electronSave = (window as unknown as Record<string, unknown>)['electronAPI'] as
      | { saveFile?: (content: string, suggestedName: string) => Promise<void> }
      | undefined;
    if (electronSave?.saveFile) {
      await electronSave.saveFile(markdown, name);
      return;
    }
    const blob = new Blob([markdown], { type: 'text/markdown;charset=utf-8' });
    triggerDownload(blob, name);
  } catch (err) {
    console.error('[export] exportMarkdown failed:', err);
    throw err;
  }
}

/**
 * Render markdown to a self-contained HTML5 document and download as `.html`.
 * Embeds typography CSS and links KaTeX + highlight.js from CDN.
 */
export async function exportHtml(markdown: string, fileName: string, theme: string): Promise<void> {
  try {
    const name = sanitizeFileName(fileName, 'html');
    const bodyHtml = await Promise.resolve(marked.parse(markdown));
    const title = fileName.replace(/\.[^./\\]+$/, '');

    const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>${escapeHtml(title)}</title>
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.10/dist/katex.min.css" crossorigin="anonymous" />
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/highlight.js@11.9.0/styles/github.min.css" crossorigin="anonymous" />
  <style>
${EMBEDDED_CSS}
  </style>
</head>
<body data-theme="${escapeHtml(theme)}">
  <div class="markdown-body">
${bodyHtml}
  </div>
  <footer class="export-footer">Exported from Atlas</footer>
</body>
</html>`;

    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    triggerDownload(blob, name);
  } catch (err) {
    console.error('[export] exportHtml failed:', err);
    throw err;
  }
}

/** Escape HTML special characters for safe embedding in attributes/text. */
function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Capture a live DOM element by id and export it as a multi-page A4 PDF.
 * Uses html2canvas (scale 2) + jsPDF.
 */
export async function exportPdf(elementId: string, fileName: string): Promise<void> {
  try {
    const name = sanitizeFileName(fileName, 'pdf');
    const target = document.getElementById(elementId);
    if (!target) {
      throw new Error(`[export] exportPdf: element #${elementId} not found`);
    }

    const bgColor = getComputedStyle(target).backgroundColor;

    const canvas = await html2canvas(target, {
      scale: 2,
      useCORS: true,
      backgroundColor: bgColor || '#ffffff',
      logging: false,
    });

    // A4 dimensions in mm
    const A4_W = 210;
    const A4_H = 297;

    const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

    const imgData = canvas.toDataURL('image/png');
    const canvasWidthMm = A4_W;
    const canvasHeightMm = (canvas.height / canvas.width) * canvasWidthMm;

    let remainingHeight = canvasHeightMm;
    let yOffset = 0;

    while (remainingHeight > 0) {
      if (yOffset > 0) {
        pdf.addPage();
      }

      const sliceHeightMm = Math.min(A4_H, remainingHeight);
      const sliceRatio = sliceHeightMm / canvasHeightMm;
      const sliceHeightPx = Math.round(canvas.height * sliceRatio);
      const yOffsetPx = Math.round((yOffset / canvasHeightMm) * canvas.height);

      // Slice the canvas for this page
      const pageCanvas = document.createElement('canvas');
      pageCanvas.width = canvas.width;
      pageCanvas.height = sliceHeightPx;
      const ctx = pageCanvas.getContext('2d');
      if (ctx) {
        ctx.drawImage(canvas, 0, -yOffsetPx);
      }

      const pageImgData = pageCanvas.toDataURL('image/png');
      pdf.addImage(pageImgData, 'PNG', 0, 0, canvasWidthMm, sliceHeightMm);

      yOffset += sliceHeightMm;
      remainingHeight -= sliceHeightMm;
    }

    // Suppress unused variable warning — imgData used as fallback reference
    void imgData;

    const blob = pdf.output('blob');
    triggerDownload(blob, name);
  } catch (err) {
    console.error('[export] exportPdf failed:', err);
    throw err;
  }
}

// ---------------------------------------------------------------------------
// DOCX helpers
// ---------------------------------------------------------------------------

type DocxChild = Paragraph | Table;

/** Convert an array of inline marked tokens into docx TextRun / ExternalHyperlink children. */
function inlineTokensToRuns(
  tokens: Tokens.Generic[],
  opts: { bold?: boolean; italic?: boolean; code?: boolean; strike?: boolean } = {},
): (TextRun | ExternalHyperlink)[] {
  const runs: (TextRun | ExternalHyperlink)[] = [];

  for (const token of tokens) {
    switch (token.type) {
      case 'text': {
        const t = token as Tokens.Text;
        if (t.tokens && t.tokens.length > 0) {
          runs.push(...inlineTokensToRuns(t.tokens as Tokens.Generic[], opts));
        } else {
          runs.push(new TextRun({
            text: t.text,
            bold: opts.bold,
            italics: opts.italic,
            strike: opts.strike,
          }));
        }
        break;
      }
      case 'strong': {
        const t = token as Tokens.Strong;
        runs.push(...inlineTokensToRuns(t.tokens as Tokens.Generic[], { ...opts, bold: true }));
        break;
      }
      case 'em': {
        const t = token as Tokens.Em;
        runs.push(...inlineTokensToRuns(t.tokens as Tokens.Generic[], { ...opts, italic: true }));
        break;
      }
      case 'del': {
        const t = token as Tokens.Del;
        runs.push(...inlineTokensToRuns(t.tokens as Tokens.Generic[], { ...opts, strike: true }));
        break;
      }
      case 'codespan': {
        const t = token as Tokens.Codespan;
        runs.push(new TextRun({
          text: t.text,
          font: 'Courier New',
          bold: opts.bold,
          italics: opts.italic,
          strike: opts.strike,
          shading: { fill: 'EFEFEF' },
        }));
        break;
      }
      case 'link': {
        const t = token as Tokens.Link;
        const linkRuns = t.tokens
          ? inlineTokensToRuns(t.tokens as Tokens.Generic[], opts)
          : [new TextRun({ text: t.text, style: 'Hyperlink' })];
        runs.push(new ExternalHyperlink({
          link: t.href,
          children: linkRuns.map(r =>
            r instanceof TextRun ? new TextRun({ ...r, style: 'Hyperlink' }) : r,
          ),
        }));
        break;
      }
      case 'image': {
        const t = token as Tokens.Image;
        runs.push(new TextRun({ text: `[image: ${t.text || t.href}]`, italics: true }));
        break;
      }
      case 'br': {
        runs.push(new TextRun({ text: '', break: 1 }));
        break;
      }
      case 'html': {
        // Strip HTML tags for plain text fallback
        const t = token as Tokens.HTML;
        const stripped = t.text.replace(/<[^>]+>/g, '');
        if (stripped.trim()) {
          runs.push(new TextRun({ text: stripped }));
        }
        break;
      }
      default: {
        // Fallback: try to get raw text
        const raw = (token as Tokens.Generic).raw ?? '';
        if (raw.trim()) {
          runs.push(new TextRun({ text: raw }));
        }
        break;
      }
    }
  }

  return runs;
}

const HEADING_LEVEL_MAP: Record<number, (typeof HeadingLevel)[keyof typeof HeadingLevel]> = {
  1: HeadingLevel.HEADING_1,
  2: HeadingLevel.HEADING_2,
  3: HeadingLevel.HEADING_3,
  4: HeadingLevel.HEADING_4,
  5: HeadingLevel.HEADING_5,
  6: HeadingLevel.HEADING_6,
};

/** Recursively convert a list token's items into docx Paragraphs. */
function listItemsToParagraphs(
  items: Tokens.ListItem[],
  ordered: boolean,
  depth: number,
): Paragraph[] {
  const paragraphs: Paragraph[] = [];

  for (const item of items) {
    const runs: (TextRun | ExternalHyperlink)[] = [];

    for (const t of item.tokens as Tokens.Generic[]) {
      if (t.type === 'text') {
        const textToken = t as Tokens.Text;
        if (textToken.tokens && textToken.tokens.length > 0) {
          runs.push(...inlineTokensToRuns(textToken.tokens as Tokens.Generic[]));
        } else {
          runs.push(new TextRun({ text: textToken.text }));
        }
      } else if (t.type === 'list') {
        // Nested list — handled below after the parent paragraph
      } else {
        runs.push(...inlineTokensToRuns([t]));
      }
    }

    if (ordered) {
      paragraphs.push(new Paragraph({
        children: runs,
        numbering: { reference: 'default-numbering', level: depth },
      }));
    } else {
      paragraphs.push(new Paragraph({
        children: runs,
        bullet: { level: depth },
      }));
    }

    // Handle nested lists
    for (const t of item.tokens as Tokens.Generic[]) {
      if (t.type === 'list') {
        const nested = t as Tokens.List;
        paragraphs.push(...listItemsToParagraphs(nested.items, nested.ordered, depth + 1));
      }
    }
  }

  return paragraphs;
}

/** Convert a top-level marked Token to one or more docx block children. */
function tokenToDocxChildren(token: Tokens.Generic): DocxChild[] {
  switch (token.type) {
    case 'heading': {
      const t = token as Tokens.Heading;
      const level = HEADING_LEVEL_MAP[t.depth] ?? HeadingLevel.HEADING_6;
      const runs = inlineTokensToRuns(t.tokens as Tokens.Generic[]);
      return [new Paragraph({ heading: level, children: runs })];
    }

    case 'paragraph': {
      const t = token as Tokens.Paragraph;
      const runs = inlineTokensToRuns(t.tokens as Tokens.Generic[]);
      return [new Paragraph({ children: runs })];
    }

    case 'list': {
      const t = token as Tokens.List;
      return listItemsToParagraphs(t.items, t.ordered, 0);
    }

    case 'code': {
      const t = token as Tokens.Code;
      return [new Paragraph({
        children: [new TextRun({
          text: t.text,
          font: 'Courier New',
          size: 18,
          shading: { fill: 'F0F0F0' },
        })],
        spacing: { before: 120, after: 120 },
        indent: { left: 720 },
      })];
    }

    case 'blockquote': {
      const t = token as Tokens.Blockquote;
      const children: DocxChild[] = [];
      for (const inner of t.tokens as Tokens.Generic[]) {
        // Re-convert inner tokens with blockquote styling applied directly
        if (inner.type === 'paragraph') {
          const para = inner as Tokens.Paragraph;
          const runs = inlineTokensToRuns(para.tokens as Tokens.Generic[]);
          children.push(new Paragraph({
            children: runs,
            indent: { left: 720 },
            style: 'IntenseQuote',
          }));
        } else {
          children.push(...tokenToDocxChildren(inner));
        }
      }
      return children;
    }

    case 'hr': {
      return [new Paragraph({
        children: [],
        border: { bottom: { style: 'single', size: 6, color: 'AAAAAA', space: 1 } },
      })];
    }

    case 'table': {
      const t = token as Tokens.Table;
      const headerRow = new TableRow({
        children: t.header.map(cell =>
          new TableCell({
            children: [new Paragraph({
              children: inlineTokensToRuns(cell.tokens as Tokens.Generic[]),
            })],
          }),
        ),
        tableHeader: true,
      });

      const bodyRows = t.rows.map(row =>
        new TableRow({
          children: row.map(cell =>
            new TableCell({
              children: [new Paragraph({
                children: inlineTokensToRuns(cell.tokens as Tokens.Generic[]),
              })],
            }),
          ),
        }),
      );

      return [new Table({ rows: [headerRow, ...bodyRows] })];
    }

    case 'html': {
      const t = token as Tokens.HTML;
      const stripped = t.text.replace(/<[^>]+>/g, '').trim();
      if (!stripped) return [];
      return [new Paragraph({ children: [new TextRun({ text: stripped })] })];
    }

    case 'space': {
      return [new Paragraph({ children: [] })];
    }

    default: {
      const raw = token.raw?.trim();
      if (!raw) return [];
      return [new Paragraph({ children: [new TextRun({ text: raw })] })];
    }
  }
}

/**
 * Parse markdown and export as a `.docx` file with proper heading, list,
 * table, code block, blockquote, and hyperlink formatting.
 */
export async function exportDocx(markdown: string, fileName: string): Promise<void> {
  try {
    const name = sanitizeFileName(fileName, 'docx');
    const tokens = marked.lexer(markdown) as Tokens.Generic[];

    const children: DocxChild[] = [];
    for (const token of tokens) {
      children.push(...tokenToDocxChildren(token));
    }

    const docOptions: IPropertiesOptions = {
      numbering: {
        config: [
          {
            reference: 'default-numbering',
            levels: [
              { level: 0, format: 'decimal', text: '%1.', alignment: 'start' },
              { level: 1, format: 'decimal', text: '%2.', alignment: 'start' },
              { level: 2, format: 'decimal', text: '%3.', alignment: 'start' },
            ],
          },
        ],
      },
      sections: [
        {
          children,
        },
      ],
    };

    const doc = new Document(docOptions);
    const blob = await Packer.toBlob(doc);
    triggerDownload(blob, name);
  } catch (err) {
    console.error('[export] exportDocx failed:', err);
    throw err;
  }
}
