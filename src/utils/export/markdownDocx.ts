/**
 * Markdown → DOCX export (X2 / UX-05, UX-06, wave1 export follow-ups).
 *
 * Parses markdown with `marked`'s token lexer (not the live remark/rehype
 * preview pipeline — Word has no equivalent renderer to reuse) and walks the
 * resulting tokens into `docx` document children. Three fidelity fixes over
 * the pre-X2 version, all covered by the characterization suite:
 *
 *   1. Task-list checkboxes (`item.task`/`item.checked`) are rendered as a
 *      real ☑/☐ glyph instead of leaking the raw `[x] `/`[ ] ` markdown
 *      syntax as literal text (UX-06).
 *   2. Hyperlink text is no longer dropped: applying the `Hyperlink` style by
 *      spreading an already-constructed `TextRun` instance (`{ ...run, style:
 *      'Hyperlink' }`) doesn't work — `TextRun`'s visible text isn't a plain
 *      `{ text }` own-property the spread can pick up, so the run rendered
 *      empty. The style is now applied at the *original* construction site
 *      instead, by threading it through as a run option (wave1 follow-up).
 *   3. Math (`$..$` / `$$..$$`, this parser has no `remark-math` equivalent
 *      so these are detected with a small regex scan of leaf text tokens) and
 *      ```mermaid fences are rasterized to PNG (see `docxMedia.ts`) and
 *      embedded as images instead of exporting as raw source text.
 */
import { marked } from 'marked';
import type { Tokens } from 'marked';
import {
  AlignmentType,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  ImageRun,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
} from 'docx';
import type { IPropertiesOptions } from 'docx';

import { sanitizeFileName, saveBinaryOutput, type SaveFilter } from './download';
import { renderMathImage, renderMermaidImage, scaleToMaxWidth } from './docxMedia';
import { toFriendlyError } from '../friendlyLibraryError';

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const DOCX_FILTERS: readonly SaveFilter[] = [{ name: 'Word Document', extensions: ['docx'] }];

type DocxRun = TextRun | ExternalHyperlink | ImageRun;
type DocxChild = Paragraph | Table;

interface RunOpts {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly strike?: boolean;
  /** Named docx paragraph/run style to apply (e.g. `'Hyperlink'`). */
  readonly style?: string;
}

// ---------------------------------------------------------------------------
// Inline math detection (no remark-math equivalent on this parser)
// ---------------------------------------------------------------------------

type TextSegment = { readonly kind: 'text'; readonly value: string } | { readonly kind: 'math'; readonly value: string };

const INLINE_MATH_RE = /\$([^$\n]+?)\$/g;

/**
 * Splits `text` into plain-text and inline-math segments. Guards against the
 * common "$5 and $10" currency false-positive the same way `remark-math`
 * does: a `$...$` span only counts as math when its content has no leading
 * or trailing whitespace.
 */
function splitInlineMath(text: string): readonly TextSegment[] {
  const segments: TextSegment[] = [];
  let lastIndex = 0;
  for (const match of text.matchAll(INLINE_MATH_RE)) {
    const inner = match[1] ?? '';
    if (inner.length === 0 || inner.trim() !== inner) continue;
    const index = match.index ?? 0;
    if (index > lastIndex) segments.push({ kind: 'text', value: text.slice(lastIndex, index) });
    segments.push({ kind: 'math', value: inner });
    lastIndex = index + match[0].length;
  }
  if (lastIndex < text.length) segments.push({ kind: 'text', value: text.slice(lastIndex) });
  return segments.length > 0 ? segments : [{ kind: 'text', value: text }];
}

/** Converts a leaf text string into runs, rasterizing any inline `$..$` math it contains. */
async function textToRuns(text: string, opts: RunOpts): Promise<DocxRun[]> {
  const segments = splitInlineMath(text);
  const runs: DocxRun[] = [];

  for (const segment of segments) {
    if (segment.kind === 'text') {
      if (segment.value.length === 0) continue;
      runs.push(new TextRun({ text: segment.value, bold: opts.bold, italics: opts.italic, strike: opts.strike, style: opts.style }));
      continue;
    }

    const image = await renderMathImage(segment.value, false);
    if (image) {
      const { width, height } = scaleToMaxWidth(image.width, image.height, 300);
      runs.push(new ImageRun({ type: 'png', data: image.bytes, transformation: { width, height } }));
    } else {
      // Rasterization failed (e.g. invalid LaTeX) — fall back to the literal
      // source rather than silently dropping the equation.
      runs.push(new TextRun({ text: `$${segment.value}$`, italics: true, bold: opts.bold, style: opts.style }));
    }
  }

  return runs.length > 0 ? runs : [new TextRun({ text, bold: opts.bold, italics: opts.italic, strike: opts.strike, style: opts.style })];
}

// ---------------------------------------------------------------------------
// Inline token walking
// ---------------------------------------------------------------------------

async function inlineTokensToRuns(tokens: readonly Tokens.Generic[], opts: RunOpts = {}): Promise<DocxRun[]> {
  const runs: DocxRun[] = [];

  for (const token of tokens) {
    switch (token.type) {
      case 'text': {
        const t = token as Tokens.Text;
        if (t.tokens && t.tokens.length > 0) {
          runs.push(...(await inlineTokensToRuns(t.tokens as Tokens.Generic[], opts)));
        } else {
          runs.push(...(await textToRuns(t.text, opts)));
        }
        break;
      }
      case 'strong': {
        const t = token as Tokens.Strong;
        runs.push(...(await inlineTokensToRuns(t.tokens as Tokens.Generic[], { ...opts, bold: true })));
        break;
      }
      case 'em': {
        const t = token as Tokens.Em;
        runs.push(...(await inlineTokensToRuns(t.tokens as Tokens.Generic[], { ...opts, italic: true })));
        break;
      }
      case 'del': {
        const t = token as Tokens.Del;
        runs.push(...(await inlineTokensToRuns(t.tokens as Tokens.Generic[], { ...opts, strike: true })));
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
          style: opts.style,
          shading: { fill: 'EFEFEF' },
        }));
        break;
      }
      case 'link': {
        const t = token as Tokens.Link;
        // Apply the 'Hyperlink' style at construction time (not by spreading
        // an already-built TextRun afterward — see file header, fix #2).
        const linkRuns = t.tokens
          ? await inlineTokensToRuns(t.tokens as Tokens.Generic[], { ...opts, style: 'Hyperlink' })
          : [new TextRun({ text: t.text, style: 'Hyperlink' })];
        runs.push(new ExternalHyperlink({ link: t.href, children: linkRuns }));
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
        const t = token as Tokens.HTML;
        const stripped = t.text.replace(/<[^>]+>/g, '');
        if (stripped.trim()) runs.push(new TextRun({ text: stripped }));
        break;
      }
      case 'checkbox': {
        // Handled explicitly by listItemsToParagraphs (UX-06) — ignored here
        // so the raw "[x] "/"[ ] " syntax never leaks through the generic
        // fallback below if a checkbox token ever reaches this level.
        break;
      }
      default: {
        const raw = (token as Tokens.Generic).raw ?? '';
        if (raw.trim()) runs.push(new TextRun({ text: raw }));
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

/** Recursively converts a list token's items into docx Paragraphs, preserving task-list checkbox state (UX-06). */
async function listItemsToParagraphs(items: readonly Tokens.ListItem[], ordered: boolean, depth: number): Promise<Paragraph[]> {
  const paragraphs: Paragraph[] = [];

  for (const item of items) {
    const runs: DocxRun[] = [];

    if (item.task) {
      runs.push(new TextRun({ text: item.checked ? '☑ ' : '☐ ' }));
    }

    for (const t of item.tokens as Tokens.Generic[]) {
      if (t.type === 'checkbox') continue; // consumed by item.task/item.checked above
      if (t.type === 'text') {
        const textToken = t as Tokens.Text;
        if (textToken.tokens && textToken.tokens.length > 0) {
          runs.push(...(await inlineTokensToRuns(textToken.tokens as Tokens.Generic[])));
        } else {
          runs.push(...(await textToRuns(textToken.text, {})));
        }
      } else if (t.type === 'list') {
        // Nested list — handled below, after the parent paragraph.
      } else {
        runs.push(...(await inlineTokensToRuns([t])));
      }
    }

    paragraphs.push(
      ordered
        ? new Paragraph({ children: runs, numbering: { reference: 'default-numbering', level: depth } })
        : new Paragraph({ children: runs, bullet: { level: depth } }),
    );

    for (const t of item.tokens as Tokens.Generic[]) {
      if (t.type === 'list') {
        const nested = t as Tokens.List;
        paragraphs.push(...(await listItemsToParagraphs(nested.items, nested.ordered, depth + 1)));
      }
    }
  }

  return paragraphs;
}

/** Whole-paragraph display math: a paragraph whose entire (trimmed) text is one `$$...$$` block. */
const BLOCK_MATH_RE = /^\$\$([\s\S]+)\$\$$/;

async function paragraphChildren(t: Tokens.Paragraph): Promise<DocxChild[]> {
  const blockMathMatch = BLOCK_MATH_RE.exec(t.text.trim());
  if (blockMathMatch) {
    const image = await renderMathImage(blockMathMatch[1]!.trim(), true);
    if (image) {
      const { width, height } = scaleToMaxWidth(image.width, image.height);
      return [new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ type: 'png', data: image.bytes, transformation: { width, height } })] })];
    }
    // Rasterization failed — fall through to plain-text rendering below.
  }

  const runs = await inlineTokensToRuns(t.tokens as Tokens.Generic[]);
  return [new Paragraph({ children: runs })];
}

/** Converts a top-level marked Token into one or more docx block children. */
async function tokenToDocxChildren(token: Tokens.Generic): Promise<DocxChild[]> {
  switch (token.type) {
    case 'heading': {
      const t = token as Tokens.Heading;
      const level = HEADING_LEVEL_MAP[t.depth] ?? HeadingLevel.HEADING_6;
      const runs = await inlineTokensToRuns(t.tokens as Tokens.Generic[]);
      return [new Paragraph({ heading: level, children: runs })];
    }

    case 'paragraph': {
      return paragraphChildren(token as Tokens.Paragraph);
    }

    case 'list': {
      const t = token as Tokens.List;
      return listItemsToParagraphs(t.items, t.ordered, 0);
    }

    case 'code': {
      const t = token as Tokens.Code;
      if (t.lang === 'mermaid') {
        const image = await renderMermaidImage(t.text);
        if (image) {
          const { width, height } = scaleToMaxWidth(image.width, image.height);
          return [new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ type: 'png', data: image.bytes, transformation: { width, height } })] })];
        }
        // Rasterization failed — fall through to the plain code block below,
        // so a bad Mermaid diagram degrades to visible source, not silence.
      }
      return [new Paragraph({
        children: [new TextRun({ text: t.text, font: 'Courier New', size: 18, shading: { fill: 'F0F0F0' } })],
        spacing: { before: 120, after: 120 },
        indent: { left: 720 },
      })];
    }

    case 'blockquote': {
      const t = token as Tokens.Blockquote;
      const children: DocxChild[] = [];
      for (const inner of t.tokens as Tokens.Generic[]) {
        if (inner.type === 'paragraph') {
          const para = inner as Tokens.Paragraph;
          const runs = await inlineTokensToRuns(para.tokens as Tokens.Generic[]);
          children.push(new Paragraph({ children: runs, indent: { left: 720 }, style: 'IntenseQuote' }));
        } else {
          children.push(...(await tokenToDocxChildren(inner)));
        }
      }
      return children;
    }

    case 'hr': {
      return [new Paragraph({ children: [], border: { bottom: { style: 'single', size: 6, color: 'AAAAAA', space: 1 } } })];
    }

    case 'table': {
      const t = token as Tokens.Table;
      const headerRow = new TableRow({
        children: await Promise.all(t.header.map(async cell =>
          new TableCell({ children: [new Paragraph({ children: await inlineTokensToRuns(cell.tokens as Tokens.Generic[]) })] }),
        )),
        tableHeader: true,
      });
      const bodyRows = await Promise.all(t.rows.map(async row =>
        new TableRow({
          children: await Promise.all(row.map(async cell =>
            new TableCell({ children: [new Paragraph({ children: await inlineTokensToRuns(cell.tokens as Tokens.Generic[]) })] }),
          )),
        }),
      ));
      return [new Table({ rows: [headerRow, ...bodyRows] })];
    }

    case 'html': {
      const t = token as Tokens.HTML;
      const stripped = t.text.replace(/<[^>]+>/g, '').trim();
      return stripped ? [new Paragraph({ children: [new TextRun({ text: stripped })] })] : [];
    }

    case 'space': {
      return [new Paragraph({ children: [] })];
    }

    default: {
      const raw = token.raw?.trim();
      return raw ? [new Paragraph({ children: [new TextRun({ text: raw })] })] : [];
    }
  }
}

/**
 * Parses markdown and exports it as a `.docx` file with heading, list, table,
 * code block, blockquote, hyperlink, task-list checkbox, and rendered
 * math/Mermaid formatting.
 */
export async function exportDocx(markdown: string, fileName: string): Promise<void> {
  try {
    const name = sanitizeFileName(fileName, 'docx');
    const tokens = marked.lexer(markdown) as Tokens.Generic[];

    const children: DocxChild[] = [];
    for (const token of tokens) {
      children.push(...(await tokenToDocxChildren(token)));
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
      sections: [{ children }],
    };

    const doc = new Document(docOptions);
    const blob = await Packer.toBlob(doc);
    const bytes = new Uint8Array(await blob.arrayBuffer());
    await saveBinaryOutput(bytes, name, DOCX_MIME, DOCX_FILTERS);
  } catch (err) {
    console.error('[export] exportDocx failed:', err);
    throw toFriendlyError(err, 'DOCX export failed');
  }
}
