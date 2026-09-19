/**
 * markdownExportCss.ts <-> src/index.css crossref (UX-20).
 *
 * `buildMarkdownExportCss()`'s heading rules are a hand-typed subset of
 * `index.css`'s live `.markdown-body h1..h6` rules (the exported document is
 * a standalone file and can't reference the app's real stylesheet). This had
 * already silently drifted once — the export copy's h4/h5/h6 sizes disagreed
 * with the live preview until this test was added. Reads `index.css` off
 * disk and diffs each heading's real `font-size` against the generated
 * export CSS's, so a future edit to one and not the other fails here instead
 * of quietly drifting again.
 */
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { buildMarkdownExportCss } from '../markdownExportCss';

const INDEX_CSS_PATH = path.resolve(process.cwd(), 'src/index.css');
const indexCss = readFileSync(INDEX_CSS_PATH, 'utf-8');
const exportCss = buildMarkdownExportCss();

const HEADINGS = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'] as const;

/** Extracts `font-size: <value>` from the first `.markdown-body <tag> { ... }`
 * rule found in `css` (heading rules in both files are single-line, possibly
 * with sibling declarations on the same line). */
function extractHeadingFontSize(css: string, tag: string): string {
  const re = new RegExp(`\\.markdown-body ${tag}\\s*\\{[^}]*font-size:\\s*([^;]+);`);
  const match = re.exec(css);
  if (!match) throw new Error(`could not find ".markdown-body ${tag}" font-size rule`);
  return match[1]!.trim();
}

describe('markdownExportCss <-> index.css heading crossref', () => {
  it.each(HEADINGS)('%s font-size matches the live preview', (tag) => {
    expect(extractHeadingFontSize(exportCss, tag)).toBe(extractHeadingFontSize(indexCss, tag));
  });
});
