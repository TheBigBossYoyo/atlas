/**
 * exportThemeTokens.ts <-> src/index.css cross-reference (UX-20).
 *
 * `exportThemeTokens.ts` is a second, hand-typed copy of the app's five theme
 * colors (index.css's real source of truth is CSS custom properties, which an
 * exported *standalone* HTML document cannot reference at all — it has to
 * carry its own literal color values). Nothing enforces these two places stay
 * in sync at the type level, so this test reads `index.css` straight off disk
 * and diffs its real `--bg-primary`/`--text-primary`/`--code-bg`/
 * `--border-primary`/`--accent` values, per theme, against
 * `EXPORT_THEME_TOKENS`. A future theme edit made in one place and not
 * mirrored in the other now fails loudly here instead of quietly drifting
 * (the UX-20 finding: this exact drift had already happened once before).
 */
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { EXPORT_THEME_ORDER, EXPORT_THEME_TOKENS, type ExportThemeId } from '../exportThemeTokens';

const INDEX_CSS_PATH = path.resolve(process.cwd(), 'src/index.css');
const indexCss = readFileSync(INDEX_CSS_PATH, 'utf-8');

/** Selector index.css actually uses for each theme (light lives on bare `:root`). */
function selectorFor(id: ExportThemeId): string {
  return id === 'light' ? ':root' : `[data-theme="${id}"]`;
}

/** Extracts the `{ ... }` block body for `selector` (the *first* occurrence —
 * every theme selector in index.css appears exactly once as a top-level
 * rule). Failing to find it is itself a signal the CSS structure changed. */
function extractBlock(selector: string): string {
  const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(`${escaped}\\s*\\{([\\s\\S]*?)\\n\\}`, 'm');
  const match = re.exec(indexCss);
  if (!match) throw new Error(`index.css: could not find a "${selector}" rule block`);
  return match[1]!;
}

function extractVar(block: string, name: string): string {
  const re = new RegExp(`--${name}:\\s*([^;]+);`);
  const match = re.exec(block);
  if (!match) throw new Error(`index.css block is missing --${name}`);
  return match[1]!.trim();
}

describe('exportThemeTokens <-> index.css crossref', () => {
  it.each(EXPORT_THEME_ORDER)('the "%s" export tokens match index.css\'s real theme variables', (id) => {
    const block = extractBlock(selectorFor(id));
    const tokens = EXPORT_THEME_TOKENS[id];

    expect(tokens.bg).toBe(extractVar(block, 'bg-primary'));
    expect(tokens.text).toBe(extractVar(block, 'text-primary'));
    expect(tokens.codeBg).toBe(extractVar(block, 'code-bg'));
    expect(tokens.border).toBe(extractVar(block, 'border-primary'));
    expect(tokens.accent).toBe(extractVar(block, 'accent'));
  });

  it('covers every theme index.css actually defines (light, dark, sepia, nord, dracula)', () => {
    expect([...EXPORT_THEME_ORDER].sort()).toEqual(['dark', 'dracula', 'light', 'nord', 'sepia']);
  });
});
