/**
 * WCAG contrast ratio for `--text-tertiary` against both background tokens,
 * in all 5 themes (UX-10 / X3).
 *
 * `--text-tertiary` backs ~26 usages across the app (icons, secondary
 * labels, placeholder-ish text) and previously fell as low as 2.89:1 against
 * `--bg-secondary` in 4 of 5 themes — well under WCAG 2.1 AA's 4.5:1 minimum
 * for normal text. This test reads the *real* colors straight out of
 * `index.css` (not a hand-copied duplicate) and computes the standard
 * relative-luminance contrast ratio itself, so a future edit to either the
 * text or background token that regresses contrast in any theme fails here
 * immediately instead of shipping a silent accessibility regression.
 */
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';

const INDEX_CSS_PATH = path.resolve(process.cwd(), 'src/index.css');
const indexCss = readFileSync(INDEX_CSS_PATH, 'utf-8');

const WCAG_AA_NORMAL_TEXT_MIN_RATIO = 4.5;

const THEME_SELECTORS = {
  light: ':root',
  dark: '[data-theme="dark"]',
  sepia: '[data-theme="sepia"]',
  nord: '[data-theme="nord"]',
  dracula: '[data-theme="dracula"]',
} as const;

type ThemeId = keyof typeof THEME_SELECTORS;

function extractBlock(selector: string): string {
  const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(`${escaped}\\s*\\{([\\s\\S]*?)\\n\\}`, 'm');
  const match = re.exec(indexCss);
  if (!match) throw new Error(`index.css: could not find a "${selector}" rule block`);
  return match[1]!;
}

function extractHexVar(block: string, name: string): string {
  const re = new RegExp(`--${name}:\\s*(#[0-9a-fA-F]{6});`);
  const match = re.exec(block);
  if (!match) throw new Error(`theme block is missing a hex --${name}`);
  return match[1]!;
}

function hexToRgb(hex: string): readonly [number, number, number] {
  return [parseInt(hex.slice(1, 3), 16), parseInt(hex.slice(3, 5), 16), parseInt(hex.slice(5, 7), 16)];
}

/** WCAG 2.1 relative luminance (https://www.w3.org/TR/WCAG21/#dfn-relative-luminance). */
function relativeLuminance([r, g, b]: readonly [number, number, number]): number {
  const channel = (c: number) => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  };
  return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b);
}

/** WCAG 2.1 contrast ratio (https://www.w3.org/TR/WCAG21/#dfn-contrast-ratio). */
function contrastRatio(hexA: string, hexB: string): number {
  const lumA = relativeLuminance(hexToRgb(hexA));
  const lumB = relativeLuminance(hexToRgb(hexB));
  const lighter = Math.max(lumA, lumB);
  const darker = Math.min(lumA, lumB);
  return (lighter + 0.05) / (darker + 0.05);
}

describe('--text-tertiary WCAG AA contrast (UX-10)', () => {
  const themeIds = Object.keys(THEME_SELECTORS) as ThemeId[];

  it.each(themeIds)('meets >=4.5:1 against --bg-primary in the "%s" theme', (id) => {
    const block = extractBlock(THEME_SELECTORS[id]);
    const ratio = contrastRatio(extractHexVar(block, 'text-tertiary'), extractHexVar(block, 'bg-primary'));
    expect(ratio).toBeGreaterThanOrEqual(WCAG_AA_NORMAL_TEXT_MIN_RATIO);
  });

  it.each(themeIds)('meets >=4.5:1 against --bg-secondary in the "%s" theme', (id) => {
    const block = extractBlock(THEME_SELECTORS[id]);
    const ratio = contrastRatio(extractHexVar(block, 'text-tertiary'), extractHexVar(block, 'bg-secondary'));
    expect(ratio).toBeGreaterThanOrEqual(WCAG_AA_NORMAL_TEXT_MIN_RATIO);
  });

  it('covers all 5 themes the app ships', () => {
    expect(themeIds.sort()).toEqual(['dark', 'dracula', 'light', 'nord', 'sepia']);
  });
});
