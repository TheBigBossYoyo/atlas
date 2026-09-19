/**
 * WCAG AA contrast for the token pairs Atlas actually renders text, controls,
 * links, error/danger banners, buttons and the search-match highlight with,
 * in all 5 themes (UX-10, extended by A11Y-2).
 *
 * This test reads the *real* colors straight out of `index.css` (not a
 * hand-copied duplicate) and computes the standard relative-luminance
 * contrast ratio itself, so a future edit to any of these tokens that
 * regresses contrast in any theme fails here immediately instead of shipping
 * a silent accessibility regression.
 *
 * Scope (A11Y-2 pass): --text-primary/--text-secondary/--text-tertiary
 * against every background token they're actually painted on
 * (--bg-primary/--bg-secondary/--bg-tertiary/--bg-elevated), --accent as a
 * link/text color and as a button background with its paired button-text
 * color, --hl-keyword as the reused "danger" text color, and the
 * `mark.search-highlight`/`--active` search-match highlight (alpha-composited
 * onto --bg-primary, since it renders over ordinary document text).
 *
 * Deliberately out of scope: --border-primary and other decorative
 * boundaries (WCAG 1.4.11 exempts purely decorative, non-essential
 * boundaries — every theme's borders are intentionally subtle) and
 * `:disabled` controls (WCAG 1.4.3/1.4.11 exempt inactive UI components
 * outright; Atlas dims them with `opacity` rather than a separate token).
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
type Rgb = readonly [number, number, number];

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

/** Like extractHexVar, but returns null instead of throwing when the theme doesn't define it. */
function extractHexVarOptional(block: string, name: string): string | null {
  const re = new RegExp(`--${name}:\\s*(#[0-9a-fA-F]{6});`);
  return re.exec(block)?.[1] ?? null;
}

/** Extracts an rgba(...) declaration and returns its channels plus alpha. */
function extractRgbaVar(block: string, name: string): readonly [number, number, number, number] {
  const re = new RegExp(`--${name}:\\s*rgba\\((\\d+),\\s*(\\d+),\\s*(\\d+),\\s*([\\d.]+)\\);`);
  const match = re.exec(block);
  if (!match) throw new Error(`theme block is missing an rgba --${name}`);
  return [Number(match[1]), Number(match[2]), Number(match[3]), Number(match[4])];
}

function hexToRgb(hex: string): Rgb {
  return [parseInt(hex.slice(1, 3), 16), parseInt(hex.slice(3, 5), 16), parseInt(hex.slice(5, 7), 16)];
}

/** WCAG 2.1 relative luminance (https://www.w3.org/TR/WCAG21/#dfn-relative-luminance). */
function relativeLuminance([r, g, b]: Rgb): number {
  const channel = (c: number) => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  };
  return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b);
}

/** WCAG 2.1 contrast ratio (https://www.w3.org/TR/WCAG21/#dfn-contrast-ratio). */
function contrastRatioRgb(a: Rgb, b: Rgb): number {
  const lumA = relativeLuminance(a);
  const lumB = relativeLuminance(b);
  const lighter = Math.max(lumA, lumB);
  const darker = Math.min(lumA, lumB);
  return (lighter + 0.05) / (darker + 0.05);
}

function contrastRatio(hexA: string, hexB: string): number {
  return contrastRatioRgb(hexToRgb(hexA), hexToRgb(hexB));
}

/** Alpha-composites an rgba foreground over an opaque hex background. */
function compositeOver([r, g, b, a]: readonly [number, number, number, number], bgHex: string): Rgb {
  const bg = hexToRgb(bgHex);
  return [r * a + bg[0] * (1 - a), g * a + bg[1] * (1 - a), b * a + bg[2] * (1 - a)];
}

const themeIds = Object.keys(THEME_SELECTORS) as ThemeId[];

const BG_TOKENS = ['bg-primary', 'bg-secondary', 'bg-tertiary', 'bg-elevated'] as const;

describe('--text-tertiary WCAG AA contrast (UX-10)', () => {
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

describe('A11Y-2: text tokens vs every background they are actually painted on', () => {
  const textTokens = ['text-primary', 'text-secondary', 'text-tertiary'] as const;

  for (const textToken of textTokens) {
    for (const bgToken of BG_TOKENS) {
      it.each(themeIds)(`--${textToken} meets >=4.5:1 against --${bgToken} in the "%s" theme`, (id) => {
        const block = extractBlock(THEME_SELECTORS[id]);
        const ratio = contrastRatio(extractHexVar(block, textToken), extractHexVar(block, bgToken));
        expect(ratio).toBeGreaterThanOrEqual(WCAG_AA_NORMAL_TEXT_MIN_RATIO);
      });
    }
  }
});

describe('A11Y-2: --accent as a link/text color (e.g. markdown links, secondary-button hover)', () => {
  for (const bgToken of BG_TOKENS) {
    it.each(themeIds)(`--accent meets >=4.5:1 against --${bgToken} in the "%s" theme`, (id) => {
      const block = extractBlock(THEME_SELECTORS[id]);
      const ratio = contrastRatio(extractHexVar(block, 'accent'), extractHexVar(block, bgToken));
      expect(ratio).toBeGreaterThanOrEqual(WCAG_AA_NORMAL_TEXT_MIN_RATIO);
    });
  }
});

describe('A11Y-2: --bg-primary as button text on an --accent / --accent-hover background', () => {
  // .welcome__btn--primary, .confirm-dialog__btn--primary, .viewer-fallback__btn
  // and .code-viewer__button--primary/.slide-deck__save-button all use
  // `background: var(--accent); color: var(--bg-primary);` (previously a
  // hardcoded #ffffff, which failed AA against dark/nord/dracula's lighter
  // --accent) and swap to --accent-hover on :hover with the same text color.
  it.each(themeIds)('--bg-primary meets >=4.5:1 against --accent in the "%s" theme', (id) => {
    const block = extractBlock(THEME_SELECTORS[id]);
    const ratio = contrastRatio(extractHexVar(block, 'bg-primary'), extractHexVar(block, 'accent'));
    expect(ratio).toBeGreaterThanOrEqual(WCAG_AA_NORMAL_TEXT_MIN_RATIO);
  });

  it.each(themeIds)('--bg-primary meets >=4.5:1 against --accent-hover in the "%s" theme', (id) => {
    const block = extractBlock(THEME_SELECTORS[id]);
    const ratio = contrastRatio(extractHexVar(block, 'bg-primary'), extractHexVar(block, 'accent-hover'));
    expect(ratio).toBeGreaterThanOrEqual(WCAG_AA_NORMAL_TEXT_MIN_RATIO);
  });
});

describe('A11Y-2: --hl-keyword as the reused "danger" text color (error banners, discard button)', () => {
  for (const bgToken of BG_TOKENS) {
    it.each(themeIds)(`--hl-keyword meets >=4.5:1 against --${bgToken} in the "%s" theme`, (id) => {
      const block = extractBlock(THEME_SELECTORS[id]);
      const ratio = contrastRatio(extractHexVar(block, 'hl-keyword'), extractHexVar(block, bgToken));
      expect(ratio).toBeGreaterThanOrEqual(WCAG_AA_NORMAL_TEXT_MIN_RATIO);
    });
  }
});

describe('A11Y-2: mark.search-highlight / --active vs the text it covers', () => {
  // The mark itself is painted over ordinary document text (--text-primary,
  // via `color: inherit`, or the theme's `--search-highlight[-active]-text`
  // override when the alpha-composited highlight is too light for that).
  // Light/sepia's highlight tokens are solid hex; dark/nord/dracula's are
  // translucent and get alpha-composited onto --bg-primary first, matching
  // how they actually render over the main content pane.
  function resolvedMarkTextColor(block: string, variant: 'base' | 'active'): string {
    const overrideName = variant === 'active' ? 'search-highlight-active-text' : 'search-highlight-text';
    const override = extractHexVarOptional(block, overrideName);
    if (override) return override;
    if (variant === 'active') {
      const baseOverride = extractHexVarOptional(block, 'search-highlight-text');
      if (baseOverride) return baseOverride;
    }
    return extractHexVar(block, 'text-primary');
  }

  function resolvedMarkBackground(block: string, variant: 'base' | 'active'): Rgb {
    const name = variant === 'active' ? 'search-highlight-active' : 'search-highlight';
    const hex = extractHexVarOptional(block, name);
    if (hex) return hexToRgb(hex);
    const rgba = extractRgbaVar(block, name);
    return compositeOver(rgba, extractHexVar(block, 'bg-primary'));
  }

  for (const variant of ['base', 'active'] as const) {
    it.each(themeIds)(`the ${variant} mark's text meets >=4.5:1 against its highlight in the "%s" theme`, (id) => {
      const block = extractBlock(THEME_SELECTORS[id]);
      const textHex = resolvedMarkTextColor(block, variant);
      const bgRgb = resolvedMarkBackground(block, variant);
      const ratio = contrastRatioRgb(hexToRgb(textHex), bgRgb);
      expect(ratio).toBeGreaterThanOrEqual(WCAG_AA_NORMAL_TEXT_MIN_RATIO);
    });
  }
});
