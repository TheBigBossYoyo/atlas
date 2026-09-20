/**
 * UX-20 — canonical per-theme color tokens used to generate exported HTML's
 * embedded CSS (`markdownHtml.ts`). This used to be a second, hand-typed copy
 * of the same five themes' colors living only in `export.ts`'s template
 * string, silently drifting from `src/index.css`'s `:root`/`[data-theme]`
 * blocks over time (the two were already out of sync before X2).
 *
 * This is now the single place those five *export* theme colors are typed,
 * and `exportThemeTokens.crossref.test.ts` diffs it against `index.css`'s
 * real `--bg-primary`/`--text-primary`/`--code-bg`/`--border-primary`/
 * `--accent` values on every test run, so a future theme edit in one place
 * that isn't mirrored here fails loudly instead of drifting silently again.
 */

export type ExportThemeId = 'light' | 'dark' | 'sepia' | 'nord' | 'dracula';

export interface ExportThemeTokens {
  readonly bg: string;
  readonly text: string;
  readonly codeBg: string;
  readonly border: string;
  readonly accent: string;
}

export const EXPORT_THEME_TOKENS: Record<ExportThemeId, ExportThemeTokens> = {
  light: { bg: '#ffffff', text: '#1f2328', codeBg: '#eff1f3', border: '#d0d7de', accent: '#0967d7' },
  dark: { bg: '#0d1117', text: '#e6edf3', codeBg: '#262c36', border: '#30363d', accent: '#58a6ff' },
  sepia: { bg: '#f4ecd8', text: '#5b4636', codeBg: '#ebe2c8', border: '#d4c9a8', accent: '#815428' },
  nord: { bg: '#2e3440', text: '#eceff4', codeBg: '#3b4252', border: '#434c5e', accent: '#90c5d3' },
  dracula: { bg: '#282a36', text: '#f8f8f2', codeBg: '#44475a', border: '#44475a', accent: '#caa8fa' },
};

export const EXPORT_THEME_ORDER: readonly ExportThemeId[] = ['light', 'dark', 'sepia', 'nord', 'dracula'];
