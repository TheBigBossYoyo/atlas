import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';

/**
 * Lint-style guard, not a type check: every shell/viewer-chrome component
 * converted to `useTranslate()` in this i18n wave must not have a literal
 * `title="..."` / `aria-label="..."` / `placeholder="..."` left over — after
 * conversion, every one of these attributes is `{t('some.key', ...)}`, an
 * expression, never a bare string literal. A regression here (someone adding
 * a new button and forgetting `t()`) fails this test immediately instead of
 * silently shipping an English string into the French UI.
 *
 * Scoped to the files this wave actually converted (see the i18n section of
 * the plan) — not the whole repo, since large surfaces like `DocxViewer.tsx`
 * and the slide editor are explicitly out of scope for this wave (see the
 * final report) and still have hard-coded English throughout.
 */
const CONVERTED_FILES = [
  'src/components/DraftRecoveryBanner.tsx',
  'src/components/DropZone.tsx',
  'src/components/EmptyFileNotice.tsx',
  'src/components/ExportMenu.tsx',
  'src/components/FileStatusBanner.tsx',
  'src/components/LanguageMenu.tsx',
  'src/components/NewDocumentMenu.tsx',
  'src/components/SearchOverlay.tsx',
  'src/components/ShortcutsModal.tsx',
  'src/components/Sidebar.tsx',
  'src/components/StatusBar.tsx',
  'src/components/TabBar.tsx',
  'src/components/ThemeMenu.tsx',
  'src/components/Toast.tsx',
  'src/components/Toolbar.tsx',
  'src/components/UnsavedChangesDialog.tsx',
  'src/components/WelcomeScreen.tsx',
  'src/viewers/CodeViewer.tsx',
  'src/viewers/CsvViewer.tsx',
  'src/viewers/SpreadsheetViewer.tsx',
  'src/viewers/pdf/PdfToolbar.tsx',
  'src/viewers/shared/LegacyFormatBanner.tsx',
  'src/viewers/spreadsheet/SpreadsheetEditToolbar.tsx',
];

// A literal (non-expression) value on one of these attributes — the pattern
// every hard-coded label/tooltip/placeholder used before conversion, and
// none should use after it.
const LITERAL_ATTR_PATTERN = /\b(title|aria-label|placeholder)="[^"]/;

// A handful of exact English phrases that existed before conversion and
// must not have crept back in as raw JSX text (as opposed to `{t(...)}`).
// Deliberately short and exact (not a broad word-boundary scan across all
// English words) to avoid false positives on code comments, class names, or
// the untranslated product name "Atlas".
const BANNED_LITERAL_TEXT: ReadonlyArray<{ file: string; text: string }> = [
  { file: 'src/components/Toolbar.tsx', text: '>Save<' },
  { file: 'src/components/Toolbar.tsx', text: '>Open<' },
  { file: 'src/components/ShortcutsModal.tsx', text: '>Keyboard Shortcuts<' },
  { file: 'src/components/UnsavedChangesDialog.tsx', text: '>Unsaved changes<' },
  { file: 'src/components/WelcomeScreen.tsx', text: '>Open File<' },
  { file: 'src/components/StatusBar.tsx', text: '>Untitled<' },
];

const REPO_ROOT = resolve(__dirname, '../../..');

describe('no hard-coded UI strings in converted i18n components', () => {
  it.each(CONVERTED_FILES)('%s has no literal title/aria-label/placeholder attribute', (relPath) => {
    const content = readFileSync(resolve(REPO_ROOT, relPath), 'utf-8');
    const match = content.match(LITERAL_ATTR_PATTERN);
    expect(match, `found a literal attribute in ${relPath}: ${match?.[0]}`).toBeNull();
  });

  it('has none of the known pre-conversion English literals left as raw JSX text', () => {
    const offenders = BANNED_LITERAL_TEXT.filter(({ file, text }) => {
      const content = readFileSync(resolve(REPO_ROOT, file), 'utf-8');
      return content.includes(text);
    });
    expect(offenders).toEqual([]);
  });
});
