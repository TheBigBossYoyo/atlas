import { readFileSync, readdirSync, type Dirent } from 'node:fs';
import { join, relative, resolve } from 'node:path';
import { describe, expect, it } from 'vitest';

/**
 * Lint-style guard, not a type check — TEST-6.
 *
 * This used to be an OPT-IN allowlist (`CONVERTED_FILES`): a component was
 * only checked once someone remembered to add its path here. A brand-new
 * component (`RawEditor.tsx` being the concrete example that shipped
 * unguarded — see I18N-1) was invisible to the test by construction, so "the
 * i18n test passes" never actually meant "no hard-coded strings" — it meant
 * "no hard-coded strings in the ~35 files somebody previously remembered to
 * list here".
 *
 * It is now OPT-OUT: every `.tsx` file under the component/viewer surface
 * roots below is scanned by default (`discoverSurfaceFiles`), and a file
 * only escapes the scan by being named in `OPT_OUTS`, with its own comment
 * explaining why. A brand-new component with a literal `title`/`aria-label`/
 * `placeholder` now fails this test the moment it's added, with no action
 * required to "opt it in" — see the "scans by default" test below, which
 * proves this against a synthetic file that is (deliberately) not, and can
 * never be, in any allowlist.
 */

const REPO_ROOT = resolve(__dirname, '../../..');

// Directories that hold component/viewer JSX. Anything under these is a
// "surface" for this guard's purposes; everything else (hooks, parsers,
// layout engines, model code, ...) either has no JSX or is out of scope for
// a UI-string guard.
const SURFACE_ROOTS = ['src/components', 'src/viewers', 'src/docx/editor', 'src/docx/render'];

// Single files outside the roots above that still render user-facing JSX.
const SURFACE_FILES = ['src/App.tsx'];

const EXCLUDED_DIR_NAMES = new Set(['__tests__', '__fixtures__', '__snapshots__', '__styles__']);

/**
 * Every entry here is a KNOWN, reported gap or a deliberate non-finding —
 * never a blind spot. Unlike the old `CONVERTED_FILES` allowlist, leaving a
 * file off this list does not exempt it: it gets scanned. A file only lands
 * here because it is named individually, with its own comment.
 */
const OPT_OUTS: ReadonlyArray<{ readonly file: string; readonly reason: string }> = [
  // --- Genuine gaps, owned by another agent working concurrently in a
  // sibling worktree this batch (phase4 batch 3). Not edited here per this
  // batch's file-ownership rules; reported instead. ---
  {
    file: 'src/viewers/spreadsheet/FrozenRowsStrip.tsx',
    reason:
      "TODO(i18n): hard-coded aria-label template `Frozen row ${r + 1}, column ${c + 1}` (~line 83) " +
      'on the per-cell input — needs `t(\'spreadsheet.frozenRowCellAria\', { row, col })` or similar. ' +
      'Owned by the spreadsheet-surface agent this batch (anything under src/viewers/spreadsheet/ is ' +
      'off-limits per this batch\'s ownership rules); left for them to convert.',
  },

  // --- Not a real gap: a deliberately non-translatable literal that this
  // guard's broadened pattern would otherwise misfire on. ---
  {
    file: 'src/docx/editor/toolbar/Toolbar.tsx',
    reason:
      "Its only literal is `placeholder={'#000000'}` on the custom hex-color input (~line 403) — an " +
      'example hex value, identical in every locale, not English prose. The file itself already ' +
      'documents this ("Not a translatable label — an example hex value, same in every locale") and ' +
      "deliberately used the `{'...'}` form to dodge the old guard's narrower bare-quote-only pattern; " +
      "this guard's pattern is broader (it also catches `={'...'}`/`={`...`}`), so the same exemption " +
      'is now made explicit here instead of silently relying on regex shape. Not otherwise touched — ' +
      "it's part of the docx editor toolbar surface another agent owns this batch.",
  },
];

const OPT_OUT_FILES = new Set(OPT_OUTS.map((entry) => entry.file));

function collectTsxFiles(absoluteDir: string, out: string[]): void {
  let entries: Dirent[];
  try {
    entries = readdirSync(absoluteDir, { withFileTypes: true, encoding: 'utf-8' });
  } catch {
    return; // A configured root that doesn't exist (yet) is not this test's problem.
  }
  for (const entry of entries) {
    if (EXCLUDED_DIR_NAMES.has(entry.name)) continue;
    const full = join(absoluteDir, entry.name);
    if (entry.isDirectory()) {
      collectTsxFiles(full, out);
    } else if (entry.name.endsWith('.tsx')) {
      out.push(relative(REPO_ROOT, full).split('\\').join('/'));
    }
  }
}

/** Every `.tsx` file under a surface root/file, scanned by default. */
function discoverSurfaceFiles(): string[] {
  const out: string[] = [];
  for (const root of SURFACE_ROOTS) {
    collectTsxFiles(resolve(REPO_ROOT, root), out);
  }
  for (const file of SURFACE_FILES) {
    out.push(file);
  }
  return out.sort();
}

/**
 * A literal (non-expression) value on `title`/`aria-label`/`placeholder` —
 * the pattern every hard-coded label/tooltip/placeholder used before
 * conversion, and none should use after it. Broader than the old
 * `LITERAL_ATTR_PATTERN`: that one only matched a bare `attr="..."`, which a
 * literal written as `attr={'...'}` or `` attr={`...`} `` slipped past
 * entirely (see the Toolbar.tsx opt-out above, which documents exactly that
 * gap being used deliberately). This one catches all four forms — bare
 * double/single-quoted, and brace-wrapped string/template-literal — while
 * still not matching a real expression (`attr={t('key')}`, `attr={variable}`),
 * since neither starts with a quote or backtick.
 */
const LITERAL_ATTR_PATTERN = /\b(title|aria-label|placeholder)\s*=\s*(\{\s*)?["'`]/;

interface LiteralAttrViolation {
  readonly attr: string;
  readonly snippet: string;
}

function findLiteralAttrViolation(content: string): LiteralAttrViolation | null {
  const match = content.match(LITERAL_ATTR_PATTERN);
  if (!match) return null;
  return { attr: match[1], snippet: match[0] };
}

describe('no hard-coded UI strings in component/viewer surfaces (TEST-6)', () => {
  const surfaceFiles = discoverSurfaceFiles();
  const scannedFiles = surfaceFiles.filter((file) => !OPT_OUT_FILES.has(file));

  it('found more than a handful of surface files to scan (the discovery walk itself is not broken)', () => {
    // A sanity check on the scan, not the product: if this drops to a
    // handful of files, `SURFACE_ROOTS`/`SURFACE_FILES` broke, not the app.
    expect(surfaceFiles.length).toBeGreaterThan(50);
  });

  it('every opt-out actually exists under a surface root (no stale entries)', () => {
    const stale = OPT_OUTS.filter((entry) => !surfaceFiles.includes(entry.file));
    expect(stale.map((entry) => entry.file)).toEqual([]);
  });

  it('every opt-out has its own non-trivial comment explaining why', () => {
    const uncommented = OPT_OUTS.filter((entry) => entry.reason.trim().length < 20);
    expect(uncommented.map((entry) => entry.file)).toEqual([]);
  });

  it.each(scannedFiles)('%s has no literal title/aria-label/placeholder attribute', (relPath) => {
    const content = readFileSync(resolve(REPO_ROOT, relPath), 'utf-8');
    const violation = findLiteralAttrViolation(content);
    expect(violation, `found a literal ${violation?.attr} in ${relPath}: ${violation?.snippet}`).toBeNull();
  });

  it('scans by default: a brand-new component is checked without being added to any list', () => {
    // Stand-in for "someone adds RawButton.tsx tomorrow and forgets i18n" —
    // deliberately NOT read from disk and NOT listed anywhere in this file,
    // so passing this test cannot be explained by an allowlist entry. The
    // old (opt-in) design would have silently ignored this exact component;
    // this one must catch it purely from the literal attribute shape.
    const brandNewComponentSource = `
      export function RawButton() {
        return <button aria-label="Do the thing">Go</button>;
      }
    `;
    const violation = findLiteralAttrViolation(brandNewComponentSource);
    expect(violation).not.toBeNull();
    expect(violation?.attr).toBe('aria-label');
  });

  it('also catches a literal written as `attr={\'...\'}` or a template literal, not just bare `attr="..."`', () => {
    expect(findLiteralAttrViolation(`<input placeholder={'0.00'} />`)).not.toBeNull();
    expect(findLiteralAttrViolation('<input title={`Row ${n}`} />')).not.toBeNull();
    // A real expression must NOT trip the guard — this is what conversion to
    // `t(...)` looks like, and must keep passing.
    expect(findLiteralAttrViolation(`<input placeholder={t('x.y')} />`)).toBeNull();
    expect(findLiteralAttrViolation(`<input aria-label={label} />`)).toBeNull();
  });

  // I18N-1 / acceptance criterion: "the inverted guard fails on today's
  // RawEditor.tsx and passes after your fix". `PRE_FIX_RAW_EDITOR_SOURCE`
  // below is exactly what `src/components/RawEditor.tsx` contained before
  // this change (captured from git history — see the batch report for the
  // exact commands used to double-check this live against the pre-fix file
  // via `git show`), so this is a permanent regression test for the guard
  // itself, not just a one-time manual check.
  const PRE_FIX_RAW_EDITOR_SOURCE = `
interface RawEditorProps {
  markdown: string;
  onChange: (value: string) => void;
}

export function RawEditor({ markdown, onChange }: RawEditorProps) {
  return (
    <div className="editor-panel">
      <div className="editor-panel__header">
        <span>Markdown Source</span>
      </div>
      <textarea
        className="editor-panel__textarea"
        value={markdown}
        onChange={(e) => onChange(e.target.value)}
        spellCheck={false}
        placeholder="Type or paste markdown here..."
      />
    </div>
  );
}
`;

  it('fails on the pre-fix RawEditor.tsx (its placeholder was a bare literal)', () => {
    const violation = findLiteralAttrViolation(PRE_FIX_RAW_EDITOR_SOURCE);
    expect(violation).not.toBeNull();
    expect(violation?.attr).toBe('placeholder');
  });

  it("passes on today's RawEditor.tsx (fixed — header and placeholder both come from t())", () => {
    const content = readFileSync(resolve(REPO_ROOT, 'src/components/RawEditor.tsx'), 'utf-8');
    expect(findLiteralAttrViolation(content)).toBeNull();
  });
});

/**
 * Regression guard, kept from the pre-inversion version of this file: a
 * handful of exact English phrases that existed before their component was
 * converted, and must not creep back in as raw JSX text (as opposed to
 * `{t(...)}`) even though `LITERAL_ATTR_PATTERN` above only looks at
 * attributes, not text nodes. Deliberately short/exact rather than a broad
 * scan for English words, to avoid false positives on code comments, class
 * names, or the untranslated product name "Atlas". Every file below is
 * currently owned by another agent (see the OPT_OUTS comments above for
 * `Toolbar.tsx`/`FrozenRowsStrip.tsx`'s sibling files this batch) — this is
 * read-only regression coverage, not something this batch acts on.
 */
describe('no known pre-conversion English literals left as raw JSX text', () => {
  const BANNED_LITERAL_TEXT: ReadonlyArray<{ file: string; text: string }> = [
    { file: 'src/components/Toolbar.tsx', text: '>Save<' },
    { file: 'src/components/Toolbar.tsx', text: '>Open<' },
    { file: 'src/components/ShortcutsModal.tsx', text: '>Keyboard Shortcuts<' },
    { file: 'src/components/UnsavedChangesDialog.tsx', text: '>Unsaved changes<' },
    { file: 'src/components/WelcomeScreen.tsx', text: '>Open File<' },
    { file: 'src/components/StatusBar.tsx', text: '>Untitled<' },
    // I18N-1 — RawEditor.tsx's own former literals, now converted.
    { file: 'src/components/RawEditor.tsx', text: '>Markdown Source<' },
    // Small unowned fixes made alongside I18N-1 (see the batch report).
    { file: 'src/components/ViewerErrorBoundary.tsx', text: '>Viewer crashed<' },
    { file: 'src/components/ViewerErrorBoundary.tsx', text: '>Try again<' },
  ];

  it('has none of the known pre-conversion English literals left as raw JSX text', () => {
    const offenders = BANNED_LITERAL_TEXT.filter(({ file, text }) => {
      const content = readFileSync(resolve(REPO_ROOT, file), 'utf-8');
      return content.includes(text);
    });
    expect(offenders).toEqual([]);
  });
});
