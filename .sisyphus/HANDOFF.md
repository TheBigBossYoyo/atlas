# Atlas — session handoff (2026-09-20)

Atlas is a Windows document viewer/editor (Electron 44 + React 19 + Vite 8 + TypeScript 6), private
repo github.com/TheBigBossYoyo/atlas, branch `main`. The owner (Youssef, GitHub TheBigBossYoyo) writes
in **French** — answer him in French. Code, commits, docs and UI stay in English.

## Read first
- `.sisyphus/plans/atlas-phase3-improvement.md` (plan; the status section near the end is kept current)
- `.sisyphus/plans/atlas-phase3-findings-register.md` (§14/§14b owner-reported issues, §14c per-wave
  resolution tables — waves 4, 5 and 6 are recorded there with commit SHAs)
- `CHANGELOG.md`, `STRICT_MODE_TODO.md`, `docs/KNOWN_LIMITATIONS.md`, `docs/RELEASE.md`
- `git log --oneline -30`

## Current state (3.3.0)
- `main` is pushed and released as **3.3.0**; installer `release\Atlas-Setup-3.3.0.exe`
  (SHA256 `87d63a4d01c9d47b15e49f60876927128492d7a9f17f673394b749c9715bed32`, unsigned).
- 2839 unit tests, 57/57 Playwright e2e, `npm audit` 0, bundle gate green.
- `npx tsc -b` now covers four projects: `src/`, `vite.config.ts`, `tests/e2e`, and
  `electron/**/*.cjs` (`checkJs` + `strict`). Never weaken these.

### Shipped since 3.2.0 (waves 5 and 6)
New documents (toolbar **New** + Ctrl+N) from blank templates in `electron/templates/`, and a 0-byte
file of an editable format opens as a blank document (the owner hit this with an Explorer-created
empty `.docx`); spreadsheet save-through-original now covers sheet add/delete/rename/reorder and
re-anchors conditional formatting, data validation, hyperlinks and autoFilter across row/column
inserts; header/footer editing preserves images, fields and tables; DOCX page virtualization
(~145 ms per keystroke on the 35-page fixture, from ~0.5 s); `electron/*.cjs` type-checked; PDF
component coverage closed; packaged exe metadata/icon fixed (`signAndEditExecutable: false` had been
disabling resource editing entirely).

### Shipped since 3.3.0 (waves 7 and 8, not yet in a version bump)
Save As (`Ctrl+Shift+S`) fixed for every non-markdown format; the OPC/OOXML/ODF spec-level package
validator (`scripts/validate-office-file.mjs`); two a11y passes (focus traps on every remaining
dialog/overlay, WCAG AA contrast fixed across all 5 themes); header/footer editing is now per-SEGMENT
(a paragraph mixing text with a field/image is editable, not just read-only); spreadsheet formula
references and defined names are now re-anchored through every structural edit, closing the leftovers
listed below; an English/French UI (`src/i18n/`, DEFER-6 resolved — the owner is confirmed French-
speaking); `scripts/*.mjs` and `tests/e2e/fixtures/*.mjs` are now type-checked too (P4.1 fully closed).

## What is left
1. **Code signing (P5.1)** — the only item needing the owner's money. `docs/RELEASE.md` lists the
   certificate options and what changes in `electron-builder.yml`.
2. **No Office verification** — nothing Atlas writes has ever been opened in real Microsoft Office or
   LibreOffice (neither is installed here). Structural validation is the substitute (see above).
3. `.doc`/`.ppt` are read-only, text only. No split view.
4. i18n: `DocxViewer`, the slide editor, and PDF find/thumbnail internals are not translated
   (`docs/KNOWN_LIMITATIONS.md`'s "Internationalization" section has the full list).
5. A version bump + CHANGELOG/findings-register update for waves 7-8 hasn't happened yet — everything
   above is on `main` but still sits under CHANGELOG's `[Unreleased]` heading.

## Environment rules (learned the hard way)
- Windows 11; PowerShell primary, Git Bash available. Bash heredocs mangle `\\`, quotes and `\u`
  escapes — use the Write/Edit tools for source files.
- `npx vitest run --maxWorkers=3`; Playwright `workers: 1`; never build + vitest + Playwright at once
  (the laptop has hard-powered-off under full load). Kill apps with `taskkill /PID <pid> /T /F`.
- Before any Electron probe or e2e run: `npx vite build` (the app loads `dist/` when it exists).
- Running e2e rewrites `tests/e2e/fixtures/*` — `git checkout -- tests/e2e/fixtures/` before committing.
- Agent worktrees get `node_modules` as a junction to the main install: remove it with
  `cmd /c rmdir <path>\node_modules` before `git worktree remove`, NEVER `rm -rf` through it.
- `npm run build` runs a bundle-size gate; refresh intentionally with
  `node scripts/capture-bundle-baseline.mjs`.
- The owner's personal `.docx` in `Downloads\` may only be used via temp copies, never committed.
- Conventional commits ending with `Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>`.
  Verify fully before pushing; CI runs lint, `tsc -b`, coverage, build, bundle gate and the Windows
  Electron e2e suite.
- When an e2e test flakes only on CI, suspect the product first: the last three "flaky" failures were
  a real bug (the PDF find bar stayed open and swallowed the next shortcut).
