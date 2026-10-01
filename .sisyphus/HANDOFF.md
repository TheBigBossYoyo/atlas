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

## Current state (3.8.0)
- `main` is pushed and released as **3.8.0** (2026-09-27); installer
  `release\Atlas-Setup-3.8.0.exe` (SHA256
  `3b83b808bbd9001fb6d7ff604e9d9d324de944a9af86327630255dc4fe108a17`, 136 MB, unsigned).
- **Embedded exe metadata verified on the real build** (the check `docs/RELEASE.md` said still
  needed a lead build + Properties -> Details): ProductName/FileDescription/CompanyName
  `Atlas`, FileVersion `3.8.0`, ProductVersion `3.8.0.0`, LegalCopyright
  `Copyright (c) 2026 Atlas`. ELEC-10 is therefore confirmed fixed against a real artifact,
  not just against the yml.
- **Step 8 (packaged installer smoke test) is DONE for 3.8.0, upgrade path included.** The
  3.8.0 installer was run over the existing 3.7.0 install (owner accepted the UAC prompt;
  `/S` silent, so that was the only interaction). Result: one Control Panel entry at 3.8.0
  (no duplicate), Start Menu shortcut intact, and `recent-files.json` / `window-state.json` /
  `Local Storage` all **byte-identical** afterwards. **This closes P5.3's installer
  upgrade-path item.** F6c was then re-verified against `C:\Program Files\Atlas\Atlas.exe`
  itself — 5/5 at 8x CPU throttle. Full table in the phase-4 plan's "3.8.0 release
  verification". Uninstall verification deliberately skipped: it would remove the working
  install, and the upgrade already exercised NSIS's internal uninstall-and-replace.
- 3.8.0 fixes **F6c** (HIGH, silent data loss: a second cell edit discarded and the following
  Ctrl+S writing nothing). 3.7.0 and earlier all have it. Anyone still on 3.7.0 should upgrade.
- At 3.8.0: **3839 unit tests, 142 Playwright e2e, coverage 82.58/70.41/85.10/84.44,
  `npm audit` 0, bundle gate green, CI green.**
- **P5.1 code signing is CLOSED as WON'T DO** (owner decision, 2026-09-27). Atlas ships
  unsigned, permanently: expect the SmartScreen "unknown publisher" warning on every release,
  treat the published SHA256 as the only integrity check, and keep auto-update out of scope
  (DEFER-7) — an updater running unsigned downloads is a malware path. Do not re-open this as
  a backlog item; see `docs/RELEASE.md`'s "Code signing — DECIDED AGAINST".
- Still open: RPR-STYLES-1 (unknown `w:rPr` children dropped from `styles.xml` on save, re-verified
  2026-09-27), REDO-REPLAY-1, and the comments pane having no un-resolve affordance (new in
  3.8.0 — resolving now persists, so it is no longer undone by reopening the file).

## Previous state (3.7.0)
- `main` is pushed and released as **3.7.0**; installer `release\Atlas-Setup-3.7.0.exe`
  (SHA256 `ddb181307d210e42155bcb2abcb8fe861a22675d6186f6e58cb06f4a28f3ec40`, 136 MB, unsigned).
  Earlier installers (3.4.0-3.6.0) are in `release\` too.
- **The 3.7.0 installer was smoke-tested on 2026-09-24** (`docs/RELEASE.md` steps 1-5; results
  in the phase-4 plan, "packaged installer smoke test"). All passed except Ctrl+O inside a
  document's text (SHORTCUT-FIELD-1). Step 6 (uninstall) not done. 3.7.0 also still has F6b
  (first keystroke lost / Ctrl+S overtaking an edit on a slow machine), fixed on `main` since.
  Cut a 3.7.1.
- 2026-09-25 night run merged ~20 fixes found by driving the app (see the phase-4 plan's
  "2026-09-25 night run" section and `..\atlas-night\NIGHT-LOG.md`). Unit tests: 3672. DOCX-3 (focus lost
  after Save As, via ViewerRouter's path-keyed remount) is the main open item from it.
- (at 3.7.0: 3581 unit tests, 114 Playwright e2e, `npm audit` 0, bundle gate green, CI green.)
- `npx tsc -b` covers five projects: `src/`, `vite.config.ts`, `tests/e2e`,
  `electron/**/*.cjs` and `scripts/**/*.mjs`. Never weaken these.

### Shipped in 3.7.0 (phase 4) — read `.sisyphus/plans/atlas-phase4-backlog-and-batches-2026-09-20.md`
Everything in this release was found by **driving the real application**, not by the test
suite, which was green throughout. Three defects were actively *asserted as correct* by
existing tests — most starkly Insert Hyperlink, whose test mocked `window.prompt` (which
Electron does not implement), so jsdom passed while the feature had never worked for anyone.

Features that had never worked: Insert Hyperlink / Add Comment / Reply; Table Properties
(borders, width and alignment were all discarded, because one invalid control silently
blocks a whole HTML form submit); the first character typed into a spreadsheet cell;
Insert Table; the caret inside a table cell; Insert Text Box on a slide. Invalid OOXML:
every colour (`w:val="#ff0000"`) and every new list. Silent save losses: theme fonts, a
class of character effects, Excel table filters, whole-column conditional formats,
wrappers in notes/comments, Save As hijacking another tab, text-returning formulas.
Security: a compromised renderer could silently overwrite any existing file.

**Method note for the next session:** the value came from hands-on sweeps with byte-level
verification of saved files, not from reading code and not from the suite. Budget for that.

### Shipped since 3.3.0 (waves 7-11, one autonomous overnight run)
Data loss fixed: Save As left a tab pointing at the file it was opened from, so later saves went
nowhere; every saved DOCX image lost its picture properties; legacy `.doc`/`.ppt` failed on any
file over ~4 KB (every fixture was smaller); `.ods` was written with a non-conforming `mimetype`;
several parts were written with their elements in the wrong schema order — Word's "unreadable
content" trigger. A spec-level package validator (`scripts/validate-office-file.mjs`) now runs in
CI and would catch all of those. Spreadsheet formulas, defined names, shared formulas and media
follow every structural edit. Unedited content controls and shape fallbacks round-trip
byte-for-byte, and a save that still drops something says so in plain language. The UI can be
French (English stays the default, `src/i18n/`, 509 keys). Markdown parses in a Worker, so large
documents no longer freeze the window. Start-up is ~33% faster (entry bundle -89%). Accessibility:
WCAG AA contrast in all five themes, focus traps everywhere, a keyboard-only end-to-end journey.

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

## 3.10.0 — built 2026-10-01

`release\Atlas-Setup-3.10.0.exe`, 136.2 MB, unsigned.
SHA256 `A59FB1F3FF2B3311DB8E92A4948C8A1ABA662B6140B4EC4EA184DB49C30247B6`.

Version history now captures every editable format between saves, not just
markdown (VERSIONS-2), and a restored text version keeps the file's own encoding
and line endings. Verified in the packed `app.asar` before shipping: the new
i18n string, the renderer text encoder and the history IPC are all present, as is
3.9.0's drag-region fix.

Installing needs an elevated run (perMachine), so it needs the owner at the
machine to accept the UAC prompt.

After installing, run `docs/RELEASE.md` step 8. The PROGID-1 `reg query` checks
in it were verified on 3.8.0 and re-verified on 3.9.0, so that is a re-check
rather than a first run.

## 3.9.0 — built 2026-09-30, installed 2026-10-01

`release\Atlas-Setup-3.9.0.exe`, 136.2 MB, unsigned.
SHA256 `39710E3EE3A14CF22B4BDB62567B62682CDED862C2D033709D2F82AB1C1BB4C1`.

Installed with the owner present (UAC accepted, installer exit 0), and the
toolbar-menu fix confirmed in the installed `app.asar`. Superseded by 3.10.0
above.

## What is left
1. **Code signing (P5.1)** — the only item needing the owner's money. `docs/RELEASE.md` lists the
   certificate options and what changes in `electron-builder.yml`.
2b. **Accessibility**: an automated WCAG 2.1 A/AA audit now exists
   (`tests/e2e/accessibility.spec.ts`, axe-core injected into the real Electron window,
   nine surfaces) and passes with zero violations after five fixes — see the Accessibility
   section of `docs/KNOWN_LIMITATIONS.md`. Testing with a REAL screen reader still has not
   happened, and axe cannot substitute for it.
2. **Office verification — all three write paths done (2026-09-30).** LibreOffice reads
   what Atlas writes for Word (22/22 corpus fixtures), spreadsheets (`.xlsx`,
   multi-sheet `.xlsx`, `.ods`, through BOTH the passthrough and the rebuild branch) and
   slides (`.pptx`, multi-slide `.pptx`, `.odp`). Shared plumbing in
   `src/__tests__/helpers/libreOffice.ts`; each harness has a self-check that changes one
   cell/paragraph/slide and requires the comparison to catch it. Use the HTML/flat-ODF
   filters, never `txt:Text` — that hangs on any document with a table in LibreOffice
   26.8.0. Microsoft Office itself is still untested, and
   `scripts/validate-office-file.mjs` remains necessary-not-sufficient.
3. **Memory**: re-measured 2026-09-28 and the ~24 MB claim is WRONG — a closed 100k-row
   spreadsheet tab gives back 92-94%, leaving ~0.3-0.4 MB per open/close cycle. Pinned by
   `tests/e2e/spreadsheet-memory-release.spec.ts`; method and the two measurement traps are in
   `docs/KNOWN_LIMITATIONS.md`.
4. `.doc`/`.ppt` are read-only, text only. No split view. A 3-D spreadsheet reference only shifts
   from its first sheet (matching Excel itself).
5. A content control or shape whose content was edited, or one inside a table, header or footer,
   is still unwrapped on save — reported to the user now, not silent.

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
