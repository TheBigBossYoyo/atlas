# Atlas — handoff prompt for the next session

Paste everything below the line into a fresh session.

---

You are taking over Atlas, a Windows document viewer/editor, for its owner Youssef
(GitHub TheBigBossYoyo). **Answer him in English.** Code, comments, commits, docs and the
app's source strings are English too. (Older docs in this repo say to answer in French —
that is stale; he asked for English on 2026-09-20.)

Working directory: `C:\Users\Youssef\Documents\Projects\md-reader`
Private repo: github.com/TheBigBossYoyo/atlas, branch `main`, released **3.7.0**.
Stack: Electron 44 + React 19 + Vite 8 + TypeScript 6, its own DOCX engine, SheetJS
spreadsheets, pdf.js, CodeMirror, PPTX/ODP editors, multi-document tabs.

## Read before doing anything

- `.sisyphus/HANDOFF.md` — current state and environment rules
- `.sisyphus/plans/atlas-phase4-backlog-and-batches-2026-09-20.md` — **the live plan.** The
  verified backlog, the batch structure, execution status, and every item found while
  executing. Batches 1-4 are done; **5 and 6 are not.**
- `.sisyphus/plans/atlas-phase4-sweep-A-matrix.md` and `-sweep-B-matrix.md` — the hands-on
  feature matrices
- `.sisyphus/plans/atlas-phase3-findings-register.md` — older findings (§14c has per-wave
  tables). Note its "review leftovers" list is stale: all four are fixed.
- `CHANGELOG.md` (the 3.7.0 entry summarises phase 4), `docs/KNOWN_LIMITATIONS.md`,
  `docs/ARCHITECTURE.md`, `docs/RELEASE.md`, `STRICT_MODE_TODO.md`
- `git log --oneline -40`

**Treat every doc claim as a hypothesis; the code is the truth.** Several of these files
have been wrong and had to be corrected against the code.

## The single most important thing to understand

Over 2026-09-20/21, roughly **thirty real defects** were found in this app — several of
them silent data loss, three of them features that had *never worked for any user* — while
the test suite was **fully green**: 3324 unit tests, 76 e2e specs, CI passing.

Every one was found by **driving the real application and inspecting the saved bytes**.
None were found by reading code, and none by the suite.

Worse, **three defects were actively asserted as correct by existing tests**:
- the Excel table filter test asserted `not.toContain('filterColumn')` after an edit that
  never touched the table's columns;
- a shortcut test asserted Ctrl+W was correctly swallowed in a text field;
- **Insert Hyperlink's test mocked `window.prompt`** — which Electron does not implement —
  so jsdom passed happily while the feature threw for every real user, forever.

So: when a sweep or a test says something works, ask *how it knows*. And when you fix a
bug, expect to find a test pinning it in place; updating that test is not weakening it.

## Current state (verify all of this yourself — do not trust this document)

- `main` = `95a28a8`, pushed, version **3.7.0**.
- Installer `release\Atlas-Setup-3.7.0.exe`, 136 MB, unsigned,
  SHA256 `ddb181307d210e42155bcb2abcb8fe861a22675d6186f6e58cb06f4a28f3ec40`.
- **3581 unit tests, 114 Playwright e2e** — all pass locally.
- `npx tsc -b` covers five projects (`src/`, `vite.config.ts`, `tests/e2e`,
  `electron/**/*.cjs`, `scripts/**/*.mjs`). `npm audit` 0. Bundle gate green.
- **CI is RED.** See below — this is your first job.

## Your next steps, in priority order

### 1. CI is red, and one failure may mean the 3.7.0 release is wrong

`gh run list` shows a failure on `95a28a8`, which is a **docs-only commit** — so whatever
is failing is not caused by that change.

The failing spec is:
```
tests\e2e\spreadsheet-keystroke-seed.spec.ts:91
  › F6 — click a cell and type: the leading character must not be dropped (xlsx)
  › typing at 120ms/char with a 1000ms pause after the click saves the full word
  Expected: "HELLO"   Received: "Name"
```

**Read that carefully.** `"Name"` is the fixture's *original* cell content. So the whole
edit was lost — the cell was never modified at all. That is **not** the bug this test
guards (which dropped the *leading character*, giving `"ELLO"`). Something else is wrong on
the CI runner: most likely the grid is not yet interactive when the test clicks, so the
click and the typing go nowhere.

Two possibilities, and they matter very differently:
- **Test readiness issue.** The spec waits 1500ms after the viewer appears before clicking.
  The agent that wrote it already had to raise this from 800ms and noted that
  click-to-selection has its own async latency. On a slower runner 1500ms may still be too
  little. Fix by waiting on a real signal (the grid reporting a selected cell), not a sleep.
- **A real product problem.** If the grid can silently swallow a click and subsequent typing
  while it is still initialising, a real user on a slow machine loses work the same way.
  **This is the same shape as the bug 3.7.0 shipped to fix**, so do not assume it is only a
  test issue.

Decide which with evidence. The F6 fix itself was verified independently: a probe at
`...\scratchpad\keystroke-probe.mjs` (see "Useful artefacts") went from 24/24 failures to
48/48 correct. So the fix works locally; the question is whether it holds on a slow machine.

**Also relevant:** three *other* CI-only failures happened that night, all on the
`windows-latest` runner, all budget/timeout-shaped, all passing locally or on retry —
`export.spec.ts` twice and `perf.spec.ts` once (a 592ms cold-start task against a 300ms
CI-adjusted budget; passed on re-run of the same commit). There may be a single systemic
story here: **the CI runner is slower and more variable than several of these budgets
assume.** Worth investigating as one question rather than four.

One thing already done to help: main-process stderr is now piped into the export tests, so
a silent `PrintToPdfError` is no longer indistinguishable from slowness. Consider doing the
same for other e2e specs — the absence of main-process diagnostics is why these are hard.

### 2. The packaged installer has never been smoke-tested

`docs/RELEASE.md` step 8 was deliberately left undone — it needs an interactive UAC prompt
and a human at the machine. **Ask Youssef to do it, or do it with him present.** Install,
launch from the Start Menu, open a `.docx`/`.xlsx`/`.pdf` via Ctrl+O *and* via an Explorer
double-click (the only check of the file associations the installer registers), make an
edit and save. Until then 3.7.0 is verified only as source, not as a shipped artefact.

### 3. Batches 5 and 6

Both are specified in the plan file with per-agent objectives, file ownership, files not to
touch, acceptance criteria and required tests.

**Batch 5** (remaining fidelity): DOCX-3 paragraph-mark `rPr` · DOCX-4/5 `docPr` ids and
anchor attributes · DOCX-13 hidden text · FID-2 hidden rows/columns · FID-4/5 PDF
destinations and annotation text · SEC-2/3 printToPdf guards and legacy encryption
detection · SHEET-7 / SLIDE-1 / UX-FR / CODE-1 missing controls and small UX fixes.

**Batch 6** (large work, only if there is a reason): FID-1 per-cell spreadsheet styling ·
FID-3 cell comments · FID-6 PDF layers · DOCX-6…DOCX-11 · SEC-4 (a `patch-package` or
watchdog for a cyclic-sector-chain hang in the vendored SheetJS CFB reader).

### 4. Open items worth promoting into a batch

> **Reconciled 2026-09-29.** Every item this section listed had already been fixed and
> marked so in `plans/atlas-phase4-backlog-and-batches-2026-09-20.md` — all seven of
> them — while this list still presented them as open work. That is the same drift the
> backlog file warns about in its own reconciliation note, and it has already cost two
> sessions re-investigating closed items from scratch. Checked against the backlog's
> status markers, not from memory:
> INSERT-TEXT-THROWS-1 **FIXED** · QUIT-DRAFT-1 **RESOLVED** · ZIP64-1 **FIXED** ·
> I18N-TEXT-1 **FIXED** · FROZENROWS-I18N-1 **RESOLVED** · TEST-9/TEST-10 **RESOLVED** ·
> FIXTURE-1 **RESOLVED** · CHARTSHEET-1 **closed as a feature request**.
>
> **If you fix something, mark it in the backlog AND delete it from here.** This section
> is a pointer, not a second source of truth.

What is genuinely open, highest value first:

- **Office verification: DOCX done, the rest open.** LibreOffice reads what Atlas writes
  for all 22 corpus fixtures, tables included (OFFICE-VERIFY-1, see section 5). Still
  open: the XLSX and PPTX write paths, which have no equivalent harness, and Microsoft
  Office itself.
- ~~Five drawing/numbering fidelity gaps~~ **ALL FIXED 2026-09-29** (FID-NUM-1 `1884f62`,
  FID-DRAW-1 `8879e6f`). The corpus fidelity detector now reports zero unexplained loss;
  the only accepted entries left are the two `w:tblW` ones that note 6 in
  `lossySaveWarnings.test.ts` established are safe by reading `parseWidth`. Note none of
  these change a character of text, so the LibreOffice comparison is blind to them — keep
  both checks.
- **PROGID-1's install verification.** The ProgIDs an install registers changed from
  generic labels to `Atlas.*`; source and build are verified, the registry behaviour is
  not. `docs/RELEASE.md` step 8 carries the `reg query` checks. Needs an elevated install
  on a real machine.
- ~~MATRIX-FLAKE-1~~ **FIXED 2026-09-30.** The cell editor claims DOM focus for itself;
  `autoFocus` fires once and glide's re-render could leave focus on an accessibility-table
  `<td>`, losing every key after the first. Two clean full e2e runs against a 2-in-4
  baseline. **Read the backlog entry before touching lost-keystroke bugs in this grid** —
  the first fix attempt went into `KeyHold` and was wrong, which makes four in a row to
  make that same mistake.
- **BUNDLE-BASELINE-1.** The bundle baseline is from 2026-09-20 and the build is +7.96%
  against it, so a real regression has less of the 25% window to show up in. Refreshing
  it also erases the drift record, which is why it is the owner's call.

### 5. Still true, still unaddressed

- **LibreOffice reads what Atlas writes — all 22 corpus fixtures, tables included.**
  `src/docx/__tests__/libreOfficeRoundTrip.test.ts` (OFFICE-VERIFY-1) has
  LibreOffice convert each fixture and the same file after Atlas parses and
  re-saves it, and asserts the extracted text matches. 22 of 22 pass, no
  exclusions.

  Two things that had to be got right first, both worth knowing before touching
  this test:
    - **The `txt:Text` filter hangs on any document containing a table** in
      LibreOffice 26.8.0 — a real runaway, 620 MB of temp output in 120 s and
      4.6 GB if left. Bisected to be sure it was not our markup: stripping
      `tblPr`/`tblGrid`/`tblLook`/`tblBorders`/`tblStyle` changed nothing, removing
      the `<w:tbl>` fixed it, a hand-built minimal 1x1-table `.docx` hangs too, and
      that same minimal file converts to `odt` and `html` without trouble. The test
      uses the HTML filter, which is why tables are covered rather than excluded.
    - **Newlines in LibreOffice's HTML output are not content.** It wraps its own
      output, so a span can come back split across lines. An earlier version of the
      test preserved those newlines and reported `rtl-text` as a content difference
      because of it. Block boundaries are marked with a sentinel first, then all
      real whitespace collapses, then the sentinel becomes a newline.

  Still not proven: **Microsoft Office itself**, and the XLSX/PPTX write paths,
  which have no equivalent harness. LibreOffice is also more forgiving in places
  and stricter in others, and reads none of the `w15:`/`w16:` extensions this
  codebase works to preserve; text equality says nothing about formatting, which is
  what `fidelity/__tests__/lossySaveWarnings.test.ts` is for. The test SKIPS when
  `soffice` is missing (CI has none), so a green suite does not mean it ran —
  `scoop install extras/libreoffice`, per-user, no elevation.
- **Microsoft Office is still untested, and the XLSX/PPTX write paths have no external
  check at all.** `scripts/validate-office-file.mjs` remains the only substitute there,
  and it passed two genuinely invalid files clean before attribute datatype checks were
  added — treat a clean validator run as necessary, not sufficient.
- Code signing needs a paid certificate (`docs/RELEASE.md`) — the one item needing money.
- A closed spreadsheet tab releases 92-94% (re-measured 2026-09-28; the old ~24 MB figure was wrong, and `performance.memory` is quantized/cached and cannot measure this).
- `.doc`/`.ppt` are read-only text. No split view.

## How to work here

**Drive the real app.** Use Playwright's `_electron.launch` from standalone Node scripts
(pattern in any `tests/e2e/*.spec.ts`), with:
```
env: { ...process.env, CI: '1', PLAYWRIGHT: '1', ATLAS_HIDDEN_WINDOW: '1' }
```
`ATLAS_HIDDEN_WINDOW=1` makes the window invisible (opacity 0, `showInactive`,
click-through) while still fully composited. **It is NOT headless — a real window exists.
Never describe a run as headless.** `show: false` does not work: Chromium stops compositing
and `requestAnimationFrame`, so Playwright's actionability checks hang forever.

**Verify against the saved bytes**, not the screen. Unzip the `.docx`/`.xlsx`/`.pptx` and
assert on the XML. Almost every bug found this cycle looked fine on screen.

**Run `node scripts/validate-office-file.mjs <file>`** on anything the app saves.

### Environment rules (learned the hard way)

- Windows 11. PowerShell primary, Git Bash available. **Bash heredocs mangle backslashes,
  quotes and `\u` escapes — use the Write/Edit tools for source files.** A Python heredoc
  (`python - <<'PY'`) is reliable for scripted edits.
- `npx vitest run --maxWorkers=3`. Playwright `workers: 1`. **Never build + vitest +
  Playwright at once** — this laptop has hard-powered-off under full CPU load.
- Before any Electron probe or e2e run: `npx vite build`.
- Kill Electron with `taskkill /PID <pid> /T /F`. After any e2e run:
  `git checkout -- tests/e2e/fixtures/`.
- **Agent worktrees:** the harness's own worktrees have **no `node_modules`**. Create them
  manually and add a junction:
  `cmd //c "mklink /J <worktree>\node_modules C:\Users\Youssef\Documents\Projects\md-reader\node_modules"`
  Remove it with `cmd //c rmdir <worktree>\node_modules` **before** `git worktree remove`.
  **Never `rm -rf` or `Remove-Item -Recurse` through a junction**, and never run
  `npm install` inside a worktree. Do these one path at a time — a bash `for` loop mangles
  the backslashes and silently does nothing.
- **Never `git stash`** — the stash stack is shared across all worktrees and concurrent
  agents. Use a throwaway WIP commit plus `git checkout <sha> -- <files>`.
- `npm run build` enforces a bundle-size gate. The baseline
  (`.sisyphus/baselines/atlas-phase3-bundle.json`) is **stale** — it predates the
  lazy-loading work, listing a 2.6 MB entry chunk against today's 331 KB. Refresh it
  deliberately with `node scripts/capture-bundle-baseline.mjs` and say why.
- Conventional commits ending with a blank line then
  `Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>`.
- Youssef is on a Max x5 plan and watches usage: **few, well-scoped agents**. Tell him the
  cost shape of a batch before running a big one.

### Running agents in parallel (this worked well — 15 agents, zero collisions)

- One worktree + branch per agent, off current `main`, with a `node_modules` junction.
- **Disjoint file ownership per batch.** Give each agent an explicit "files you own" and
  "files you must NOT touch" list. The usual collision points: `src/App.tsx`,
  `src/viewers/DocxViewer.tsx`, `electron/main.cjs`, `src/docx/serializer/documentWriter.ts`,
  `src/viewers/spreadsheet/xlsxPassthrough.ts`, the i18n catalogues, `CHANGELOG.md`.
- **Only one agent may touch the i18n catalogues per batch.** Others describe the key and
  English text they need; add them yourself at merge. This works well.
- Tell every agent: targeted `npx vitest run --maxWorkers=1 <paths>` only, at most one
  Playwright spec at a time, **never** the full suite or `npm run build`.
- Require: reproduce before fixing, and prove each new test **fails before the fix and
  passes after**. Several agents caught themselves this way.
- **Verify their claims.** Re-check every critical finding against the code yourself. One
  agent's tests didn't run in CI at all (wrong config); one reported a feature as working
  that had never worked; one characterised a deterministic bug as an intermittent race.
- Gate between batches, in order, never concurrently:
  `npx tsc -b` → `npx eslint src electron tests scripts` → `npx vitest run --maxWorkers=3`
  → `npm run build` → `npx playwright test` → push → **watch CI** (red `e2e-windows` blocks
  a release per `docs/RELEASE.md`, even when local passes).
- Check `git ls-files` for stray `zz*`/`scratch`/`repro`/`debug` files before merging —
  agents leave them in worktrees, and occasionally commit them.

## Useful artefacts from the last session

- `...\scratchpad\keystroke-probe.mjs` — parameterised probe (typing speed × pause after
  click) that saves a real `.xlsx` and reads the cell back from the bytes. This is what
  settled the first-keystroke bug when three agents had mischaracterised it. The pattern —
  sweep a parameter, N trials each, assert on saved bytes — is reusable.
- `ATLAS_E2E_PROFILE_DIR` — opt-in env var letting two consecutive `electron.launch()` calls
  share a `userData` profile, so anything persisted (recent files, window state, the
  `recentFilesStore` security ground truth) is testable across a relaunch. Sequence the
  relaunches yourself: `taskkill`, then **await the process's `exit` event**, or you hit the
  singleton lock. Call `session.flushStorageData()` before killing, or `localStorage` writes
  can be lost to the kill.

## Quality bar

Fix causes, not symptoms. **Never weaken a test, a lint rule or a tsconfig flag to make
something pass** — but do correct a test that asserts a bug as correct, and say so. Every
behaviour change gets a test; anything that only shows in the real app gets an e2e test.
Report honestly: if something was not run, say so; if a measurement did not complete, do
not quote a number. An agent's claim of evidence must match what it actually ran — that has
been wrong repeatedly.
