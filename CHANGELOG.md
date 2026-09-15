# Changelog

All notable changes to Atlas, adopted per Task P4.9/QA-24 (no changelog had
existed across 9 undocumented patch releases before this). Entries are
grouped by the internal "wave" merges that produced them — see
`.sisyphus/plans/atlas-phase3-improvement.md` for the task IDs and findings
each wave closed — rather than by individual commit, since a wave is this
project's real unit of shipped, reviewable work. Dates are merge dates from
`git log`.

## [3.1.0] — 2026-09-15 (wave 3, docs-quality slice)

### Added
- `docs/ARCHITECTURE.md` and `docs/KNOWN_LIMITATIONS.md`.
- This changelog.
- A per-chunk bundle regression gate (`scripts/check-bundle.mjs`, wired as a
  `postbuild` step) covering every `dist/assets/*.{js,css}` chunk, not just
  the main entry bundle, with a freshly captured baseline.
- Coverage thresholds (`vitest.config.ts`) and an `npm run coverage` script.
- A shared, typed `createMockElectronAPI()` test fixture
  (`src/__tests__/mocks/electronAPI.ts`) and a dedicated `useRecentFiles`
  test suite.
- CI: Node-24-compatible action majors, a coverage-artifact upload step, and
  a Windows Electron end-to-end job on pushes to `main`.

### Fixed
- `npm run electron:preview` now actually previews the production build
  instead of silently falling back to a dev server that might not be
  running (`electron/main.cjs`'s dev/prod detection now keys off whether
  `dist/index.html` exists, with an `ATLAS_DEV` override).
- README's feature/shortcut/stack documentation brought back in sync with
  what the app actually does (previously described only the Markdown
  feature set from Phase 1, with zero mention of the other 12 formats).

## [3.0.x] — 2026-09-14 (wave 2)

Eight parallel worktrees, merged into `main` over one day. Full task-ID
mapping in the improvement plan; highlights below.

- **DOCX fidelity** (`wave2/docx-model`, `wave2/docx-layout`): style-cascade
  resolution wired through layout, alignment/indent/spacing, justification,
  list markers, hyperlink rendering, manual and section page breaks, table
  style cascade, `sdt`/`AlternateContent`/symbol handling, resolved
  `commentsExtended` state, serializer round-trip fidelity, Aptos font
  mapping with lazy-loaded font assets.
- **DOCX editing** (`wave2/docx-editor`): cross-paragraph editing,
  transactional undo/redo, image insert via the same history, a spellcheck
  fix, accept/reject track changes plus a `trackChanges` setting, list/
  indent/hyperlink/page-break/table insert commands, Save As.
- **PDF viewer** (`wave2/pdf`): page virtualization, text and annotation
  layers, in-document find, print, interactive-form-field rendering, page
  rotation, a thumbnail rail, password-protected PDF support.
- **Slides** (`wave2/slides`): PPTX/ODP layout and master/slide-layer
  inheritance, run formatting, bullets, tables, grouped shapes, speaker
  notes, keyboard navigation, thumbnails.
- **Spreadsheets, CSV, code, RTF, ODT** (`wave2/data-text`): formatted cell
  values, merged cells, column widths, hidden sheets, off-main-thread
  worker parsing (with a 200ms task-perf budget spec), text/code
  virtualization, RTF/ODT size/security guards, expanded fixtures.
- **Format detection & routing** (`wave2/formats`): the single canonical
  extension manifest (`src/formats/extensionManifest.ts`, generated at
  `prebuild`), extensionless/legacy-CFB detection, magic-byte verification
  on the recognized-extension fast path, the `UnknownViewer` empty state.
- **Export & UX polish** (`wave2/export-ux`): markdown export reads the
  live rendered DOM instead of re-parsing (fixing export/live-render
  divergence), DOCX export gained math/Mermaid/hyperlink/task-list
  support, Electron-native save dialogs for exports, a toast system
  replacing blocking `alert()`s, real CSV export, accessibility and
  contrast fixes.
- **Shell** (`wave2/shell`): the centralized shortcut dispatcher
  (`ShortcutManagerProvider`, replacing six independent `keydown`
  listeners), main-process close-confirmation guard, draft recovery,
  close-file + dynamic window title, in-document search extended beyond
  markdown to text/code/RTF/ODT.

## [3.0.x] — 2026-09-13 (wave 1)

- **CI & lint** (`wave1/ci-lint`): CI workflow stood up, lint brought to
  zero warnings/errors.
- **Markdown characterization** (`wave1/markdown-characterization`): the
  snapshot suite that is the executable definition of "markdown behavior
  must not change" for every later wave.
- **DOCX round-trip corpus** (`wave1/docx-corpus`): a fixed fixture set
  parsed → serialized → re-parsed and diffed, gating every later
  serializer-touching change.
- **DOCX comments fix** (`wave1/docx-comments`): corrected
  `extractCommentText` operator-precedence bug.
- **Viewer quickfixes** (`wave1/viewer-quickfixes`): assorted viewer polish
  and error-recovery fixes.
- **DOCX save integrity** (`wave1/docx-save-integrity`): save-path
  robustness hardening.
- **Electron hardening** (`wave1/electron-hardening`): main-process
  security hardening — sandboxing, IPC allowlisting, and related fixes.
- **Shell session** (`wave1/shell-session`): the unified dirty-state/save
  capability contract (`ViewerContext`'s `isDirty`/`setDirty`/
  `registerSave`) and session UX (recent-files dedup, discard-changes
  confirmation, browser-mode guard).
- **Dependency & build-chain security** (`wave1/deps-security`): dependency
  hardening.

## Before 3.0 — Phases 1 & 2

Phase 1 (view-only, 13 formats, Atlas rebrand from "MD Reader") and Phase 2
(hand-coded DOCX editing engine) predate this changelog's adoption. See
`.sisyphus/plans/atlas-phase1.md` and `.sisyphus/plans/atlas-phase2-docx.md`
— both now carry a dated "Actual State" addendum reconciling their
historical claims against what shipped.
