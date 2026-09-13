# Atlas Phase 1 — Universal Document Viewer (Plan)

**Repo**: `C:\Users\Youssef\Documents\Projects\md-reader`
**Target artifact**: `release/Atlas-Setup-2.0.0.exe` (in-place upgrade over `MD Reader` via shared `appId: com.mdreader.app`)
**Mode**: VIEW-ONLY. No editing surfaces.
**Style invariants**: named exports, `useCallback`/`useMemo`, no `any`, lucide-react icons, BEM CSS, theme CSS variables only, lazy-loaded viewer libs.

## Session Progress

### PHASE 1 — COMPLETE ✅ (all 8 slices shipped)

**Final state**: tsc 0 errors, lint 0 errors, **151/151 unit tests**, **13/13 Playwright smoke**, installer `release/Atlas-Setup-2.0.0.exe` (123.16 MB), `appId: com.mdreader.app` preserved for in-place upgrade.

**Bundle delta vs Phase 0 baseline (G3)**: main +4.44%, gzip +1.56% — well under the 25% cap.

**13 formats supported (view-only)**: markdown, text, code, csv, tsv, xlsx, ods, docx, rtf, odt, pdf, pptx, odp.

### Branding Polish — COMPLETE ✅
- ✅ Atlas brand mark designed: Lucide globe (ISC) on indigo `#4F46E5` rounded-square tile (1024×1024 SVG, 180px radius)
- ✅ Master SVG at `assets/atlas.svg` — single source of truth for all icon outputs
- ✅ Build script `scripts/build-icon.mjs` — sharp rasterizes SVG → multi-size PNGs → `png-to-ico` packs `.ico` (sizes 16/24/32/48/64/128/256)
- ✅ Generated `build/icon.ico` (363.8 KB, 7-size multi-res) — used by app exe + ProgID `DefaultIcon` for ALL 43 file associations
- ✅ Generated `build/icon.png` (51.5 KB, 1024×1024) — fallback for non-Windows
- ✅ Replaced `public/favicon.svg` (was MD-branded blue square) with Atlas globe glyph
- ✅ `electron-builder.yml`: `win.icon: build/icon.ico` (was `.png`)
- ✅ **NSIS install-folder fix**: added `!macro preInit` to `build/installer.nsh` — clears stale `HKLM\Software\<APP_GUID>\InstallLocation` (legacy `C:\Program Files\MD Reader\Atlas\` from MD Reader), wipes `Software\MD Reader` keys (HKLM+HKCU, x64+x32 views), forces `$INSTDIR = $PROGRAMFILES64\Atlas` so upgrades land in `C:\Program Files\Atlas\`
- ✅ Final installer: `release/Atlas-Setup-2.0.0.exe` (119.26 MB, 125,055,568 bytes)
- ✅ SHA256: `4AB08012D719CE1C13394CA98EED4B215B66F144917114216673B2F6F69551BF`
- ✅ Verified: tsc 0, lint 0, 151/151 unit, 13/13 Playwright smoke

### Slice 8 — COMPLETE (Wave C: bundle gate + Playwright + installer)
- ✅ W5.1 `scripts/check-bundle.mjs` — reads `.sisyphus/baselines/atlas-phase0.json`, fails on >25% regression
- ✅ W5.2 `playwright.config.ts` + `tests/e2e/smoke.spec.ts` + `tests/e2e/fixtures/generate.mjs` (13 fixtures, all passing)
- ✅ W6.1 NSIS installer built — `Atlas-Setup-2.0.0.exe`, 123,157,594 bytes
- ✅ Fixture path corrected to `tests/e2e/fixtures/`; smoke spec asserts mounted viewer + filename + zero console errors

### Hotfixes (mid-Wave-C, applied to handle live bugs)
- ✅ `electron/main.cjs` — null guards in `sendFileToWindow`/`extractFilePath`/`open-file`; replaced `.split('.').pop()` with `lastIndexOf('.')` pattern
- ✅ Installed `react-responsive-carousel` via `--legacy-peer-deps` (peer of glide-data-grid)
- ✅ Swapped `html2canvas` → `html2canvas-pro` in `src/utils/export.ts` (modern CSS color support)
- ✅ Recent files CSS: added `.recent__open` (flex/gap), `.recent__time`, fixed `.recent__header`
- ✅ **Empty-viewer fix** in `src/index.css`: added `.content--viewer .preview-panel` flex chain + `.docx-viewer` styling — viewers no longer collapse to 0×0

### Slice 7 — COMPLETE (Wave B: ViewerRouter wiring + chrome generalization)
- ✅ `ViewerRouter` (Suspense + ErrorBoundary + lazy viewer) wired into `App.tsx`
- ✅ Conditional render: MD editor only when `file.format === 'markdown'`; all others go through `ViewerRouter`
- ✅ Sidebar/StatusBar generalized via `ViewerProvider` (nav items + stats discriminated union)
- ✅ G2 result: tsc 0, lint 0, **151/151 tests** ✅

### Wave A — COMPLETE (Slices 3-6 in parallel)
- ✅ **Slice 3** SpreadsheetViewer (xlsx + ods via SheetJS, glide-data-grid, theme bridge) + CsvViewer (papaparse) + tsv alias
- ✅ **Slice 4** CodeViewer (Shiki, theme bridge, extToLang map covering 30+ languages)
- ✅ **Slice 5** PdfViewer (pdfjs-dist v5, worker URL, page virtualization)
- ✅ **Slice 6** PptxViewer + OdpViewer + shared SlideDeck component
- ✅ G2 result: tsc 0, lint 0, **141/141 tests** ✅

### Slice 2 — COMPLETE (DocxViewer + RtfViewer + OdtViewer + lint debt cleared)
- ✅ W2.DOCX `DocxViewer.tsx` — lazy `docx-preview`, page count from `.docx-wrapper > section`, words from `innerText`, error fallback, cancellation flag
- ✅ W2.RTF `RtfViewer.tsx` — lazy `rtf.js@3.0.9` (full TS types built-in, `DocumentFacade.render() => Promise<HTMLElement[]>`), `RTFJS/WMFJS/EMFJS.loggingEnabled(false)`, error fallback
- ✅ W2.ODT `OdtViewer.tsx` — lazy `odf-kit/reader` subpath (`odtToHtml(Uint8Array, { fragment: true }) => string`), DOMPurify sanitization (FORBID script/style/iframe/object/embed + on*-handlers), error fallback
- ✅ Installed `odf-kit@0.13.4` + `dompurify`
- ✅ **Lint debt cleared**: split `ViewerContext.tsx` into 3 files
  - `viewerContextValue.ts` — context object + value type only
  - `useViewerContext.ts` — 4 consumer hooks (`useNavItems`, `useSetNavItems`, `useViewerStats`, `useSetViewerStats`)
  - `ViewerContext.tsx` — `ViewerProvider` only (component file fast-refresh-clean)
  - Reset-on-filePath uses React 19 setState-during-render idiom (avoids `react-hooks/refs` AND `react-hooks/set-state-in-effect`)
  - Updated imports in `TextViewer.tsx`, `MarkdownViewer.tsx`, `__tests__/ViewerContext.test.tsx`
- **G2 result**: tsc 0 errors ✅, lint 0 errors ✅, **138/138 tests pass** ✅ (was 138 baseline going in — no regressions)

### Slice 1 — COMPLETE (Markdown viewer + full Atlas rebrand)
- ✅ W2.MD `MarkdownViewer.tsx` (54 lines) + fixture `sample.md`
- ✅ W2.TXT `TextViewer.tsx` (101 lines, virtualized via react-window v2 List aliased to FixedSizeList) + fixture `sample.txt`
- ✅ W4.1 `package.json` (name=atlas, version=2.0.0) + `electron-builder.yml` (productName=Atlas, expanded fileAssociations, appId KEPT)
- ✅ W4.2 UI strings rebranded across `index.html`, `WelcomeScreen.tsx`, `Toolbar.tsx`, `constants.ts`, `utils/export.ts`, `README.md`
- ✅ W4.3 `src/lib/migrateLocalStorage.ts` + 4 tests; hooks updated (`useTheme`, `useAutosave`, `useRecentFiles`, `useFontSize`); `runMigration()` injected in `main.tsx` before `createRoot`
- ✅ W4.4 `build/installer.nsh` rewritten — 5 ProgIDs (Atlas.Document/Spreadsheet/Presentation/Pdf/Code) × 43 extensions, 245 lines, legacy MDReader cleanup intact
- ✅ Fixed regression: installed missing `@testing-library/dom` peer dep

### Next session entry point — Phase 2 (DOCX editing)
Phase 1 is fully shipped. Move to `.sisyphus/plans/atlas-phase2-docx.md` — hand-coded MS-Word-parity DOCX editing (no SDKs, no bundles): OOXML WordprocessingML parser+serializer, contenteditable rich-text engine, undo stack, paste-from-Word sanitizer, color pickers, formatting toolbar. Ships as `Atlas-Setup-3.0.0.exe`.

### Phase 2 plan
- See `.sisyphus/plans/atlas-phase2-docx.md` (Waves A–F, hand-coded DOCX editing parity, ships as `Atlas-Setup-3.0.0.exe`)

---

## Locked Decisions (user-confirmed)

- **Code syntax themes**: per-theme mapping — each of the 5 app themes maps to a corresponding Shiki theme.
- **PPTX/ODP rendering**: canvas thumbnails sidebar + DOM detail pane (text selectable in main view).
- **PDF export**: when current file is PDF, "Export to PDF" saves a copy of original bytes (preserve vector). Other formats rasterize via existing jspdf+html2canvas pipeline.
- **ODT library**: `odf-kit` (best available in-browser; Phase 1 best-effort fidelity).
- **Installer upgrade test**: manual on user's machine.

## Locked Library Stack

- Markdown: existing pipeline (marked + react-markdown + mermaid)
- DOCX: `docx-preview`
- RTF: `rtf.js`
- ODT: `odf-kit`
- XLSX/ODS: `xlsx` (SheetJS, native ODS support)
- CSV/TSV: `papaparse`
- Spreadsheet grid: `@glideapps/glide-data-grid`
- PDF: `react-pdf` (pdfjs-dist worker via `new URL('pdfjs-dist/build/pdf.worker.min.mjs', import.meta.url)`)
- PPTX/ODP: `jszip` + custom XML parser
- Code syntax highlighting: `shiki`
- Code language detection: extension map first, `highlight.js` auto-detect fallback
- All viewer libs lazy-loaded via dynamic `import()`

## Architecture Contracts (locked)

- New dirs: `src/formats/` (detect.ts, types.ts, registry.ts), `src/viewers/`
- `FormatId = 'markdown' | 'docx' | 'xlsx' | 'pptx' | 'pdf' | 'csv' | 'tsv' | 'text' | 'code' | 'odt' | 'ods' | 'odp' | 'rtf' | 'unknown'`
- `LoadedFile = { kind: 'text'; content: string; path: string; format: FormatId } | { kind: 'binary'; content: ArrayBuffer; path: string; format: FormatId }`
- Detection: extension map first, then magic-byte sniff (PDF=`25504446`, ZIP=`504B0304` → inspect inner OOXML/ODF, RTF=`7B5C72746`)
- Electron IPC: add new `dialog:openFileBinary` channel returning ArrayBuffer (do NOT mutate existing text channel)
- Single `ViewerRouter` dispatches by `FormatId` to lazy-loaded viewer via Suspense + ErrorBoundary
- Sidebar generalized via `NavItem[]` published by viewers through context
- StatusBar consumes `ViewerStats` discriminated union via context
- ExportMenu format-aware
- Universal shortcuts: Ctrl+P export PDF, Ctrl+E export menu, PgUp/PgDn navigate

## Rebrand

- `package.json`: name `md-reader` → `atlas`, version `1.0.1` → `2.0.0`
- `electron-builder.yml`: productName → `Atlas`, artifactName → `Atlas-Setup-${version}.${ext}`, shortcutName → `Atlas`, expand fileAssociations to ~30 extensions
- KEEP `appId: com.mdreader.app` for in-place upgrade
- `build/installer.nsh`: 5 ProgIDs (Atlas.Document, Atlas.Spreadsheet, Atlas.Presentation, Atlas.Pdf, Atlas.Code)
- localStorage migration: `md-reader-*` → `atlas-*` (idempotent, one-time)
- UI text: WelcomeScreen, Toolbar, README, index.html title

---

# Tasks

## Wave 0 — Baseline (sequential)

### W0.1 — Repo audit + baseline green
- **Description**: Run `npm ci`, `npm run lint`, `npx tsc --noEmit -p tsconfig.app.json`, `npm run build`. Record current bundle size, lint warning count, and tsc errors as baseline.
- **Files**: read-only audit; produce `.sisyphus/baselines/atlas-phase0.json`.
- **Agent**: Category `quick`, Skills `[]`.
- **QA Scenario**:
  - Tool: `bash` from project root.
  - Steps: run each command sequentially.
  - Expected: all four commands exit 0; `.sisyphus/baselines/atlas-phase0.json` exists with keys `{ lintWarnings, lintErrors, tscErrors, distSizeKB, mainChunkKB }`.
  - Pass condition: file written, all commands exit 0.

### W0.2 — Vitest + RTL bootstrap
- **Description**: Add `vitest`, `@testing-library/react`, `@testing-library/jest-dom`, `jsdom` to devDeps; add `vitest.config.ts` mirroring `vite.config.ts` aliases; add `npm test` script. Add `src/__tests__/smoke.test.ts` asserting `expect(1+1).toBe(2)`.
- **Files**: `package.json`, `vitest.config.ts` (new), `src/__tests__/smoke.test.ts` (new), `tsconfig.app.json` (add vitest types).
- **Agent**: Category `quick`, Skills `[]`.
- **QA Scenario**:
  - Tool: `bash`.
  - Steps: `npm install` then `npm test`.
  - Expected: vitest runs smoke test, exits 0, output shows `1 passed`.
  - Pass condition: `npm test` exits 0 with smoke test green.

---

## Wave 1 — Foundation

### W1.1 — Format type system
- **Description**: Create `src/formats/types.ts` exporting:
  - `FormatId` union (14 values listed above)
  - `LoadedFile` discriminated union on `kind: 'text' | 'binary'`
  - `ViewerStats` discriminated union: `{kind:'markdown', words, headings}`, `{kind:'spreadsheet', sheet, rows, cols}`, `{kind:'pdf', page, pageCount}`, `{kind:'slides', slide, slideCount}`, `{kind:'code', language, lines}`, `{kind:'text', lines, chars}`, `{kind:'document', words, pages}`
  - `NavItem = { id: string; label: string; icon?: LucideIcon; level?: number; onSelect: () => void }`
  - `assertNever(x: never): never` helper
- **Files**: `src/formats/types.ts` (new), `src/formats/__tests__/types.test.ts` (new).
- **Agent**: Category `ultrabrain`, Skills `[]`.
- **QA Scenario**:
  - Tool: `bash`.
  - Steps: `npx tsc --noEmit -p tsconfig.app.json && npm test -- types`.
  - Expected: tsc exits 0; exhaustiveness test fails to compile if a `FormatId` case is omitted from a switch (guard via `assertNever`).
  - Pass condition: tsc green AND types test green.

### W1.2 — Detection module + tests (TDD-first)
- **Description**: TDD. First write `src/formats/__tests__/detect.test.ts` covering: 40+ code extensions, all doc extensions, magic bytes (PDF `%PDF`, ZIP `PK\x03\x04` then inspect inner `word/`/`xl/`/`ppt/`/`mimetype` payload to disambiguate ooxml/odf, RTF `{\rtf`), unknown fallback. Then implement `src/formats/detect.ts`:
  - `detectByExtension(path: string): FormatId`
  - `detectByMagic(buf: ArrayBuffer): FormatId`
  - `detectFormat(path: string, buf?: ArrayBuffer): FormatId`
- **Files**: `src/formats/detect.ts` (new), `src/formats/__tests__/detect.test.ts` (new).
- **Agent**: Category `deep`, Skills `[]`.
- **QA Scenario**:
  - Tool: `bash`.
  - Steps: `npm test -- detect`.
  - Expected: ≥30 test cases pass; specifically docx vs xlsx vs pptx vs odt vs ods vs odp all distinguished from raw ZIP magic via inner archive inspection.
  - Pass condition: every `FormatId` (except `unknown`) reachable through at least one test case; coverage report shows 100% of branches in detect.ts.

### W1.3 — Lazy viewer registry
- **Description**: Create `src/formats/registry.ts` exporting `viewerRegistry: Record<FormatId, () => Promise<ComponentType<{ file: LoadedFile }>>>` using dynamic imports. Each entry returns a Promise of the named export. `unknown` returns `UnknownViewer`.
- **Files**: `src/formats/registry.ts` (new), `src/viewers/UnknownViewer.tsx` (new placeholder), `src/formats/__tests__/registry.test.ts` (new).
- **Agent**: Category `quick`, Skills `[]`.
- **QA Scenario**:
  - Tool: `bash`.
  - Steps: `npm test -- registry && npx tsc --noEmit -p tsconfig.app.json`.
  - Expected: snapshot test asserts `Object.keys(viewerRegistry).sort()` equals all FormatId values sorted; tsc green.
  - Pass condition: keys exhaustive, tsc green.

### W1.4 — Generalized binary IPC (dialog + path) and argv path detection
- **Description**: Generalize Electron's path-based file flow so OS "Open with" works for all 13 formats, not just markdown.
  1. **Add `dialog:openFileBinary` IPC** in `electron/main.cjs` — opens picker with universal filter (all 30+ extensions), reads via `fs.promises.readFile`, returns `{ canceled: boolean, path: string, buffer: ArrayBuffer }` (transfer Buffer as `Uint8Array.buffer.slice(byteOffset, byteOffset+byteLength)`).
  2. **Add `file:readBinaryByPath` IPC** in `electron/main.cjs` — accepts `filePath: string`, validates `fs.existsSync`, reads buffer, returns `{ path, buffer: ArrayBuffer }`. This unblocks OS "Open with" for non-markdown formats.
  3. **Generalize `extractFilePath(argv)`** in `electron/main.cjs` — replace the `/\.(md|markdown)$/i` regex with a universal known-extension whitelist (the same 30+ extensions used in `electron-builder.yml` fileAssociations). Argv detection no longer markdown-only.
  4. **Generalize `sendFileToWindow`** — instead of always calling `readMarkdownFile`, it now sends a `file-opened-path` event carrying just the absolute path; the renderer (W1.5 useFileHandler) decides text-vs-binary based on `detectByExtension`. Keep existing `file-opened` event firing for `.md`/`.markdown` so existing markdown listener stays a regression-free fallback during transition (remove in W3.5).
  5. **Generalize `get-initial-file`** — return `{ path: string } | null` instead of full markdown payload; renderer reads via the new IPCs.
  6. **Expose new methods** in `electron/preload.cjs`: extend the existing `electronAPI` object (NOT `window.api`) with `openFileBinary()`, `readBinaryByPath(path)`, and rename `getInitialFile` to return `{path}|null`. Existing `openFileDialog`, `openFileByPath`, `saveFile`, `onFileOpened`, `setTheme` stay intact.
  7. **Extend `src/electron.d.ts`** (existing file at `src/electron.d.ts`, NOT `src/types/electron.d.ts`) with new method signatures and a new `onFileOpenedPath` callback.
- **Files**: `electron/main.cjs`, `electron/preload.cjs`, `src/electron.d.ts` (extend existing).
- **Agent**: Category `unspecified-high`, Skills `[]`.
- **QA Scenario**:
  - Tool: manual via `npm run electron:dev` then DevTools console (NOT `npm run dev` which is Vite-only); plus argv smoke test via `& "release\win-unpacked\Atlas.exe" "C:\path\to\sample.pdf"` after W6.1 builds.
  - Steps:
    1. Launch `npm run electron:dev`. In DevTools console run `await window.electronAPI.openFileBinary()` and pick a PDF — confirm `new Uint8Array((await ...).buffer.slice(0,4))` is `[0x25, 0x50, 0x44, 0x46]`.
    2. In DevTools run `await window.electronAPI.readBinaryByPath('C:\\Users\\Youssef\\Documents\\Projects\\md-reader\\README.md')` — confirm returned `buffer.byteLength > 0` and `new TextDecoder().decode(new Uint8Array(buf.buffer).slice(0,2))` starts with `#` (markdown heading byte). This proves path-based binary read works on a Wave-1-available file without needing any future fixture.
    3. Confirm the existing markdown open dialog (`await window.electronAPI.openFileDialog()`) still works (regression check).
  - Expected: PDF magic bytes returned in step 1; non-empty buffer with markdown content in step 2; markdown dialog still works in step 3.
  - Pass condition: all 3 checks pass.

### W1.5 — `useFileHandler` returns `LoadedFile` (path + dialog + drag-drop)
- **Description**: Refactor `src/hooks/useFileHandler.ts` to a unified loader that produces `LoadedFile` from any of: dialog open, OS "Open with" argv path, drag-drop. Algorithm:
  1. Define `loadFromPath(absPath)`: call `detectByExtension(absPath)`. If text-class (md, csv, tsv, text, code), call existing `window.electronAPI.openFileByPath(absPath)` and wrap as `{kind:'text', content, path, format}`. If binary-class (pdf, docx, xlsx, pptx, odt, ods, odp, rtf), call new `window.electronAPI.readBinaryByPath(absPath)` and wrap as `{kind:'binary', content: arraybuf, path, format}`. If `unknown`, call `readBinaryByPath` then `detectByMagic` to refine.
  2. Define `loadFromDialog()`: call `window.electronAPI.openFileBinary()` (universal dialog), then route through `loadFromPath`-like switch using both extension and magic-byte detection.
  3. Wire boot: subscribe to new `onFileOpenedPath(callback)` event so OS "Open with" routes through `loadFromPath`. Also call `getInitialFile()` at boot — if it returns `{path}`, run `loadFromPath`.
  4. Drag-drop: when a `File` is dropped, derive path from `file.path` (Electron exposes this on dropped files), then call `loadFromPath`.
  5. On any success, push to recent files list with `{path, format, openedAt}` shape (format icon hint added later in Sidebar).
- **Files**: `src/hooks/useFileHandler.ts`, `src/hooks/__tests__/useFileHandler.test.ts` (new).
- **Agent**: Category `unspecified-high`, Skills `[]`.
- **QA Scenario**:
  - Tool: `bash` for unit tests; manual via `npm run electron:dev` for live smoke (NOT `npm run dev`).
  - Steps:
    1. `npm test -- useFileHandler && npx tsc --noEmit -p tsconfig.app.json`.
    2. `npm run electron:dev`; drag a `.md` onto window — opens as text.
    3. In same dev session, click File→Open and select a PDF — opens as binary.
    4. Quit, then from PowerShell at the repo root: `npx electron . "C:\Users\Youssef\Documents\Projects\md-reader\README.md"` — Electron launches with README.md already loaded as text via the argv path. (This uses the dev Electron binary that ships in `node_modules/.bin` after `npm install`, so it is runnable in Wave 1 without any packaged Atlas.exe.) Optional post-W6.1 check: re-run with `& "release\win-unpacked\Atlas.exe" "<any-fixture>"`.
  - Expected: hook return type narrows to `LoadedFile | null`; unit tests with mocked `window.electronAPI` cover text path, binary path, unknown→magic-byte path, and `onFileOpenedPath` boot path; manual smoke shows text/binary/argv all work.
  - Pass condition: tests green, tsc green, all 3 manual smoke paths work.

### W1.6 — ViewerRouter skeleton
- **Description**: Create `src/components/ViewerRouter.tsx` accepting `{ file: LoadedFile }`, looking up `viewerRegistry[file.format]`, rendering inside `<Suspense fallback={<ViewerLoading/>}>` wrapped in `<ViewerErrorBoundary>`. Mount in App.tsx behind a feature flag (do NOT replace existing render path yet — that's W3.5).
- **Files**: `src/components/ViewerRouter.tsx` (new), `src/components/ViewerLoading.tsx` (new), `src/components/ViewerErrorBoundary.tsx` (new), `src/components/__tests__/ViewerRouter.test.tsx` (new).
- **Agent**: Category `unspecified-high`, Skills `[]`.
- **QA Scenario**:
  - Tool: `bash`.
  - Steps: `npm test -- ViewerRouter`.
  - Expected: RTL test renders with mock registry, asserts loading state then resolves; throws-and-catches error from a failing viewer.
  - Pass condition: tests green.

### Gate G1 (after Wave 1)
- Tool: `bash`.
- Steps: `npm test && npm run lint && npx tsc --noEmit -p tsconfig.app.json`.
- Expected: all exit 0.
- Failure action: dispatch `quick`-category fix task targeting failing file before Wave 2.

---

## Wave 2 — Viewers + Rebrand (PARALLEL)

> Pattern for every viewer task: 1 fixture in `src/viewers/__fixtures__/`, viewer component as named export, lazy-imported by registry, receives `{ file: LoadedFile }`, publishes `NavItem[]` via `useSetNavItems()`, publishes `ViewerStats` via `useSetViewerStats()`, all colors/spacing via theme CSS variables (no hardcoded colors).

### W2.X1 — Shiki theme bridge
- **Description**: `src/viewers/shared/shikiTheme.ts` exporting `getShikiThemeForAppTheme(themeId: ThemeId): BundledTheme`. The locked app theme IDs (from `src/types.ts` and `src/index.css`) are exactly: `light`, `dark`, `sepia`, `nord`, `dracula`. Map: `light→github-light`, `dark→github-dark`, `sepia→solarized-light` (warm cream-on-brown matches sepia palette), `nord→nord`, `dracula→dracula`. Use exhaustive switch with `assertNever` to enforce all 5 IDs covered.
- **Files**: `src/viewers/shared/shikiTheme.ts` (new), `src/viewers/shared/__tests__/shikiTheme.test.ts` (new).
- **Agent**: Category `quick`, Skills `[]`.
- **QA Scenario**:
  - Tool: `bash`.
  - Steps: `npm test -- shikiTheme`.
  - Expected: each of the 5 app theme IDs (`light|dark|sepia|nord|dracula`) maps to a real BundledTheme name accepted by Shiki (`github-light|github-dark|solarized-light|nord|dracula`); `assertNever` guards exhaustiveness.
  - Pass condition: 5 mapping cases green; `npx tsc` fails if a `ThemeId` is added without updating the switch.

### W2.X2 — Nav/Stats context providers
- **Description**: `src/viewers/shared/ViewerContext.tsx` exposing `<ViewerProvider>`, `useNavItems()`, `useSetNavItems()`, `useViewerStats()`, `useSetViewerStats()`. Cleared automatically when `file.path` changes.
- **Files**: `src/viewers/shared/ViewerContext.tsx` (new), `src/viewers/shared/__tests__/ViewerContext.test.tsx` (new).
- **Agent**: Category `unspecified-high`, Skills `[]`.
- **QA Scenario**:
  - Tool: `bash`.
  - Steps: `npm test -- ViewerContext`.
  - Expected: setting nav items then changing file resets to empty; subscribers re-render.
  - Pass condition: 3 cases green (set, get, reset-on-file-change).

### W2.MD — MarkdownViewer
- **Description**: `src/viewers/MarkdownViewer.tsx` wrapping existing markdown components (Editor preview pipeline). Publish headings as `NavItem[]` (clicking scrolls to heading). Publish stats `{kind:'markdown', words, headings}`. Mermaid still rendered. Also creates a small dedicated fixture under `src/viewers/__fixtures__/sample.md` so the W3.5 "13 fixtures" matrix is satisfied without depending on the project README.
- **Files**: `src/viewers/MarkdownViewer.tsx` (new), `src/viewers/__fixtures__/sample.md` (new — covers headings, lists, code fence, mermaid block, table, image), reuses existing markdown components.
- **Agent**: Category `unspecified-high`, Skills `[]`.
- **QA Scenario**:
  - Tool: manual Electron via `npm run electron:dev`.
  - Steps: launch app via `npm run electron:dev`, open `src/viewers/__fixtures__/sample.md` through ViewerRouter feature flag.
  - Expected: rendered identically to current MD reader; sidebar shows heading outline; stats bar shows word count; mermaid block renders.
  - Pass condition: visual parity confirmed by user.

### W2.DOCX — DocxViewer (`docx-preview`)
- **Description**: Lazy-import `docx-preview`, call `parseAsync` + `renderDocument` (NOT `renderAsync`, to avoid built-in CSS), filter STYLE nodes out, render into scoped `.viewer-docx` container. Publish heading outline as NavItems; stats `{kind:'document', words, pages}`.
- **Files**: `src/viewers/DocxViewer.tsx` (new), `src/viewers/__fixtures__/sample.docx` (new), `src/viewers/__styles__/viewer-docx.css` (new scoped styles).
- **Agent**: Category `unspecified-high`, Skills `[]`.
- **QA Scenario**:
  - Tool: manual Electron via `npm run electron:dev`.
  - Steps: launch via `npm run electron:dev`, open `src/viewers/__fixtures__/sample.docx` via app.
  - Expected: headings, lists, tables, inline images, bold/italic visible; theme colors apply (not docx-preview's defaults).
  - Pass condition: visual confirmation + DevTools `document.head.querySelectorAll('style')` length unchanged after open (no `<style>` injection by docx-preview).

### W2.RTF — RtfViewer (`rtf.js`)
- **Description**: Lazy-import `rtf.js` (RTFJS, WMFJS, EMFJS), disable logging, build `RTFJS.Document(arrayBuffer)`, render nodes into scoped container.
- **Files**: `src/viewers/RtfViewer.tsx` (new), `src/viewers/__fixtures__/sample.rtf` (new).
- **Agent**: Category `unspecified-high`, Skills `[]`.
- **QA Scenario**:
  - Tool: manual Electron via `npm run electron:dev`.
  - Steps: launch via `npm run electron:dev`, open `src/viewers/__fixtures__/sample.rtf`.
  - Expected: bold, italic, color, basic tables visible.
  - Pass condition: visual confirmation.

### W2.ODT — OdtViewer (`odf-kit`)
- **Description**: Lazy-import `odf-kit`, call `odtToHtml(uint8array)`, sanitize HTML (DOMPurify if needed), render in `.viewer-odt` scoped container.
- **Files**: `src/viewers/OdtViewer.tsx` (new), `src/viewers/__fixtures__/sample.odt` (new).
- **Agent**: Category `unspecified-high`, Skills `[]`.
- **QA Scenario**:
  - Tool: manual Electron via `npm run electron:dev`.
  - Steps: launch via `npm run electron:dev`, open `src/viewers/__fixtures__/sample.odt`.
  - Expected: headings, paragraphs, lists visible; no script execution from injected HTML.
  - Pass condition: visual confirmation + no console security warnings.

### W2.XLSX — SpreadsheetViewer (`xlsx` + `@glideapps/glide-data-grid`)
- **Description**: Parse with SheetJS `XLSX.read(buf, {cellFormula:true, cellStyles:true, xlfn:true, sheetStubs:true})`. Sheet tabs as `NavItem[]`. Active sheet rendered via Glide Data Grid with virtualization, freezeColumns: 0 (configurable), header row, copy-to-clipboard via `getCellsForSelection`, sort-on-header-click. Stats `{kind:'spreadsheet', sheet, rows, cols}`.
- **Files**: `src/viewers/SpreadsheetViewer.tsx` (new), `src/viewers/__fixtures__/sample.xlsx` (new with 10k+ rows test sheet), `src/viewers/__styles__/viewer-spreadsheet.css` (new with theme CSS var bridge to glide theme).
- **Agent**: Category `visual-engineering`, Skills `[frontend-ui-ux]`.
- **QA Scenario**:
  - Tool: manual Electron via `npm run electron:dev` + DevTools Performance tab.
  - Steps: launch via `npm run electron:dev`, open `src/viewers/__fixtures__/sample.xlsx`, scroll to row 10000, switch sheets, click header to sort, select range and Ctrl+C.
  - Expected: scroll fluid (≥30fps via Performance tab FPS meter), sort instant on <10k rows, clipboard contains tab-separated values (paste into a text editor to verify).
  - Pass condition: all interactions work; Performance tab shows no main-thread block >100ms during scroll.

### W2.ODS — uses SpreadsheetViewer
- **Description**: Registry maps `ods` → `SpreadsheetViewer` (SheetJS handles ODS natively, no new component). Also creates a multi-sheet ODS fixture for QA so the W3.5 13-fixture matrix is satisfied.
- **Files**: `src/formats/registry.ts` (entry only), `src/viewers/__fixtures__/sample.ods` (new — produced via LibreOffice export of `sample.xlsx` from W2.XLSX, or hand-built minimal multi-sheet ODS).
- **Agent**: Category `quick`, Skills `[]`.
- **QA Scenario**:
  - Tool: manual Electron via `npm run electron:dev`.
  - Steps: launch via `npm run electron:dev`, open `src/viewers/__fixtures__/sample.ods`.
  - Expected: opens via SpreadsheetViewer, multiple sheets, formulas visible as values.
  - Pass condition: visual confirmation; fixture file exists at expected path.

### W2.CSV — CsvViewer (`papaparse` + Glide)
- **Description**: Stream-parse with PapaParse (`worker: true, header: false, skipEmptyLines: true`), render in Glide grid. Auto-detect delimiter for TSV reuse via `delimiter: ext === 'tsv' ? '\t' : undefined`. Stats `{kind:'spreadsheet', sheet:'data', rows, cols}`.
- **Files**: `src/viewers/CsvViewer.tsx` (new), `src/viewers/__fixtures__/sample.csv` (new), `src/viewers/__fixtures__/large.csv` (1M rows for stress test).
- **Agent**: Category `visual-engineering`, Skills `[frontend-ui-ux]`.
- **QA Scenario**:
  - Tool: manual Electron via `npm run electron:dev` + DevTools Performance tab.
  - Steps: launch via `npm run electron:dev`, open `src/viewers/__fixtures__/sample.csv` then `src/viewers/__fixtures__/large.csv`.
  - Expected: small file instant; large file shows progress, parses in worker without blocking UI (verify via Performance tab no main-thread block >100ms).
  - Pass condition: large file reachable, scroll fluid, Performance tab confirms worker thread is doing parsing.

### W2.TSV — uses CsvViewer
- **Description**: Registry maps `tsv` → `CsvViewer` with delimiter prop. Also creates a TSV fixture for QA so the W3.5 13-fixture matrix is satisfied.
- **Files**: `src/formats/registry.ts` (entry only), `src/viewers/__fixtures__/sample.tsv` (new — 3 columns × 5 rows tab-delimited, matches `sample.csv` content for parity check).
- **Agent**: Category `quick`, Skills `[]`.
- **QA Scenario**:
  - Tool: manual Electron via `npm run electron:dev`.
  - Steps: launch via `npm run electron:dev`, open `src/viewers/__fixtures__/sample.tsv`.
  - Expected: parses as tab-delimited, columns correct.
  - Pass condition: visual confirmation; fixture file exists at expected path.

### W2.TXT — TextViewer
- **Description**: Plain text rendered in virtualized list (use `react-window` if not present, else custom virtualization). Stats `{kind:'text', lines, chars}`.
- **Files**: `src/viewers/TextViewer.tsx` (new), `src/viewers/__fixtures__/sample.txt` (new), `src/viewers/__fixtures__/large.log` (50MB).
- **Agent**: Category `quick`, Skills `[]`.
- **QA Scenario**:
  - Tool: manual Electron via `npm run electron:dev`; `bash` PowerShell for line count cross-check.
  - Steps: launch via `npm run electron:dev`, open `src/viewers/__fixtures__/large.log`. Cross-check line count: `(Get-Content -LiteralPath src/viewers/__fixtures__/large.log | Measure-Object -Line).Lines` in PowerShell.
  - Expected: opens within 2s, scroll smooth, StatusBar line count matches PowerShell line count.
  - Pass condition: no UI freeze, line count exactly matches.

### W2.CODE — CodeViewer (`shiki` + `highlight.js` detect)
- **Description**: Build `extToLang` map covering 40+ extensions (.ts, .tsx, .js, .jsx, .py, .rs, .go, .java, .c, .cpp, .h, .hpp, .cs, .rb, .php, .sh, .ps1, .sql, .html, .css, .scss, .json, .yaml, .yml, .toml, .ini, .xml, .vue, .svelte, .swift, .kt, .scala, .clj, .ex, .erl, .lua, .pl, .r, .dart, .nim). Unknown ext → `hljs.highlightAuto(code, [topN])` to detect. Render via Shiki using `getShikiThemeForAppTheme(currentTheme)`. Stats `{kind:'code', language, lines}`. Symbol outline via simple regex per language family (function/class/def/fn) → NavItems.
- **Files**: `src/viewers/CodeViewer.tsx` (new), `src/viewers/extToLang.ts` (new), `src/viewers/__fixtures__/sample.ts`, `sample.py`, `sample.rs` (new).
- **Agent**: Category `visual-engineering`, Skills `[frontend-ui-ux]`.
- **QA Scenario**:
  - Tool: `bash` for unit tests; manual Electron via `npm run electron:dev` for theme-switch smoke.
  - Steps: `npm test -- CodeViewer extToLang`; then `npm run electron:dev` and open each fixture, switch app theme via theme picker.
  - Expected: ≥40 extensions covered in unit test; opening unknown extension `.xyz` triggers hljs detect; switching theme re-renders Shiki with matching theme.
  - Pass condition: extension test green; live theme switch re-highlights code.

### W2.PDF — PdfViewer (`react-pdf`)
- **Description**: Lazy-import `react-pdf`, configure worker via `pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('pdfjs-dist/build/pdf.worker.min.mjs', import.meta.url).toString()`. Render `<Document file={arrayBuffer}>` with `<Page>` per page. PgUp/PgDn navigation. Outline extracted from `pdfDocument.getOutline()` → recursive NavItems. Stats `{kind:'pdf', page, pageCount}`. Verify worker loads in packaged Electron build (not just dev).
- **Files**: `src/viewers/PdfViewer.tsx` (new), `src/viewers/__fixtures__/sample.pdf` (new).
- **Agent**: Category `deep`, Skills `[]`.
- **QA Scenario**:
  - Tool: manual Electron via `npm run electron:dev`; `bash` for build + asset audit.
  - Steps: 1) `npm run electron:dev`, open `src/viewers/__fixtures__/sample.pdf`, navigate pages with PgUp/PgDn, click outline entry. 2) `npm run build` then `Get-ChildItem dist/assets/ -Filter '*pdf.worker*'` to confirm worker chunk emitted.
  - Expected: pages render at 100% width, outline navigates, stats update on page change. Worker chunk present in `dist/assets/`.
  - Pass condition: visual confirmation + `Get-ChildItem` returns ≥1 file matching `pdf.worker`.

### W2.PPTX — PptxViewer (custom jszip + XML)
- **Description**: Lazy-import `jszip`, extract `.pptx`. Parse `ppt/presentation.xml` for slide list, then per slide `ppt/slides/slideN.xml` for text runs (`<a:t>` elements) and shape positions. Extract images from `ppt/media/`. Render slide rail (left) as canvas thumbnails (low-res rasterized via canvas-text), main pane as DOM detail of selected slide (text selectable, images positioned absolutely scaled to viewport). Slides as NavItems; stats `{kind:'slides', slide, slideCount}`.
- **Files**: `src/viewers/PptxViewer.tsx` (new), `src/viewers/shared/SlideDeck.tsx` (new shared with ODP), `src/viewers/__fixtures__/sample.pptx` (new).
- **Agent**: Category `ultrabrain`, Skills `[]`.
- **QA Scenario**:
  - Tool: manual Electron via `npm run electron:dev` + DevTools Console.
  - Steps: launch via `npm run electron:dev`, open `src/viewers/__fixtures__/sample.pptx`, click slide thumbnails, PgUp/PgDn navigate, select text in main pane.
  - Expected: thumbnails render for all slides; main pane shows text + images; text selectable; navigation works.
  - Pass condition: visual confirmation, text selection works, DevTools Console has 0 XML parse errors.

### W2.ODP — OdpViewer (sibling of PPTX)
- **Description**: Same approach against ODF schema. Parse `content.xml` for `<draw:page>` elements, text in `<text:p>`, images from `Pictures/`. Reuse `<SlideDeck>` from W2.PPTX.
- **Files**: `src/viewers/OdpViewer.tsx` (new), `src/viewers/__fixtures__/sample.odp` (new).
- **Agent**: Category `ultrabrain`, Skills `[]`.
- **QA Scenario**:
  - Tool: manual Electron via `npm run electron:dev`.
  - Steps: launch via `npm run electron:dev`, open `src/viewers/__fixtures__/sample.odp`.
  - Expected: slides shown via shared `<SlideDeck>` component.
  - Pass condition: visual confirmation.

### W4.1 — Rebrand: build configs
- **Description**: `package.json`: name `md-reader` → `atlas`, version → `2.0.0`. `electron-builder.yml`: `productName: Atlas`, `artifactName: Atlas-Setup-${version}.${ext}`, `nsis.shortcutName: Atlas`, expand `fileAssociations` to ~30 extensions (md, markdown, mdown, mkd, docx, rtf, odt, xlsx, xlsm, ods, csv, tsv, pptx, odp, pdf, txt, log, json, yaml, yml, toml, ini, js, ts, tsx, jsx, py, rs, go, java, c, cpp, h, hpp, cs, rb, php, sh, ps1, sql, html, css, scss). KEEP `appId: com.mdreader.app` unchanged.
- **Files**: `package.json`, `electron-builder.yml`.
- **Agent**: Category `unspecified-high`, Skills `[git-master]`.
- **QA Scenario**:
  - Tool: `bash` (PowerShell).
  - Steps:
    1. `npm run build; if ($?) { npx electron-builder --dir --win }`.
    2. `Test-Path release/win-unpacked/Atlas.exe` — must be `True`.
    3. `Select-String -Path electron-builder.yml -Pattern '^appId:\s*com\.mdreader\.app' -SimpleMatch:$false` — must return ≥1 match (proves source config still uses unchanged `appId`).
    4. `Select-String -Path electron-builder.yml -Pattern '^productName:\s*Atlas' -SimpleMatch:$false` — must return ≥1 match.
    5. `(Get-Item release/win-unpacked/Atlas.exe).VersionInfo.ProductName` — should report `Atlas`.
  - Expected: build produces `Atlas.exe` (NOT `MD Reader.exe`); `electron-builder.yml` source still has `appId: com.mdreader.app`; Atlas.exe ProductName VersionInfo is `Atlas`.
  - Pass condition: all 5 checks pass.

### W4.2 — Rebrand: UI strings
- **Description**: Replace "MD Reader" → "Atlas" in `src/components/WelcomeScreen.tsx`, `src/components/Toolbar.tsx`, `index.html` `<title>`, `README.md`. Update WelcomeScreen body to mention multi-format support (Word/Excel/PowerPoint/PDF/code/markdown).
- **Files**: `src/components/WelcomeScreen.tsx`, `src/components/Toolbar.tsx`, `index.html`, `README.md`.
- **Agent**: Category `writing`, Skills `[]`.
- **QA Scenario**:
  - Tool: `bash` (PowerShell).
  - Steps: `Select-String -Path src\*,src\**\*,index.html,README.md -Pattern 'MD Reader' -SimpleMatch` (or use the `grep` tool with pattern `"MD Reader"` against `src/`, `index.html`, `README.md`).
  - Expected: zero hits except `src/lib/migrateLocalStorage.ts` (W4.3 needs `md-reader-` prefix string for migration).
  - Pass condition: only allowed-list hits remain.

### W4.3 — localStorage migration
- **Description**: `src/lib/migrateLocalStorage.ts` exports `runMigration()` called once at app boot in `App.tsx` `useEffect`. Logic: if `localStorage.getItem('atlas-migration-v1') === '1'` skip. Otherwise enumerate keys, for each starting with `md-reader-`, copy to `atlas-` prefix only if `atlas-` equivalent absent; delete old. Set flag.
- **Files**: `src/lib/migrateLocalStorage.ts` (new), `src/App.tsx` (call at boot), `src/lib/__tests__/migrateLocalStorage.test.ts` (new).
- **Agent**: Category `quick`, Skills `[]`.
- **QA Scenario**:
  - Tool: `bash`.
  - Steps: `npm test -- migrateLocalStorage`.
  - Expected: test cases — (a) fresh state with md-reader-* keys → migrated, flag set; (b) re-run with flag → no-op; (c) atlas-* already exists → md-reader-* removed without clobber; (d) no md-reader-* keys → flag set, no errors.
  - Pass condition: 4 cases green; idempotent.

### W4.4 — installer.nsh ProgIDs
- **Description**: Expand `build/installer.nsh`: define 5 ProgIDs:
  - `Atlas.Document` for `.docx, .odt, .rtf, .md, .markdown, .mdown, .mkd, .txt, .log`
  - `Atlas.Spreadsheet` for `.xlsx, .xlsm, .ods, .csv, .tsv`
  - `Atlas.Presentation` for `.pptx, .odp`
  - `Atlas.Pdf` for `.pdf`
  - `Atlas.Code` for `.json, .yaml, .yml, .toml, .ini, .js, .ts, .tsx, .jsx, .py, .rs, .go, .java, .c, .cpp, .h, .hpp, .cs, .rb, .php, .sh, .ps1, .sql, .html, .css, .scss`
  Register each under `RegisteredApplications` and `OpenWithProgids`. Call `SHChangeNotify`. Uninstall section reverses each ProgID.
- **Files**: `build/installer.nsh`.
- **Agent**: Category `deep`, Skills `[]`.
- **QA Scenario**:
  - Tool: `bash` (PowerShell) for build + registry inspection; manual click-test in Explorer for "Open with" UI.
  - Steps: 1) `npx electron-builder --win nsis --x64`. 2) Run produced installer. 3) `Get-ItemProperty -Path HKLM:\SOFTWARE\RegisteredApplications`. 4) `cmd /c "assoc .docx"`. 5) Right-click `.docx`, `.xlsx`, `.pdf`, `.pptx`, `.md` in Explorer → check Open With menu.
  - Expected: Atlas appears in RegisteredApplications; `assoc .docx` returns `.docx=Atlas.Document`; "Open with" menu shows Atlas for each of the 5 sample extensions.
  - Pass condition: at least 5 sample extensions verified to show Atlas in Open With.

### Gate G2 (after Wave 2)
- Tool: `bash`.
- Steps: `npm test && npm run lint && npx tsc --noEmit -p tsconfig.app.json`.
- Expected: all exit 0.
- Failure action: dispatch `quick` fix tasks per failing file.

---

## Wave 3 — Integration

### W3.1 — Sidebar generalize
- **Description**: Refactor `src/components/Sidebar.tsx` to consume `useNavItems()`. Render with lucide icon, label, level-based indent (CSS `padding-left: calc(var(--nav-indent, 12px) * level)`). Click → `onSelect()`. Empty state when no items.
- **Files**: `src/components/Sidebar.tsx`, `src/components/__tests__/Sidebar.test.tsx`.
- **Agent**: Category `visual-engineering`, Skills `[frontend-ui-ux]`.
- **QA Scenario**:
  - Tool: `bash` for unit tests; manual Electron via `npm run electron:dev` for cross-format smoke.
  - Steps: `npm test -- Sidebar`; then `npm run electron:dev` and open MD then DOCX then PDF then XLSX fixtures.
  - Expected: Sidebar shows headings, then DOCX outline, then PDF outline, then sheet list — same component, different content.
  - Pass condition: unit tests green; 4 formats render correctly through unified Sidebar.

### W3.2 — StatusBar generalize
- **Description**: Refactor `src/components/StatusBar.tsx` to consume `useViewerStats()` with exhaustive switch on `stats.kind`. Per-kind segments (e.g., spreadsheet shows `Sheet: Foo · 1234 rows × 12 cols`).
- **Files**: `src/components/StatusBar.tsx`, `src/components/__tests__/StatusBar.test.tsx`.
- **Agent**: Category `visual-engineering`, Skills `[frontend-ui-ux]`.
- **QA Scenario**:
  - Tool: `bash`.
  - Steps: `npm test -- StatusBar && npx tsc --noEmit -p tsconfig.app.json`.
  - Expected: switch over `ViewerStats` is exhaustive (tsc fails if a kind is omitted via `assertNever`); each kind renders its specific segments.
  - Pass condition: tests green, tsc green, exhaustiveness enforced.

### W3.3 — ExportMenu format-aware
- **Description**: `ExportMenu` reads `format` from current `LoadedFile`. For `pdf` → "Save a copy" copies the original `ArrayBuffer` via Electron's `dialog.showSaveDialog` + `fs.writeFile`. For other formats → "Export to PDF" via `html2canvas` + `jspdf` of active viewer DOM (uses existing pipeline). Disable irrelevant items (e.g., "Export to DOCX" only shown for markdown).
- **Files**: `src/components/ExportMenu.tsx`, `electron/main.cjs` (add `dialog:saveBinary` IPC), `electron/preload.cjs`.
- **Agent**: Category `unspecified-high`, Skills `[]`.
- **QA Scenario**:
  - Tool: manual Electron via `npm run electron:dev`; `bash` (PowerShell) for SHA hash verification.
  - Steps: 1) `npm run electron:dev`. 2) Open `sample.pdf` → Export → "Save a copy" → save to `out.pdf`. 3) Open `README.md` → Export → "Export to PDF". 4) Open `sample.xlsx` → Export → "Export to PDF". 5) Verify PDF byte-identity: `(Get-FileHash src\viewers\__fixtures__\sample.pdf -Algorithm SHA256).Hash -eq (Get-FileHash out.pdf -Algorithm SHA256).Hash`.
  - Expected: PDF copy SHA256 matches source; MD export is non-empty PDF; XLSX export is non-empty PDF (rasterized grid).
  - Pass condition: SHA equality returns `True`; 3 export paths produce valid files (size >0).

### W3.4 — Universal shortcuts
- **Description**: `src/hooks/useUniversalShortcuts.ts`: Ctrl+P → trigger ExportMenu's "Export to PDF" or "Save a copy"; Ctrl+E → open ExportMenu dropdown; PgUp/PgDn → dispatch `viewer:nav-prev` / `viewer:nav-next` custom events. Viewers (PDF, PPTX, ODP, Spreadsheet) listen via `useEffect` and advance page/slide/sheet.
- **Files**: `src/hooks/useUniversalShortcuts.ts` (new), updates to PDF/PPTX/ODP/Spreadsheet viewers to listen.
- **Agent**: Category `unspecified-high`, Skills `[]`.
- **QA Scenario**:
  - Tool: manual Electron via `npm run electron:dev`.
  - Steps: launch via `npm run electron:dev`. Open `sample.pdf` and press PgDn ×3; open `sample.pptx` and press PgDn ×3; press Ctrl+P on each; press Ctrl+E on each.
  - Expected: PDF advances 3 pages (StatusBar shows page 4/N); PPTX advances 3 slides; Ctrl+P opens save dialog; Ctrl+E opens ExportMenu dropdown.
  - Pass condition: all 4 shortcuts (PgUp/PgDn/Ctrl+P/Ctrl+E) work in PDF and PPTX.

### W3.5 — ViewerRouter wires all viewers
- **Description**: Replace registry stubs with real lazy imports for all 13 viewers. Replace existing direct markdown render in `App.tsx` with `<ViewerRouter file={loadedFile}/>`. Remove the temporary feature flag from W1.6.
- **Files**: `src/formats/registry.ts`, `src/App.tsx`.
- **Agent**: Category `unspecified-high`, Skills `[]`.
- **QA Scenario**:
  - Tool: manual Electron via `npm run electron:dev` + DevTools Console.
  - Steps: launch via `npm run electron:dev`, open one fixture per format from `src/viewers/__fixtures__/` (13 total).
  - Expected: each fixture routes to its viewer (verify via DevTools `document.querySelector('[data-viewer]').dataset.viewer`); no fallback to UnknownViewer; DevTools Console has 0 errors.
  - Pass condition: all 13 fixtures render via ViewerRouter with correct `data-viewer` attribute.

### Gate G3 (after Wave 3)
- Tool: `bash`.
- Steps: `npm test && npm run lint && npx tsc --noEmit -p tsconfig.app.json && npm run build`.
- Expected: all exit 0.
- Failure action: dispatch `quick` fix tasks per failing file.

---

## Wave 4 — Verification

### W5.1 — Static gate
- **Description**: Run `npm ci && npm run lint && npx tsc --noEmit -p tsconfig.app.json && npm test && npm run build`. Compare bundle to W0.1 baseline. Write `.sisyphus/baselines/atlas-phase1.json`.
- **Agent**: Category `quick`, Skills `[]`.
- **QA Scenario**:
  - Tool: `bash` (PowerShell).
  - Steps: `npm ci; if ($?) { npm run lint }; if ($?) { npx tsc --noEmit -p tsconfig.app.json }; if ($?) { npm test }; if ($?) { npm run build }`. Then read `.sisyphus/baselines/atlas-phase0.json` and compare main chunk size against new `dist/assets/index-*.js`.
  - Expected: all 5 commands exit 0; main chunk did not regress >25% vs W0.1 baseline (viewer libs must be lazy chunks visible in `dist/assets/` as separate files).
  - Pass condition: all green AND main chunk regression <25%; `.sisyphus/baselines/atlas-phase1.json` written.

### W5.2 — Smoke matrix (Playwright Electron)
- **Description**: `scripts/smoke.mjs` launches Electron in test mode, opens one fixture per format directly from `src/viewers/__fixtures__/` (the same fixtures created by W2.* viewer tasks — no duplication into a separate `tests/fixtures/` directory), asserts viewer mounts (DOM has `[data-viewer="<format>"]`) and StatusBar reports stats text.
- **Files**: `scripts/smoke.mjs` (new), `playwright.config.ts` (new if absent). Fixtures consumed in-place from `src/viewers/__fixtures__/*` (already produced by W2.MD, W2.DOCX, W2.XLSX, W2.PPTX, W2.PDF, W2.CSV, W2.TSV, W2.TEXT, W2.CODE, W2.ODT, W2.ODS, W2.ODP, W2.RTF — verified via `Get-ChildItem src/viewers/__fixtures__/sample.*` enumerating 13 files before `node scripts/smoke.mjs` runs).
- **Agent**: Category `deep`, Skills `[playwright]`.
- **QA Scenario**:
  - Tool: `bash` (Playwright Electron via Node).
  - Steps: `npm run build; if ($?) { node scripts/smoke.mjs }`.
  - Expected: smoke script logs `13/13 PASS`; for each format, asserts `data-viewer="<format>"` attribute exists and StatusBar text is non-empty within 5s timeout each.
  - Pass condition: exit 0 with `13/13 PASS`; PDF worker loads (no console error about workerSrc captured by Playwright).

---

## Wave 5 — Ship

### W6.1 — Build NSIS installer
- **Description**: `npm run build && npx electron-builder --win nsis --x64`. Produces `release/Atlas-Setup-2.0.0.exe`.
- **Agent**: Category `quick`, Skills `[]`.
- **QA Scenario**:
  - Tool: `bash` (PowerShell).
  - Steps: `npm run build; if ($?) { npx electron-builder --win nsis --x64 }`. Then `$f = Get-Item -LiteralPath release/Atlas-Setup-2.0.0.exe; $f.Length / 1MB`.
  - Expected: artifact `release/Atlas-Setup-2.0.0.exe` exists; reported size is between 80 and 250 MB.
  - Pass condition: `Get-Item` succeeds and size in 80–250 MB range.

### W6.2 — Manual upgrade test
- **Description**: User runs `Atlas-Setup-2.0.0.exe` over existing `MD Reader` install at `C:\Program Files\MD Reader\`. Verify: install path, Start Menu shortcut renamed to `Atlas`, file associations re-pointed for at least 5 sample extensions, user data preserved (recent files list intact, theme preference preserved via localStorage migration W4.3).
- **Agent**: Manual / user-driven.
- **QA Scenario**:
  - Tool: manual user-driven on Windows machine; PowerShell for path/registry checks.
  - Steps: 1) Pre-install: `Get-ChildItem 'C:\Program Files\MD Reader\'` and screenshot Start Menu. 2) Double-click `release/Atlas-Setup-2.0.0.exe`. 3) `Test-Path 'C:\Program Files\MD Reader\Atlas.exe'` (per-machine path preserved due to unchanged `appId`). 4) Check Start Menu shows `Atlas` shortcut. 5) Right-click a `.docx` file → Open With → confirm `Atlas` listed. 6) Launch Atlas → confirm recent files panel includes previously opened MD files. 7) Confirm theme picker shows previously selected theme.
  - Expected: all 7 checks pass; no data loss.
  - Pass condition: user confirms all 7 checks pass.

---

# Atomic Commit Strategy

One commit per task ID. Conventional Commits style. After Wave 4 (post W5.1), run `skill(name="review-work")` against the cumulative diff.

# Success Criteria (Phase 1 Done)

1. `release/Atlas-Setup-2.0.0.exe` exists and installs over existing MD Reader install with `appId: com.mdreader.app` preserved.
2. All 13 formats open from File→Open and from OS "Open with" association.
3. lint, tsc, build, test all green.
4. Main JS bundle did not regress >25%; each viewer is a separate lazy chunk.
5. Sidebar/StatusBar/ExportMenu work uniformly across formats.
6. Ctrl+P, Ctrl+E, PgUp, PgDn behave per spec in PDF + PPTX + spreadsheet.
7. localStorage migration runs once and is idempotent.
8. `grep -R "MD Reader" src/ index.html README.md` returns zero hits (outside migration shim).
