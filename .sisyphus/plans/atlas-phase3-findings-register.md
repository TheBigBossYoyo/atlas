# Atlas Findings Register

Source: whole-codebase verification audit (13 subsystem passes) + this session's synthesis, finalized after an independent critic review pass (see the "Corrections and Additions from Critic Review" section near the end, and Section 11 of the companion `atlas-phase3-improvement.md`). Every finding produced by the audit is listed below exactly once, grouped by subsystem, sorted by current severity (CRITICAL → HIGH → MEDIUM → LOW). **299 findings total** (294 from the original audit, 0 refuted, plus 5 added during critic review), **0 refuted**.

Legend — **Verify**: `CONF`=confirmed (independently reproduced/traced), `PART`=partially-confirmed (core mechanism verified, one framing detail adjusted), `UNV`=unverified (audited but not independently re-traced by the verifier pass), `RUNTIME`=observed-at-runtime (reproduced live in a running Electron instance). **Sev**: current severity; "⇩was X" / "⇧was X" marks a finding the verifier downgraded/upgraded from the drafting pass's `originalSeverity`. **Eff**: S/M/L/XL implementation effort.

Totals by current severity: **39 CRITICAL, 92 HIGH, 126 MEDIUM, 42 LOW** = 299. (Original audit: 38/92/122/42 = 294. Critic-review additions: +1 CRITICAL — ELEC-24; +4 MEDIUM — DXP-20, DXP-21, ELEC-25, ELEC-26. Two existing rows, SHELL-13 and SHELL-19, were corrected in place rather than added or removed — see the corrections section.)

---

## 1. App Shell, State Model & Chrome (`SHELL`) — Health 3/10

Four CRITICAL/HIGH defects each hit a routine workflow (drag-and-drop is dead on the shipped Electron build; unsaved edits vanish silently on close and on opening a new file; DOCX edits are never tracked as dirty at all). Two confirmed, reproducible keyboard collisions (Ctrl+B, Ctrl+E) disrupt DOCX editing on every use. Autosave writes constantly but is never read back. Zero test coverage on the entire shell layer means none of this is caught before shipping. Strengths worth preserving: the `ViewerContext`/`FormatId` discriminated-union architecture and the `localStorage` migration boundary are well-built.

| ID | Sev | Category | Verify | Location | Impact | Recommendation | Eff |
|---|---|---|---|---|---|---|---|
| SHELL-03 | CRITICAL | data-loss | CONF | `DocxViewer.tsx`; `App.tsx:82-90` | DOCX editor has no `isDirty` tracking at all — Toolbar/StatusBar dirty dot never lights for DOCX; opening another file discards all DOCX edits silently. | Extend `ViewerContext` with `setDirty(boolean)`; DocxViewer calls it whenever the model diverges from last-saved snapshot; App combines it with markdown's own `isDirty`. | M |
| SHELL-04 | CRITICAL | data-loss | CONF | `App.tsx:86-90,209-218` | Render-phase sync resets `localMarkdown`/`isDirty` purely by string-equality with no dirty check; Ctrl+O / Recent-file click / drag-drop instantly replaces unsaved markdown edits with zero confirmation. | Gate `openFile`/`openFileFromPath`/`handleOpenRecent` behind a combined dirty check; show confirm/discard dialog before proceeding. | S |
| SHELL-01 | HIGH ⇩was CRITICAL | bug | CONF | `App.tsx:135-142` | Electron 32+ removed `File.path`; `handleDrop`'s `'path' in droppedFile` guard is always false — drag-and-drop is dead for every format in the shipped app. | Add `webUtils.getPathForFile` IPC bridge in preload; call it from `handleDrop` instead of reading `.path`. | M |
| SHELL-07 | HIGH | bug | CONF | `App.tsx:69-76`; `useFileHandler.ts:22-26` | `loading`/`error` state is fully implemented and unit-tested in the hook but never destructured/rendered in `App.tsx` — file-open failures are completely silent. | Destructure `loading`/`error`; surface a dismissible toast/banner and a loading spinner. | S |
| SHELL-08 | HIGH | bug | CONF | `DocxViewer.tsx:812-900`; `useUniversalShortcuts.ts:34-73` | Ctrl+B while editing a DOCX also toggles the app sidebar every time, because DocxViewer never calls `stopPropagation()` and the global inField guard doesn't recognize `contentEditable`. | Call `stopPropagation()` in DocxViewer's handled-combo branches; widen the global `inField` check to include `isContentEditable`. | S |
| SHELL-09 | HIGH | bug | CONF | `DocxViewer.tsx:856-863` | Ctrl+E (center-align) inside the DOCX editor also pops open the Export menu on the same keypress. | Same fix as SHELL-08 — stop propagation once DocxViewer recognizes the combo. | S |
| SHELL-11 | HIGH | feature-gap | CONF | `useAutosave.ts:40-53` | `loadDraft()`/`clearDraft()` exist for crash recovery but have zero call sites anywhere — autosave writes every 800ms but is never read back, so it protects nothing despite being advertised. | Call `loadDraft()` on mount; prompt to restore if present/newer; `clearDraft()` after a successful save. | M |
| SHELL-13 | HIGH | fidelity | CONF | `App.tsx:236-238`; `SpreadsheetViewer.tsx` | Generic PDF export is a single `html2canvas-pro` screenshot of `#viewer-content`; for virtualized spreadsheet grids this exports only the currently-scrolled-into-view rows/columns with no warning. *(Corrected during critic review: the original text named the package `html2canvas` — the runtime import at `src/utils/export.ts:16` is in fact `html2canvas-pro`, a different npm package; the bare `html2canvas` package is unused dead weight. See corrections section.)* | Make export format-aware: build spreadsheet PDFs from the parsed data model, or disable/relabel PDF export for these formats. | L |
| SHELL-02 | MEDIUM ⇩was CRITICAL | data-loss | PART | `App.tsx:277-285` | The only close-guard is a renderer `beforeunload` handler; Electron/Chromium does not surface a visible confirmation dialog for it, so closing with unsaved changes is silently allowed in the packaged app. | Move the guard to the main-process `close` handler via a renderer→main dirty-state IPC signal + `dialog.showMessageBoxSync`. | M |
| SHELL-05 | MEDIUM | bug | UNV | `App.tsx:86-100` | The dirty-reset guard keys off markdown-string equality, not file identity; switching from a dirty sample/markdown session to a binary file leaves the stale dirty dot lit next to the new file's name. | Key the reset off `file.path` identity (matching `ViewerContext`'s pattern) instead of derived string equality. | S |
| SHELL-06 | MEDIUM | bug | UNV | `App.tsx:66`; `useFileHandler.ts:57`; `useRecentFiles.ts:36-46` | App and `useFileHandler` each hold an independent `useRecentFiles()` instance; removing a recent entry in one is silently undone the next time the other's stale `addRecent` fires. | Lift `useRecentFiles()` to one call in `App.tsx`; pass `addRecent` into `useFileHandler` as a parameter. | S |
| SHELL-10 | MEDIUM | bug | UNV | `useUniversalShortcuts.ts:38-47`; `App.tsx:102-112`; `DocxViewer.tsx:836-840` | Ctrl+S while editing DOCX silently no-ops: the global handler's `saveFile()` returns false for non-markdown, and DocxViewer's own handler explicitly swallows the key expecting some other path to save (none does). | Thread the active viewer's own save handler through the shared capability contract so global Ctrl+S dispatches to it when non-markdown. | M |
| SHELL-12 | MEDIUM | bug | UNV | `App.tsx:152`; `useAutosave.ts:21-37` | `useAutosave` runs unconditionally; after SHELL-05's bug, a persisted draft can pair stale markdown text with a binary file's name — low-impact only until SHELL-11 is fixed, then actively harmful. | Gate `useAutosave` on `isMarkdownDocument`, fixed together with SHELL-05. | S |
| SHELL-14 | MEDIUM | ux | UNV | `Toolbar.tsx:136-161`; `useFontSize.ts:27` | Font-size +/-/reset buttons are always visible/enabled but only affect `.markdown-body`; clicking them on any other format does nothing with no explanation. | Hide/disable the group when `!isMarkdown`, mirroring the existing `canSearch` pattern. | S |
| SHELL-15 | MEDIUM | ux | UNV | `Toolbar.tsx:98-106`; `DocxViewer.tsx:1284-1287` | A permanently-disabled global Save button (misleading enabled-looking tooltip) sits next to DocxViewer's own working Save button — two visually different Save controls on screen at once. | Hide the global Toolbar Save button entirely when `!isMarkdown`. | S |
| SHELL-16 | MEDIUM | feature-gap | UNV | `useFileHandler.ts:143-147` | `clear()` fully resets file/error/loading but is never wired to any button/menu/shortcut — no way to close the current document and return to Welcome. | Add a "Close file" action/shortcut (e.g. Ctrl+W) calling `clear()`, gated behind the same unsaved-changes confirmation. | S |
| SHELL-17 | MEDIUM | feature-gap | UNV | `App.tsx:63-100` | No multi-document/tabs support — opening file B always fully replaces file A's session, compounding SHELL-04's data-loss risk. | Architecture-level gap; flag for a dedicated future roadmap item (document-session array + tab UI), not a quick patch. | XL |
| SHELL-18 | MEDIUM | ux | UNV | `index.html:7`; `App.tsx` | Window/title bar is hardcoded `"Atlas"` — never reflects open filename or dirty state, unlike every comparable desktop editor. | Set `document.title` reactively to `${dirty?'● ':''}${fileName} — Atlas`, or route through `mainWindow.setTitle`. | S |
| SHELL-19 | MEDIUM | architecture | PART | `useUniversalShortcuts.ts:98`; `useFontSize.ts:60`; `useSearch.ts:123,140`; `ThemeMenu.tsx:28`; `ExportMenu.tsx:36`; `ShortcutsModal.tsx:17` | Six independent `window.addEventListener('keydown')` sites with no shared registry or precedence — the direct architectural cause of SHELL-08/09/UX-03 and a latent risk for every future viewer/shortcut. *(Corrected during critic review: the original text named only four of the six sites — `useUniversalShortcuts`, `useFontSize`, `useSearch`, `ThemeMenu` — and miscounted DocxViewer's handler as a "window-level" fifth to reach "six." The two omitted sites are `ExportMenu.tsx:36` and `ShortcutsModal.tsx:17` [both Escape-only]. DocxViewer's own combo-key handling at `DocxViewer.tsx:1321` is confirmed to be an element-level `onKeyDown` prop on its contentEditable div, not a `window` listener — it is a separate, 7th contributor to the collision problem via its failure to stop propagation, not one of the six being replaced. See corrections section.)* | Introduce one shared shortcut-dispatcher (focused input/contentEditable → active viewer → shell-global precedence), replacing all six `window`-level listeners and routing DocxViewer's element-level handler into the same precedence chain. | L |
| SHELL-20 | MEDIUM | bug | UNV | `useFileHandler.ts:73,82,91,134` | `window.electronAPI!` non-null assertions crash file-open with an unhandled rejection when the app runs in a plain browser tab (`npm run dev`, a documented workflow). | Guard each call; add a browser-mode fallback (`<input type=file>` + `arrayBuffer()`) so plain-browser dev stays usable. | M |
| SHELL-21 | LOW | tech-debt | UNV | `App.tsx:29,370` | `ENABLE_VIEWER_ROUTER = true` is a permanently-true, unwireable dead feature flag. | Delete the flag; collapse the ternary to the only reachable branch. | S |
| SHELL-22 | LOW | tech-debt | UNV | `App.tsx:386-387` | A hidden `<span>` "suppress unused-var" hack for `filePath` is itself dead code (filePath is already consumed twice above it). | Delete the hidden span and stale comment. | S |
| SHELL-23 | LOW | tech-debt | UNV | `useTheme.ts:20` | `theme as never` cast when calling the Electron bridge defeats type-checking on that call for no reason (types already match). | Remove the cast; pass `theme` directly. | S |
| SHELL-24 | LOW | bug | UNV | `App.tsx:128-131` | Drag-overlay `dragleave` uses the classic target-identity false-negative check; once SHELL-01 is fixed, the "Drop your file" overlay can flicker/stick. | Replace with an enter/leave counter pattern; fix alongside SHELL-01. | S |
| SHELL-25 | LOW | ux | UNV | `useTheme.ts:8-12` | System dark/light preference is read once at mount only — no `matchMedia` change listener, so a live OS theme toggle has no effect until restart. | Add a `matchMedia('change')` listener that reapplies system preference only when the user has no explicit saved theme. | S |
| SHELL-26 | LOW | docs | UNV | `README.md:61`; `useUniversalShortcuts.ts:79-88` | README's Ctrl+2/Ctrl+3 shortcut table disagrees with actual behavior and the in-app ShortcutsModal (Split vs Editor swapped). | Fix README row to read Ctrl+1 Preview / Ctrl+2 Split / Ctrl+3 Editor. | S |
| SHELL-27 | LOW | testing | UNV | `App.tsx`; `src/components/__tests__`; `src/hooks/__tests__` | Zero dedicated tests for `App.tsx`, Toolbar, WelcomeScreen, ThemeMenu, ShortcutsModal, DropZone, SearchOverlay, ExportMenu, and most hooks — exactly the class of regression SHELL-04–11 represent. | Prioritize `App.tsx` integration tests using the existing `ViewerProvider`-harness pattern. | M |

## 2. Electron Main/Preload, IPC, Security & Packaging (`ELEC`) — Health 3/10

Drag-and-drop is fully non-functional on the pinned Electron version; the file IPC surface allows unrestricted absolute-path reads and silent overwrites with no sender/path validation; there is no CSP and the sandbox is explicitly disabled; a confirmed file-association regression (`.mdown`) and no real close-confirmation round out the CRITICAL/HIGH picture. A critic-review pass additionally confirmed that the `save-file` handler used by markdown itself shares the identical non-atomic-write risk previously flagged only for `save-binary-file` (ELEC-24, new), and that no crash/error logging exists anywhere in the main process. Basic wiring (`contextIsolation`, single-instance lock, correct relative asset base) is sound.

| ID | Sev | Category | Verify | Location | Impact | Recommendation | Eff |
|---|---|---|---|---|---|---|---|
| ELEC-02 | CRITICAL | security | CONF | `main.cjs:233-240,258-262` | `file:readBinaryByPath`/`open-file-by-path` allow reading any file on disk with no sender/path validation and no size cap. | Restrict to paths the main process itself vouches for; add sender-frame check; cap read size. | M |
| ELEC-03 | CRITICAL | security | CONF | `main.cjs:264-289,291-316` | `save-file`/`save-binary-file` silently overwrite any renderer-supplied `existingPath` with no ownership check and no confirmation. | Track "path this window may silently save to" as main-process state; reject other paths, fall back to save dialog. | M |
| ELEC-24 | CRITICAL | data-loss | CONF | `main.cjs:264-289` (`save-file` handler) | **Added during critic review.** The `save-file` IPC handler — the one markdown (the app's own "perfect" feature) and every other text-class format saves through — calls `fs.writeFileSync(targetPath, req.content, 'utf-8')` directly with no temp-file staging, no `fsync`, and no backup, confirmed by direct code read at `main.cjs:283`. This is the exact same crash-mid-write / no-recovery risk already flagged as CRITICAL for the adjacent `save-binary-file` handler (see DXS-12/RUN-10), but was never flagged for this handler in the original audit. | Apply the identical temp-file-then-`fsync`-then-atomic-rename-plus-rolling-`.bak` pattern used to fix `save-binary-file`; extract a shared `atomicWriteFile()` helper so both handlers use the same code path. | S |
| ELEC-01 | HIGH ⇩was CRITICAL | bug | CONF | `App.tsx:135-141`; `preload.cjs` | Same root cause as SHELL-01 — `webUtils` is never imported/exposed anywhere in the codebase. | Expose `getPathForFile` via `contextBridge`. | S |
| ELEC-04 | HIGH | security | CONF | `index.html:1-13`; `main.cjs:106-112` | No Content-Security-Policy anywhere; `sandbox:false` with no evident need. | Add a strict CSP via `session.webRequest.onHeadersReceived`; set `sandbox:true` and verify. | M |
| ELEC-05 | HIGH | bug | CONF | `main.cjs:33-43`; `electron-builder.yml:30-33`; `installer.nsh:44,138-139` | `.mdown` is registered as a Windows file association but omitted from `KNOWN_EXTENSIONS` — double-clicking a `.mdown` file silently fails. | Add `mdown`; generate all extension lists from one shared source. | S |
| ELEC-06 | HIGH | data-loss | CONF | `main.cjs` (no close/quit handler); `App.tsx:277-285` | No main-process `close`/`before-quit` handler at all. | Add `mainWindow.on('close', ...)` with `dialog.showMessageBoxSync`. | M |
| ELEC-07 | MEDIUM | bug | UNV | `main.cjs:422-430` | Second-instance file-open requests dropped if they arrive before `mainWindow` exists. | Set `pendingFilePath` when `mainWindow` is null. | S |
| ELEC-08 | MEDIUM | performance | UNV | `main.cjs:227-230,233-240` | No size limit or streaming on any file-read IPC handler. | Add an `fs.promises.stat` size check before reading. | M |
| ELEC-09 | MEDIUM | tooling | UNV | `electron-builder.yml:1-20` | No code-signing configuration — installers are unsigned. | Obtain a code-signing certificate; wire via CI. | L |
| ELEC-10 | MEDIUM | bug | UNV | `electron-builder.yml:20`; `scripts/build-icon.mjs` | `signAndEditExecutable:false` likely also skips icon/version metadata embedding. | Build once and inspect; decouple if confirmed. | S |
| ELEC-11 | MEDIUM | feature-gap | UNV | `main.cjs:93-113` | No window state (size/position/maximized) persistence between launches. | Persist bounds to a JSON under `userData`; restore on create. | S |
| ELEC-12 | MEDIUM | security | CONF | `main.cjs` (no `Menu.setApplicationMenu`); `main.cjs:184-186` | No custom application menu — default menu exposes Reload/DevTools in production. | Build a minimal production `Menu` template omitting Reload/DevTools. | S |
| ELEC-13 | MEDIUM | crash | UNV | `main.cjs` (no `render-process-gone`/`unresponsive` handlers) | No renderer-crash handling and no persistent logging. Confirmed during critic review via a repo-wide grep: zero `crashReporter`/`uncaughtException`/`unhandledRejection` handlers exist anywhere in `electron/main.cjs`. | Add handlers with a recovery screen; route errors to a log file. **Sequencing note (critic review):** moved forward from Phase 5 to Phase 1 (Task P1.15) in the improvement plan, since Phases 1–4's shared-shell surgery is exactly the period most in need of crash diagnostics. | M |
| ELEC-14 | MEDIUM | feature-gap | UNV | `package.json`; `electron-builder.yml` | No auto-update mechanism. | Add `electron-updater` + a publish target once a release channel exists. | L |
| ELEC-15 | MEDIUM | architecture | UNV | `main.cjs:33-43,207-217`; `electron-builder.yml:21-193`; `detect.ts:3-60` | Known-extension lists hand-duplicated across 3-4 disagreeing sources. | Define one canonical table; generate the others from it at build time. | M |
| ELEC-16 | MEDIUM | tech-debt | UNV | `preload.cjs:8`; `electron.d.ts:29-58`; `DocxViewer.tsx:81-83,1164-1168` | `saveBinaryFile` missing from the shared `ElectronAPI` type, forcing a local re-declaration + cast. | Add to the single `ElectronAPI` interface; delete the shim. | S |
| ELEC-18 | MEDIUM | bug | UNV | `preload.cjs:17-21` | `onFileOpenedPath`'s unsubscribe calls `removeAllListeners` instead of `removeListener`. | Change to `removeListener('file-opened-path', handler)`. | S |
| ELEC-19 | MEDIUM | bug | UNV | `useFileHandler.ts:73,82,91,134` | Same class as SHELL-20 — non-null assertions crash the plain-browser dev workflow. | Guard each call site; add a browser-mode fallback. | S |
| ELEC-25 | MEDIUM | feature-gap | UNV | `main.cjs` (path-allowlist design, P1.2); `main.cjs:291-316` (atomic-write design, P1.8) | **Added during critic review.** No app-manifest `longPathAware` opt-in was found under `build/`, and no evidence the planned IPC path-allowlist or atomic-write temp-then-rename pattern accounts for Windows' legacy 260-character `MAX_PATH` limit or UNC network-share paths (`\\server\share\file.docx`); a cross-volume atomic rename can fail with `EXDEV` on some network filesystems. | Normalize allowlist path comparisons to handle `\\?\`-prefixed long paths and UNC paths consistently; fall back to a direct (still backed-up) write if an atomic rename fails with `EXDEV`. | S |
| ELEC-26 | MEDIUM | data-loss | UNV | `main.cjs:291-316` (`save-binary-file` handler) | **Added during critic review.** DOCX is the app's flagship problem format and Microsoft Word is near-certain to be installed on the same Windows machines; no code anywhere handles the case where the same `.docx` is concurrently open in Word (Word's exclusive/shared lock, its `~$file.docx` sibling), nor how an atomic rename-over-target interacts with a file Word may re-save moments later. | Best-effort mitigation: catch `EBUSY`/`EPERM` from the write/rename and surface a specific "this file appears to be open in another program — close it there first" message, rather than a generic failure. Proactive lock detection (checking for Word's lock-file sibling before allowing an edit session) is not attempted — unreliable and out of scope. | S |
| ELEC-20 | LOW | docs | UNV | `electron-builder.yml:21-193` | All ~30 file associations use `role: Editor`, including view-only formats. | Set `role: Viewer` for view-only formats. | S |
| ELEC-21 | LOW | feature-gap | UNV | `main.cjs:291-316`; `App.tsx:102-123` | The generic `save-binary-file` handler is only wired up by DocxViewer. | Route future binary-format saves through a format-agnostic callback. | M |
| ELEC-22 | LOW | tech-debt | UNV | `electron-builder.yml:1` | Windows `appId` is still the legacy `com.mdreader.app`. | Document the install-continuity reason, or plan a deliberate migration. | S |
| ELEC-23 | LOW | tooling | UNV | `scripts/build-icon.mjs`; `package.json:9-19` | `build-icon.mjs` is not wired into any npm script. | Add an `"icon"` script; run it pre-`electron:build`. | S |
| ELEC-17 | LOW | tech-debt | UNV | `preload.cjs:22-42` | Spellcheck IPC exposed twice (flat + namespaced) over identical channels. | Pick one shape; remove the flat duplicates. | S |

## 3. File Loading, Format Detection & Viewer Routing (`LOAD`) — Health 4/10

The type system and `detect.ts` logic are well-designed and well-tested in isolation, and markdown never enters this subsystem at all. But two CRITICAL, concretely-confirmed defects sit alongside a confirmed Windows file-association regression and eight HIGH-severity issues (silent failures, races, a viewer crash that bricks the router for every subsequently-opened good file).

| ID | Sev | Category | Verify | Location | Impact | Recommendation | Eff |
|---|---|---|---|---|---|---|---|
| LOAD-01 | CRITICAL | data-loss | CONF | `App.tsx:81-90,145-150`; `useAutosave.ts:21-28` | Opening any new file silently discards unsaved markdown edits *and* destroys the autosave backup, with zero confirmation. | Single dirty-check guard at every open-trigger entry point. | M |
| LOAD-02 | CRITICAL | bug | CONF | `App.tsx:135-142`; `preload.cjs:1-43` | Same root defect as SHELL-01/ELEC-01 — drag-and-drop dead for all 14 formats. | `webUtils.getPathForFile` preload bridge. | S |
| LOAD-03 | HIGH ⇩was CRITICAL | bug | PART | `electron-builder.yml:30-33,106-109`; `main.cjs:33-61`; `detect.ts:3-85` | `.mdown`/`.ini` are registered Windows file associations that cannot actually be opened by double-click. | One extension→format manifest feeding all consumers. | M |
| LOAD-05 | HIGH | fidelity | CONF | `main.cjs:63-71`; `useFileHandler.ts:96` | All text-class files blind-decoded as UTF-8 with no BOM sniffing — non-UTF-8 files render as mojibake. | Sniff BOM; fall back to Windows-1252 on invalid UTF-8. | M |
| LOAD-06 | HIGH | ux | CONF | `App.tsx:69-76`; `useFileHandler.ts:50-124` | Same defect as SHELL-07 — `error`/`loading` never rendered. | Destructure and surface both. | S |
| LOAD-07 | HIGH | bug | CONF | `useFileHandler.ts:63-124`; `main.cjs:422-430` | No request token on concurrent loads — a slower, earlier request can overwrite a faster, newer one. | Add a monotonically increasing request id; discard stale results. | M |
| LOAD-08 | HIGH | bug | CONF | `ViewerErrorBoundary.tsx:12-49`; `ViewerRouter.tsx:22-29` | `ViewerErrorBoundary` never resets on file change — one crashed file bricks the crash screen for every subsequent valid file. | `<ViewerErrorBoundary key={file.path}>`. | S |
| LOAD-10 | HIGH | feature-gap | CONF | `detect.ts:87-96,109-132`; `UnknownViewer.tsx:9-16` | No text-sniffing fallback for extensionless files (README, LICENSE, Dockerfile). | Apply a cheap text-sniff heuristic in the unknown-extension branch. | M |
| LOAD-11 | HIGH | feature-gap | CONF | `detect.ts:3-85,109-132`; `UnknownViewer.tsx:9-16` | Legacy `.doc`/`.xls`/`.ppt` (OLE/CFB magic) have zero detection support. | Add a CFB magic check; surface a specific, honest message. | M |
| LOAD-04 | MEDIUM ⇩was HIGH | bug | PART | `useFileHandler.ts:68-88`; `detect.ts:148-170` | Recognized-extension fast path never verifies magic bytes. | Always run `detectFormat` even on the fast path. | M |
| LOAD-09 | MEDIUM ⇩was HIGH | bug | PART | `ViewerRouter.tsx:12-20`; `ViewerErrorBoundary.tsx:31-34` | A rejected dynamic import is permanently cached by `React.lazy` — "Try again" can never recover. | Replace with a retry-capable loader keyed on a bumpable counter. | M |
| LOAD-12 | MEDIUM | bug | UNV | `detect.ts:8-13` | `EXTENSION_MAP` omits `.docm` while siblings are present. | Add `docm: 'docx'` (and template siblings). | S |
| LOAD-13 | MEDIUM | performance | UNV | `useFileHandler.ts:130-137`; `main.cjs:200-231` | Every dialog-based open reads the entire file from disk twice. | Pass the already-fetched buffer through instead of re-reading. | S |
| LOAD-14 | MEDIUM | performance | UNV | `main.cjs:200-240` | No file-size guard anywhere in the read pipeline. | Add an `fs.promises.stat` size check before reading. | M |
| LOAD-15 | MEDIUM | performance | UNV | `detect.ts:98-107,120-129` | ZIP magic-detection decodes the entire file byte-by-byte just to substring-search. | Limit the scan to the first few KB. | S |
| LOAD-16 | MEDIUM | bug | UNV | `App.tsx:66`; `useFileHandler.ts:57`; `useRecentFiles.ts:33-64` | Same duplicate-hook defect as SHELL-06. | Lift `useRecentFiles()` to one call; pass `addRecent` down. | S |
| LOAD-17 | MEDIUM | bug | UNV | `useFileHandler.ts:73,82,91,134` | Same non-null-assertion crash class as SHELL-20/ELEC-19. | Guard with the same pattern used in the mount effect. | S |
| LOAD-18 | MEDIUM | ux | UNV | `UnknownViewer.tsx:9-16` | `UnknownViewer` renders exactly `<span>{format}</span>` + path — zero actionable UX. | Build a real empty-state UI with actionable options. | M |
| LOAD-19 | MEDIUM | architecture | UNV | `ViewerRouter.tsx:12-20` | `lazy()` memo keyed only on `file.path`, not `format`. | Add `file.format` to the `useMemo` dependency array. | S |
| LOAD-20 | LOW | feature-gap | UNV | `useAutosave.ts:40-53` | Same dead-code defect as SHELL-11. | Wire `loadDraft()`/`clearDraft()` into a restore-prompt flow. | S |
| LOAD-21 | LOW | performance | UNV | `ViewerRouter.tsx:12-20`; `ViewerRouter.test.tsx:92-116` | Every file switch pays a full unmount/remount even between two same-format files (deliberate design). | Cache the `lazy()` wrapper per `FormatId` if worth optimizing. | S |
| LOAD-22 | LOW | ux | UNV | `DropZone.tsx:11-17`; `detect.ts:3-85` | DropZone copy is stale ("Drop your Markdown file") despite 14-format support. | Update copy; add explicit template-extension entries + tests. | S |

## 4. DOCX Parser, Document Model & Fonts (`DXP`) — Health 4/10

Direct-formatting parsing is genuinely thorough, and the hand-rolled TTF/canvas font-metrics engine reflects serious iterative engineering. But named styles are parsed and then never consulted by the render pipeline; footnotes/endnotes/headers/footers are parsed into unusable opaque blobs; list numbering produces no markers; `w:tblGrid` is parsed and discarded. A critic-review pass additionally confirmed two gaps: no defense-in-depth guard on the XML parser beyond zip-level size checks (DXP-20), and no `devicePixelRatio` awareness in the canvas-based glyph-measurement engine, unlike the equivalent PDF-viewer fix (DXP-21).

| ID | Sev | Category | Verify | Location | Impact | Recommendation | Eff |
|---|---|---|---|---|---|---|---|
| DXP-01 | CRITICAL | fidelity | CONF | `paginate.ts:683-984` | Named styles parsed correctly but the renderer's local merge functions never consult `document.styles`. | Replace the local merge functions with the correct cascade from `cascade.ts`. | L |
| DXP-03 | CRITICAL | feature-gap | CONF | `footnotes.ts:54-59,97-98`; `endnotes.ts:88-89` | Footnotes/endnotes are the "Option B" stub — one opaque blob, no marker, no text. | Port `comments.ts`'s synthetic-wrapper technique. | M |
| DXP-04 | CRITICAL | feature-gap | CONF | `headers.ts:34-48`; `footers.ts:34-46` | Headers/footers store the entire raw XML as one blob — always blank in the viewer. | Same fix pattern as DXP-03. | M |
| DXP-07 | CRITICAL | feature-gap | CONF | `numbering.ts:1-271`; `paginate.ts:944-948` | Numbering is parsed thoroughly but no marker is ever generated in the live pipeline. | Implement counter state per numId/ilvl in the layout pipeline. | L |
| DXP-09 | CRITICAL ⇧was HIGH | fidelity | PART | `document.ts:122-130,427-458`; `run.tsx:47-78` | Anchored/floating drawings lose all position/wrap metadata, render identically to inline. | Parse `wp:positionH/V` and wrap mode; give an absolutely-positioned render path. | M |
| DXP-06 | HIGH | feature-gap | CONF | `styles.ts:393-402`; `parser/styles.ts:303-318` | `w:tblStylePr` (table conditional formatting) is not modeled or parsed at all. | Add conditional-formats map; parse each `w:tblStylePr` by type. | L |
| DXP-08 | HIGH | feature-gap | CONF | `document.ts:150-153,217-246,257-288` | `mc:AlternateContent`, `w:sdt`, VML, `w:sym` entirely unhandled. | Unwrap `w:sdt` transparently; add `w:sym`; prefer `mc:Fallback`. | M |
| DXP-12 | HIGH | fidelity | CONF | `families.ts:40-91`; `fontResolution.ts:23` | No substitute for "Aptos" (Word's default font since late 2023). | Add an Aptos entry with a bundled substitute; log unmapped fonts. | M |
| DXP-15 | HIGH | security | CONF | `unzip.ts:44-78` | Zip extraction has no size/entry limits — a zip-bomb can hang/crash the renderer. | Check uncompressed size per-entry/running-total before `.async()`. | S |
| DXP-19 | HIGH | bug | CONF | `document.ts:530-558`; `styles.ts:349-353`; `layoutTable.ts:37-39,161-178` | `w:tblGrid` explicitly skipped by the parser; layout uses a type-only cast that's always undefined. | Parse `w:gridCol` into a `Table.grid` field. | M |
| DXP-02 | MEDIUM ⇩was HIGH | fidelity | PART | `paragraph.tsx:32,40,46`; `run.tsx:10-13` | Run character styles resolve against the paragraph's style id instead of the run's own. | Resolve `rStyle` against the run's own id. | S |
| DXP-05 | MEDIUM | tech-debt | UNV | `comments.ts:83-89`; `document.ts:367-375` | `extractCommentText`'s `Array.isArray` truthiness bug (same root cause as QA-02/DXE-10). | Delete the unreachable legacy-blocks fallback path. | S |
| DXP-10 | MEDIUM ⇩was HIGH | testing | CONF | `families.ts:7-27`; `vite.config.ts`; `vitest.config.ts` | Root-absolute font `?url` imports break Vitest — 7 DOCX suites fail to load. | Move TTFs to relative imports under `src/docx/fonts/assets/`. | S |
| DXP-14 | MEDIUM | feature-gap | UNV | `styles.ts:330-352`; `document.ts:670-721,902` | Paragraph-level `w:bidi` is never parsed. | Parse `<w:bidi>` into `ParaProps.bidi`. | S |
| DXP-16 | MEDIUM | bug | UNV | `document.ts:108-126`; `styles.ts:85-111`; `numbering.ts:43-63` | These three modules never wrap malformed-XML failures in `DocxParseError`. | Wrap top-level `xmlParser.parse()` in try/catch. | S |
| DXP-20 | MEDIUM | security | PART | `src/docx/parser/document.ts`, `styles.ts`, `numbering.ts`, `headers.ts`, `footers.ts`, `footnotes.ts`, `endnotes.ts`, `comments.ts`, `relationships.ts`, `contentTypes.ts`, `theme.ts` (all `fast-xml-parser` call sites) | **Added during critic review.** D21's zip-level size guard (DXP-15) catches a small-compressed/huge-uncompressed zip bomb, but not a modest-uncompressed-size XML part that is pathologically deep or wide once it reaches `fast-xml-parser` (confirmed via grep: `fast-xml-parser@^5.8.0` is used at all 11 DOCX parser modules listed). **Scope note:** confirmed via code read that `PptxViewer.tsx`/`OdpViewer.tsx` use the browser's native, sandboxed `DOMParser` instead (`new DOMParser().parseFromString(...)`), not `fast-xml-parser` — this finding is DOCX-only, correcting a broader "DOCX/PPTX/ODP" framing raised during review. No specific known CVE was confirmed for the pinned version; framed as defense-in-depth, not a proven exploit. | Add a depth/node-count cap or a conservative per-part size ceiling (e.g. 20MB) before parsing, alongside D21's existing zip-bomb guard. | S |
| DXP-21 | MEDIUM | fidelity | UNV | `fonts/canvasMetrics.ts` | **Added during critic review.** DOCX's hand-rolled canvas-based glyph-measurement engine has no `devicePixelRatio` handling; PDF-18 already flags and plans the identical fix for the PDF viewer, but no equivalent task existed for DOCX despite it being the app's own centerpiece fidelity work and Windows laptops commonly running 125%/150%/200% display scaling. | Make `canvasMetrics.ts`'s glyph measurement `devicePixelRatio`-aware, mirroring PDF-18's approach. | M |
| DXP-11 | LOW | tech-debt | UNV | `fontFaces.css:1-281`; `index.css:5-10` | Orphaned `fontFaces.css` contains the exact hardcoded-path bug already fixed once. | Delete the orphaned file. | S |
| DXP-13 | LOW | feature-gap | UNV | `families.ts`; `docx/index.ts` | No support for documents with embedded fonts. | Track as a documented, deferred gap. | XL |
| DXP-17 | LOW | fidelity | UNV | `canvasMetrics.ts:256-264`; `render/style.ts:212` | Small-caps measurement drifts from actual reduced-size paint width. | Measure at reduced size, or document the drift. | S |
| DXP-18 | LOW | performance | UNV | `DocxViewer.tsx:128-169,1223`; `families.ts:40-91` | All 5 bundled font families (~7MB) load eagerly regardless of which fonts are used. | Load only referenced fonts; backfill the rest lazily. | M |

## 5. DOCX Layout, Pagination & Rendering (`DXL`) — Health 3/10

A well-engineered data-flow architecture renders wrong output for nearly every real document: alignment/indentation have zero effect on line position, space-before/after is never applied, named styles never cascade into the live view, list markers never render, images render as a placeholder glyph, and justification is dead code. A second, mostly-correct-but-unused "legacy flow" renderer sits dormant in the same codebase.

| ID | Sev | Category | Verify | Location | Impact | Recommendation | Eff |
|---|---|---|---|---|---|---|---|
| DXL-01 | CRITICAL | fidelity | CONF | `paginate.ts:712-984`; `cascade.ts:26-67` | Same root cause as DXP-01 in the layout layer. | Same fix as DXP-01. | L |
| DXL-02 | CRITICAL | fidelity | CONF | `paginate.ts:575-600`; `breakLines.ts:319-329`; `PageView.tsx:365-378` | Alignment/indentation have zero effect — every paragraph renders flush-left. | Compute `leftPt` including indent/alignment offset in `placeLine`. | M |
| DXL-03 | CRITICAL | fidelity | CONF | `itemize.ts:73-84`; `types.ts:8-17`; `PageView.tsx:268-309` | Embedded images render as an invisible placeholder glyph, not a picture. | Add a real `drawing` LineItem variant; render via media resolver. | L |
| DXL-04 | HIGH ⇩was CRITICAL | feature-gap | CONF | `numbering.ts`; `paginate.ts:944-951`; `PageView.tsx` | List markers never rendered in the shipped path. | Resolve marker text via `document.numbering` in layout. | M |
| DXL-05 | HIGH | fidelity | CONF | `paginate.ts:592-599`; `breakLines.ts:403-419` | Paragraph space-before/after is never applied. | Add spacing between paragraphs, respecting `contextualSpacing`. | M |
| DXL-06 | HIGH | fidelity | CONF | `itemize.ts:63-70`; `breakLines.ts:56-70`; `types.ts:31-35` | Manual page/column breaks (Ctrl+Enter) render as ordinary line breaks. | Add break-kind tracking to `LineBox`; split placement chunks. | M |
| DXL-07 | HIGH | fidelity | CONF | `paginate.ts:172-249,424-452`; `document.ts:261-266` | 'continuous'/'nextColumn' section breaks always force a full page break. | Branch on section type; keep flowing or advance column instead. | M |
| DXL-08 | HIGH | fidelity | CONF | `layoutTable.ts:47-85,493-507`; `document.ts:300-310` | Table style cascade entirely unimplemented — banding/header shading never apply. | Resolve `tblStyle`'s basedOn chain + conditional blocks. | L |
| DXL-09 | HIGH | feature-gap | CONF | `pageTypes.ts:113-130`; `paginate.ts` | Footnotes/endnotes never laid out or rendered anywhere. | Extend layout with a footnote content map + reserved space. | XL |
| DXL-10 | HIGH | feature-gap | CONF | `styles.ts:330-352`; `breakLines.ts:448`; `PageView.tsx` | No RTL/bidi support anywhere in layout/render. | Add paragraph bidi, default right alignment, `dir=rtl`. | L |
| DXL-11 | HIGH | bug | CONF | `breakLines.ts:139-155`; `PageView.tsx:245-334` | Computed justification stretch is dead code — "Justify" renders ragged. | Render stretchable spaces at `item.width + justificationStretch`. | S |
| DXL-12 | HIGH | feature-gap | CONF | `paginate.ts:856-878`; `types.ts:8-51`; `PageView.tsx:268-309` | Hyperlinks not clickable or visually distinct. | Track hyperlink target per run; wrap spans in `<a href>`. | M |
| DXL-13 | HIGH | performance | CONF | `DocxViewer.tsx:1211-1266`; `paginate.ts:158-251` | Full-document re-pagination on every edit, no incremental relayout. | Cache LineBoxes by hash; re-run only changed units. | L |
| DXL-14 | MEDIUM | performance | UNV | `PageStack.tsx:15-28`; `page-view.css:16-24` | `PageStack` mounts every page unconditionally, no virtualization. | Virtualize with windowed range + overscan. | M |
| DXL-15 | MEDIUM | feature-gap | UNV | `paginate.ts:786-789,509-524`; `page-view.css:14` | Table rows can never split across a page boundary. | Extend cells with per-line break points for unset-cantSplit rows. | L |
| DXL-16 | MEDIUM | feature-gap | UNV | `breakLines.ts:282-292`; `paginate.ts:880-910` | Tab stop alignment (center/right/decimal) and leaders never implemented. | Compute pending-content width; render leader glyphs. | M |
| DXL-17 | MEDIUM | fidelity | UNV | `document.ts:275,287`; `paginate.ts:97-107,986-1012` | Section `w:vAlign` parsed but never applied. | Offset `contentTopPt` by slack height for center/bottom. | S |
| DXL-18 | MEDIUM | feature-gap | UNV | `itemize.ts:114-145` | No automatic-hyphenation support. | Parse `w:autoHyphenation`; run a pattern-based hyphenator. | L |
| DXL-19 | MEDIUM | feature-gap | UNV | `DocxViewer.tsx:1332`; `PageView.tsx:225-243` | Zoom hardcoded to 1 despite `PageView` fully supporting it. | Add a zoom state + toolbar controls. | S |
| DXL-20 | MEDIUM | tech-debt | UNV | `Renderer.tsx`, `paragraph.tsx`, `run.tsx`, `table.tsx` | A complete, correct-cascade "legacy flow" renderer ships unused — a maintenance hazard. | Repurpose its logic as a porting source, then delete it. | S |
| DXL-21 | LOW | ux | UNV | `page-view.css:1-8,10-12`; `index.css` | DOCX backdrop uses undefined CSS vars, always falling back to light-mode literals. | Bind to the app's real theme tokens. | S |

## 6. DOCX Editor & DocxViewer Integration (`DXE`) — Health 2/10

Basic single-paragraph, single-run editing works, but inserting an image after any other edit silently destroys those edits; multi-command operations can corrupt the undo stack; any paragraph containing a hyperlink is effectively uneditable; cross-paragraph and cross-run selections cannot be deleted/formatted; lists/indent/tables are wired-but-unimplemented stubs; Ctrl+S does not save; accepted spellcheck fixes are lost on save. 0% of the dedicated test suites currently execute.

| ID | Sev | Category | Verify | Location | Impact | Recommendation | Eff |
|---|---|---|---|---|---|---|---|
| DXE-01 | CRITICAL | data-loss | CONF | `DocxViewer.tsx:711-751`; `insertImage.ts:70-139` | Inserting an image operates on a stale `bundle.document` — every edit since file-load is silently discarded. | Call with `{...bundle, document: documentModel}`. | M |
| DXE-03 | CRITICAL | bug | CONF | `commands.ts:902-932`; `Input.ts:433-435` | Typing/deleting in any paragraph with a hyperlink fails silently while the DOM still visibly mutates. | Flatten hyperlink children in run-collection; always preventDefault on failure. | L |
| DXE-04 | CRITICAL | bug | CONF | `commands.ts:419-439,492-511,556-573` | Selections spanning more than one paragraph cannot be deleted/replaced/formatted. | Implement general cross-paragraph delete + generalize format application. | L |
| DXE-05 | CRITICAL | bug | CONF | `commands.ts:441-466`; `Input.ts:339-364` | Deleting/typing over a multi-run selection within one paragraph fails. | Generalize delete to slice every intersected run. | M |
| DXE-06 | CRITICAL | feature-gap | CONF | `commands.ts:75-94`; `toolbarAdapter.ts:208-237`; `Toolbar.tsx:188-351` | Lists/indent/table insert throw internally despite full toolbar wiring. | Implement for real, or disable with a "not yet supported" state. | XL |
| DXE-09 | CRITICAL ⇧was HIGH | data-loss | CONF | `DocxViewer.tsx:1406-1442` | Opening a different file unconditionally unmounts DocxEditor with no dirty flag or confirmation. | Track a dirty flag; confirm-before-switch. | M |
| DXE-02 | HIGH ⇩was CRITICAL | data-loss | CONF | `DocxViewer.tsx:619-637`; `Find.ts:304-320`; `commands.ts:451-453` | Multi-command ops push inverses before the whole batch is confirmed to succeed — a partial failure can corrupt the undo stack. | Build the full result first; only push inverses after full success. | M |
| DXE-07 | HIGH | bug | CONF | `DocxViewer.tsx:836-840` | Ctrl+S is captured, `preventDefault()`'d, and does nothing. | Call `handleSave()` directly. | S |
| DXE-08 | HIGH | data-loss | CONF | `useSpellCheck.ts:54-67`; `DocxViewer.tsx:1002-1011` | Accepting a spell-check suggestion bypasses the document model — lost on save. | Apply a real delete+insert command through the editor pipeline. | M |
| DXE-10 | HIGH | bug | CONF | `comments.ts:81-89`; `CommentsPane.tsx:114,132` | Raw-XML comment bodies always display "(empty comment)" (same bug as DXP-05/QA-02). | Fix the `Array.isArray` precedence. | S |
| DXE-11 | HIGH | feature-gap | CONF | `Toolbar.tsx:250-253`; `toolbarAdapter.ts:260-265`; `commandTypes.ts:83-139` | Track Changes Accept/Reject buttons map to `null` despite a working backend. | Wire `toolbarToCommand` to the existing revision-resolution logic. | M |
| DXE-12 | HIGH | feature-gap | CONF | `toolbarAdapter.ts:238-273`; `DocxViewer.tsx:1129-1131` | ~12 toolbar commands are silent no-ops (roughly half the ribbon UI). | Implement, or visually disable with a tooltip. | L |
| DXE-13 | HIGH | feature-gap | CONF | `commands.ts:873-900`; `Input.ts:42-54` | Headers/footers cannot be edited — no addressing-scheme path exists. | Extend paragraph-path addressing with a header/footer prefix. | L |
| DXE-14 | HIGH | feature-gap | CONF | `commandTypes.ts:99-113`; `Toolbar.tsx:309-351` | No table structural editing exists at all (row/col insert, merge/split, resize). | Add the missing command kinds. | XL |
| DXE-15 | HIGH | performance | CONF | `DocxViewer.tsx:792-805,1211-1266` | Every edit triggers a full un-debounced repagination — typed characters lag. | Debounce re-pagination; optimistic local echo. | M |
| DXE-16 | HIGH | bug | CONF | `History.ts:23-63`; `Input.ts:451-462` | Undo/redo does not restore the selection active when the undone edit was made. | Carry position in the inverse; restore on undo/redo. | M |
| DXE-17 | HIGH | ux | CONF | `DocxViewer.tsx:619-637`; `htmlPaste.ts:113-165`; `Find.ts:304-320` | Paste and Replace-All are not single atomic undo steps. | Introduce a composite/grouped Command. | M |
| DXE-18 | HIGH | feature-gap | CONF | `insertImage.ts:17-19,99-139,227`; `DocxViewer.tsx:735-736,746-747` | Inserted images bypass History, ignore cursor position and aspect ratio. | Route through Command/History; insert at cursor; decode natural size. | M |
| DXE-28 | HIGH | testing | CONF | `DocxViewer.editor.test.tsx:1-33`; `families.ts` | An unmocked font import leaves 0% of ~1500 lines of test code executing. | Mock `../../docx/fonts`, or fix the underlying import (resolved by DXP-10's fix). | S |
| DXE-19 | MEDIUM | feature-gap | UNV | `htmlPaste.ts:1-25` | Paste from Word/web drops tables, lists, hyperlinks, images, colors by design. | Extend the paste walker once tables/hyperlinks are editable. | XL |
| DXE-20 | MEDIUM | bug | UNV | `Composition.ts:51-68`; `DocxViewer.tsx:910-926` | IME composition can double-insert text on candidate cycling. | Track whether committed text was already applied before dispatching. | M |
| DXE-21 | MEDIUM | bug | UNV | `Input.ts:221-252` | Word-boundary delete never crosses a run or paragraph boundary. | Continue the search into adjacent runs/paragraphs. | M |
| DXE-22 | MEDIUM | architecture | UNV | `DocxViewer.tsx:542,1207-1209`; `useEditor.ts:16-19`; `useSelection.ts` | Dead duplicate "shadowEditor" state plus an unused `useSelection` hook. | Remove or wire in properly. | S |
| DXE-23 | MEDIUM | a11y | UNV | `Toolbar.tsx:145-351` | Color pickers/table-size picker not keyboard-operable; most buttons rely only on `title`. | Convert to real buttons with keyboard nav + `aria-label`. | M |
| DXE-24 | MEDIUM | feature-gap | UNV | `DocxViewer.tsx:1157-1182,1284-1287` | No "Save As" affordance anywhere. | Add an explicit Save-As action. | S |
| DXE-26 | MEDIUM | bug | UNV | `DocxViewer.tsx:1064-1098` | "Resolving" a comment is a client-only filter — reverts on reload. | Persist resolved state into the document model. | M |
| DXE-25 | LOW | ux | UNV | `DocxViewer.tsx:566,1383` | Save-failure feedback is a low-visibility banner the next keystroke clears. | Surface via a persistent toast/modal. | S |
| DXE-27 | LOW | tooling | UNV | `DocxViewer.tsx:1243-1257,1322,1408,1450`; `useSpellCheck.test.tsx:26-27` | Production console logging is the cause of 2 of the project's 6 lint errors. | Remove console calls; clean up unused directives/params. | S |

## 7. DOCX Serializer & Round-Trip Fidelity (`DXS`) — Health 2/10

The single worst finding across the whole audit lives here: saving any document with a header, footer, footnote, or endnote — i.e. most real Word documents — throws and the save fails outright. Every table loses its required `tblGrid`. Anchored images and non-picture drawings become schema-invalid or lose content. New comments/images are OPC-invalid orphans. The final disk write is a non-atomic, backup-less overwrite.

| ID | Sev | Category | Verify | Location | Impact | Recommendation | Eff |
|---|---|---|---|---|---|---|---|
| DXS-01 | CRITICAL | crash | CONF | `headers.ts:34-48`; `footers.ts:34-46`; `footnotes.ts:54-59,97-98`; `endnotes.ts:88-89` | Saving any document with a header, footer, footnote, or endnote throws and the save fails outright. | Finish the parser (same fix as DXP-03/04). | M |
| DXS-02 | CRITICAL | bug | CONF | `document.ts:530-558`; `documentWriter.ts:524-609`; `layoutTable.ts:38,172` | `w:tblGrid` (a required child) is silently dropped and never re-emitted — every table becomes schema-invalid. | Add `tblGrid` to the model; parse and emit it. | S |
| DXS-03 | CRITICAL | bug | CONF | `documentWriter.ts:411-430,457-470` | Anchored images written missing required `wp:anchor` child elements. | Extend `Drawing` to capture position/wrap; emit required children. | M |
| DXS-04 | CRITICAL | data-loss | CONF | `document.ts:427-458`; `documentWriter.ts:411-455` | Charts/SmartArt/other graphicFrame content silently dropped, replaced with a schema-invalid empty wrapper. | Preserve as `UnknownNode` when no `a:blip` is found. | S |
| DXS-05 | CRITICAL | fidelity | CONF | `styles.ts:393-402`; `stylesWriter.ts:441-456` | `w:tblStylePr` completely unmodeled — a table style's formatting is lost on every save. | Add `tblStylePr` field; parse/emit by type. | M |
| DXS-06 | CRITICAL | bug | CONF | `docx/index.ts:196-201`; `relsWriter.ts:84-98`; `contentTypesWriter.ts:105-114` | Adding the first comment never registers the relationship/content-type entry — OPC-invalid. | Call the already-existing registration helpers. | S |
| DXS-07 | CRITICAL | bug | CONF | `insertImage.ts:70-139`; `contentTypesWriter.ts:83-100` | Inserting an image into an image-free document never registers its content type. | Call `ensureMediaContentType`. | S |
| DXS-12 | CRITICAL | data-loss | CONF | `main.cjs:291-316` | DOCX save is a direct, non-atomic overwrite with no backup. | Write-temp-then-rename with a rolling `.bak`. | S |
| DXS-08 | HIGH ⇩was CRITICAL | bug | PART | `partWriterSupport.ts:61-75`; `commentsWriter.ts:18-29`; `documentWriter.ts:411-455` | Header/footer/footnote/comment parts never declare namespaces their own content can use. | Declare the full standard namespace set. | S |
| DXS-09 | HIGH | fidelity | CONF | `document.ts:427-458`; `documentWriter.ts:411-455` | Picture crop, border, effects, rotation discarded on every save. | Extend `Drawing` to carry the effect/crop/transform subtree. | L |
| DXS-10 | HIGH | fidelity | CONF | `document.ts:195-295`; `documentWriter.ts:211-266` | `w14:paraId`/`w14:textId`/`w:rsid*` stripped from every paragraph/run on every save. | Capture as optional fields; re-emit, generating fresh only for new nodes. | M |
| DXS-11 | HIGH | feature-gap | CONF | `document.ts:367-375`; `DocxViewer.tsx:1064-1098` | Comment "resolved" status neither modeled nor persisted. | Add a `resolved` field; parse/write `commentsExtended.xml`. | M |
| DXS-13 | MEDIUM | feature-gap | UNV | `docx/index.ts:172-251` | `docProps/core.xml`/`app.xml` never regenerated on save. | Regenerate `dcterms:modified`/`cp:lastModifiedBy`. | S |
| DXS-14 | MEDIUM | fidelity | UNV | `stylesWriter.ts:41-52` | `w:latentStyles` dropped from `styles.xml` on every save. | Capture as an opaque passthrough fragment. | S |
| DXS-15 | MEDIUM | architecture | UNV | `documentWriter.ts:1123-1150` | Document root always emits a fixed, hardcoded namespace set. | Capture the source root's actual attributes; union with baseline. | M |
| DXS-16 | MEDIUM | architecture | UNV | `document.ts:1393-1398` | Unknown nodes reconstructed from parsed tree, not raw source — loses ancestor namespaces. | Capture as raw substrings where feasible. | M |
| DXS-17 | MEDIUM | testing | UNV | `documentWriter.test.ts:269-285`; `headerWriter.test.ts:18-48`; `commentsWriter.test.ts:1-81` | Zero realistic round-trip tests — fixtures specifically avoid the bugs found in this audit. | Build a real-fixture corpus + genuine round-trip test. | M |
| DXS-19 | MEDIUM | architecture | UNV | `docx/index.ts:172-251`; `DocxViewer.tsx:1157-1182` | No post-serialization validation — corrupted output ships silently. | Add a cheap re-parse + invariant-check pass before/after commit. | M |
| DXS-18 | LOW | tech-debt | UNV | `relsWriter.ts:84-98`; `contentTypesWriter.ts:83-114` | The helpers needed to fix DXS-06/07 already exist and are tested but unused. | Wire them into the mutation paths — no new low-level code. | S |
| DXS-20 | LOW | feature-gap | UNV | `document.ts:132-207`; `parser/document.ts:217-246,268-289` | Fields (TOC, PAGE, REF) preserved as opaque XML, never regenerated. | Document as a limitation; model fields explicitly if ever a goal. | L |

## 8. PDF Viewer (`PDF`) — Health 3/10

Solid foundations fall badly short of what a professional viewer needs: no page virtualization at all (a genuine memory-exhaustion risk), no text selection/copy, no clickable links, no search, no printing, no forms. Several concrete bugs undermine what does exist.

| ID | Sev | Category | Verify | Location | Impact | Recommendation | Eff |
|---|---|---|---|---|---|---|---|
| PDF-01 | CRITICAL | performance | CONF | `PdfViewer.tsx:366-484,431-471,444-455` | No page virtualization — a 500-page PDF can allocate GB-scale canvas backing store. | Implement windowed rendering off the existing `IntersectionObserver`. | L |
| PDF-02 | HIGH | bug | CONF | `PdfViewer.tsx:487-505` | Fit-Width/Fit-Page never re-fits on resize (identity-preserving state update). | Drive re-fit off a value that actually changes. | S |
| PDF-03 | HIGH | bug | CONF | `PdfViewer.tsx:234-270,512,524-530` | Global shortcuts hijack the page-number input. | Add an `inField` guard excluding the input. | S |
| PDF-04 | HIGH | feature-gap | CONF | `PdfViewer.tsx:431-471` | No text layer — text can't be selected/copied/read by screen readers. | Render a `TextLayerBuilder` overlay. | L |
| PDF-05 | HIGH | feature-gap | CONF | `PdfViewer.tsx:431-471` | No annotation layer — links are inert. | Render pdf.js's `AnnotationLayer`. | M |
| PDF-06 | HIGH | feature-gap | CONF | `App.tsx:203,304`; `useSearch.ts:3` | PDF search entirely missing. | Build a search index from `getTextContent()`. | M |
| PDF-07 | HIGH | bug | CONF | `PdfViewer.tsx:175-196,431-471` | Jumping to a not-yet-rendered page silently does nothing. | Queue the pending page; prioritize/render on demand. | M |
| PDF-09 | HIGH | fidelity | CONF | `PdfViewer.tsx:411-429,438` | Fit scale computed once from page 1, reused for every page. | Compute fit scale per-page. | M |
| PDF-10 | HIGH | feature-gap | CONF | `PdfViewer.tsx` (no print action) | No way to print a PDF — Ctrl+P re-downloads instead. | Add a Print button; make Ctrl+P format-aware. | M |
| PDF-08 | MEDIUM ⇩was HIGH | bug | PART | `PdfViewer.tsx:384-405` | Current-page tracking resets baseline to 0 each callback — can report the wrong page while scrolling. | Maintain a persistent ratio map. | S |
| PDF-11 | MEDIUM | feature-gap | UNV | `PdfViewer.tsx:431-471` | No AcroForm/interactive form field rendering. | Enable `renderForms:true` once annotation layer exists. | L |
| PDF-12 | MEDIUM | feature-gap | UNV | `PdfViewer.tsx` (no rotation state) | No manual page-rotation control. | Add a rotation state + toolbar button. | S |
| PDF-13 | MEDIUM | feature-gap | UNV | `PdfViewer.tsx:102-116` | No page thumbnails — fallback nav is a plain text list. | Generate lazy low-scale thumbnails. | M |
| PDF-14 | MEDIUM | feature-gap | UNV | `PdfViewer.tsx:292-348` | Password-protected PDFs cannot be opened at all. | Wire `onPassword` to a password-entry modal. | M |
| PDF-15 | MEDIUM | architecture | UNV | `PdfViewer.tsx:298-301`; `main.cjs:145-150` | Module-worker loading under packaged `file://` is unverified. | Add a packaged-build test; warn on fake-worker fallback. | S |
| PDF-16 | MEDIUM | testing | UNV | `PdfViewer.tsx`; `smoke.spec.ts:32-61` | Zero meaningful automated coverage beyond a single-page smoke check. | Export helpers for unit testing; extend the e2e fixture. | M |
| PDF-17 | LOW | tech-debt | UNV | `package.json:36,40` | Unused `react-pdf` pins a different, older `pdfjs-dist`. | `npm uninstall react-pdf`. | S |
| PDF-18 | LOW | fidelity | UNV | `PdfViewer.tsx:444-455` | Canvas resolution doesn't update on a live DPR change. | Listen for DPR changes via `matchMedia`. | S |
| PDF-19 | LOW | tech-debt | UNV | `viewer-pdf.css:148-154` | Dead CSS transition rule with no effect. | Remove the unused rule. | S |

## 9. PPTX & ODP Slide Viewers (`SLD`) — Health 3/10

A bare text-and-image extractor, not a slide renderer. Slide masters/layouts are never read; run-level formatting, bullets, shape fills, rotation, cropping, and group transforms are absent or wrong; tables flatten into an unordered stack of cell text; soft line breaks corrupt text; export silently produces a one-slide PDF instead of the deck.

| ID | Sev | Category | Verify | Location | Impact | Recommendation | Eff |
|---|---|---|---|---|---|---|---|
| SLD-01 | CRITICAL | data-loss | CONF | `App.tsx:231-239,372-374`; `export.ts:276-341`; `SlideDeck.tsx:157-239` | Export to PDF captures only whichever slide is on screen — a 20-slide deck exports as 1-2 pages, no warning. | Build a dedicated per-slide export. | M |
| SLD-02 | HIGH | fidelity | CONF | `PptxViewer.tsx:394-534,304-320`; `SlideDeck.tsx:120-127` | Slide masters/layouts never read — placeholders collapse into a generic stacked list. | Resolve the layout→master chain; implement placeholder matching. | L |
| SLD-03 | HIGH | fidelity | CONF | `OdpViewer.tsx:354-455,358,365-371` | ODP master pages/`styles.xml` never opened — falls back to a hardcoded default. | Parse `styles.xml`'s master-page/page-layout data. | L |
| SLD-04 | HIGH | fidelity | CONF | `PptxViewer.tsx:322-334`; `OdpViewer.tsx:385-406` | Text run formatting (bold/italic/color/size) completely discarded. | Extend `SlideData` with styled runs; render per-run spans. | L |
| SLD-05 | HIGH | fidelity | CONF | `PptxViewer.tsx:322-334`; `OdpViewer.tsx:387-394` | Bullets/numbering never rendered. | Read bullet/level per paragraph; render prefix + indent. | M |
| SLD-06 | HIGH | bug | CONF | `PptxViewer.tsx:322-334`; `OdpViewer.tsx:387-394` | Soft line breaks dropped, fusing lines into one run-on string (empirically reproduced). | Walk children in document order, inserting `\n` at breaks. | S |
| SLD-07 | HIGH | fidelity | CONF | `PptxViewer.tsx:459-495`; `OdpViewer.tsx:385-441` | Tables not recognized — each cell becomes an independent stacked text box. | Add explicit table handling into a 2D grid. | M |
| SLD-08 | HIGH | bug | CONF | `PptxViewer.tsx:304-320,459-521` | Grouped shapes never compose the ancestor transform (empirically reproduced). | Traverse `p:spTree` composing group transforms. | M |
| SLD-09 | HIGH | fidelity | CONF | `SlideDeck.tsx:75-155`; `PptxViewer.tsx:459-482`; `OdpViewer.tsx:385-406` | Shapes render as bare text with no fill/border/geometry/background. | Parse `spPr` fill/line/geometry; apply backgrounds. | L |
| SLD-24 | HIGH | bug | CONF | `PptxViewer.tsx:435-455`; `OdpViewer.tsx:376-384` | One malformed slide part aborts the entire deck, discarding every parsed slide. | Wrap each slide iteration in its own try/catch. | S |
| SLD-10 | MEDIUM | fidelity | UNV | `PptxViewer.tsx:499-521`; `SlideDeck.tsx:101-119` | Picture cropping (`a:srcRect`) ignored. | Apply crop via CSS or pre-crop into a canvas. | S |
| SLD-11 | MEDIUM | fidelity | UNV | `PptxViewer.tsx:304-320` | Shape/picture rotation and flip ignored. | Apply `transform: rotate() scale()`. | S |
| SLD-12 | MEDIUM | feature-gap | UNV | `PptxViewer.tsx:394-534`; `OdpViewer.tsx:354-455` | Charts/SmartArt/OLE dropped with no placeholder. | Render a labeled "not supported" placeholder. | S |
| SLD-13 | MEDIUM | feature-gap | UNV | `PptxViewer.tsx`; `OdpViewer.tsx`; `SlideDeck.tsx` | Speaker notes never parsed or exposed. | Parse notes relationship; add a notes drawer. | M |
| SLD-14 | MEDIUM | feature-gap | UNV | `SlideDeck.tsx`; `PptxViewer.tsx`; `OdpViewer.tsx` | No presentation/fullscreen mode. | Add a "Present" control with `requestFullscreen()`. | M |
| SLD-15 | MEDIUM | feature-gap | UNV | `PptxViewer.tsx:435-531`; `OdpViewer.tsx:376-452` | Hidden slides always shown. | Read `show`; exclude hidden by default. | S |
| SLD-16 | MEDIUM | ux | UNV | `PptxViewer.tsx:70-103`; `OdpViewer.tsx:70-103` | Keyboard nav limited to PageUp/PageDown. | Extend to Arrow/Space/Home/End. | S |
| SLD-17 | MEDIUM | ux | UNV | `App.tsx:155,336-338`; `Sidebar.tsx:13-45`; `SlideDeck.tsx:171-216` | App sidebar and thumbnail rail both show a redundant slide list. | Suppress the app-level Sidebar for slide formats. | S |
| SLD-18 | MEDIUM | ux | UNV | `PptxViewer.tsx:37-47`; `OdpViewer.tsx:37-47` | Nav labels always generic "Slide N," never the actual title. | Extract a title-placeholder's text for the label. | S |
| SLD-19 | MEDIUM | performance | UNV | `PptxViewer.tsx:435-531`; `OdpViewer.tsx:376-452`; `SlideDeck.tsx:185-215` | Large decks parse serially; every thumbnail unvirtualized. | Parallelize parsing; virtualize the rail. | M |
| SLD-20 | MEDIUM | a11y | UNV | `SlideDeck.tsx:106`; `PptxViewer.tsx:499-521`; `OdpViewer.tsx:408-420` | Image `alt` text always hardcoded empty. | Read `descr`/`svg:title`; pass through to `<img alt>`. | S |
| SLD-21 | MEDIUM | fidelity | UNV | `PptxViewer.tsx:322-334`; `SlideDeck.tsx:77-86,129-148` | No text-autofit support — overflow silently clipped. | Read `a:normAutofit`; scroll or indicate truncation. | M |
| SLD-22 | LOW | fidelity | UNV | `PptxViewer.tsx:333`; `OdpViewer.tsx:394` | Paragraph spacing discarded — one fixed newline joins all paragraphs. | Render each paragraph with its own margin. | S |
| SLD-23 | LOW | ux | UNV | `SlideDeck.tsx:58-73` | No manual zoom controls beyond auto-fit. | Add zoom in/out/reset controls. | S |

## 10. Spreadsheet, CSV, Text, Code, RTF & ODT Viewers (`DAT`) — Health 3/10

Three of six format families have a CRITICAL, empirically-reproduced defect: Code/RTF/ODT content taller than one screen is silently clipped with no scrollbar. Spreadsheet parsing sits on a known-vulnerable, npm-frozen library with public unpatched CVEs. Grid copy is non-functional; cells look editable but silently discard input; dates/currency/errors render wrong in every formatted workbook.

| ID | Sev | Category | Verify | Location | Impact | Recommendation | Eff |
|---|---|---|---|---|---|---|---|
| DAT-01 | CRITICAL | security | CONF | `package.json:51`; `SpreadsheetViewer.tsx:48-49` | `xlsx@0.18.5` is npm's final, permanently-frozen release with unpatched CVEs. | Move to SheetJS's own CDN-distributed patched build. | L |
| DAT-02 | CRITICAL | bug | CONF | `CodeViewer.tsx:116-127`; `RtfViewer.tsx:76`; `OdtViewer.tsx:79`; `index.css:318-334` | CSS specificity fully clips Code/RTF/ODT past one screen — no scrollbar at all. | Add dedicated `overflow:auto` stylesheets. | S |
| DAT-03 | HIGH | bug | CONF | `main.cjs:63-71,258-262`; `useFileHandler.ts:10-12` | Text/CSV/TSV/code always decoded as UTF-8 with no BOM detection. | Sniff/strip BOM; fall back to Windows-1252. | M |
| DAT-04 | HIGH | fidelity | CONF | `SpreadsheetViewer.tsx:55` | `sheet_to_json` defaults to raw values — dates show as serial numbers, errors as blank. | Pass `raw:false`; render `.w` for error cells. | S |
| DAT-05 | HIGH | feature-gap | CONF | `SpreadsheetViewer.tsx:226-241`; `CsvViewer.tsx:212-227` | Copying a cell range does not work at all. | Pass `getCellsForSelection={true}`. | S |
| DAT-06 | HIGH | bug | CONF | `SpreadsheetViewer.tsx:132`; `CsvViewer.tsx:124` | Cells look editable but `onCellEdited` is never wired — typed edits silently vanish. | Set `allowOverlay:false` to be honest, or wire real editing+save. | S |
| DAT-07 | HIGH | performance | CONF | `SpreadsheetViewer.tsx:48-58,198-200`; `CsvViewer.tsx:47-57,185-187` | Parsing runs synchronously with no loading indicator — 100k-row file can freeze the UI. | Enable worker parsing; add a loading state. | M |
| DAT-13 | HIGH | performance | CONF | `CodeViewer.tsx:74-127` | No virtualization; shiki tokenizes the whole file synchronously. | Fall back to virtualized plain text above a size threshold. | M |
| DAT-08 | MEDIUM | performance | UNV | `SpreadsheetViewer.tsx:49` | Expensive parse options requested and immediately discarded. | Drop `cellStyles`/`cellFormula`/`sheetStubs`. | S |
| DAT-09 | MEDIUM | fidelity | UNV | `SpreadsheetViewer.tsx` | Merged cells not read — data appears missing outside the top-left cell. | Read `ws['!merges']`; backfill or use span cells. | M |
| DAT-10 | MEDIUM | fidelity | UNV | `SpreadsheetViewer.tsx:175,236` | Column widths/row heights/frozen panes ignored. | Seed from `ws['!cols']`/`ws['!rows']`. | M |
| DAT-11 | MEDIUM | fidelity | UNV | `SpreadsheetViewer.tsx:53` | Hidden sheets always shown as ordinary tabs. | Filter by the `Hidden` flag by default. | S |
| DAT-12 | MEDIUM | tech-debt | UNV | `SpreadsheetViewer.tsx:146-192`; `CsvViewer.tsx:133-179` | SpreadsheetViewer/CsvViewer ~90% duplicated code. | Extract a shared grid hook/module. | M |
| DAT-14 | MEDIUM | feature-gap | UNV | `extToLang.ts:1-43`; `detect.ts:20-80` | 22 of 61 code extensions get zero syntax highlighting. | Generate from the same source as `detect.ts`. | S |
| DAT-15 | MEDIUM | feature-gap | UNV | `App.tsx:304,323`; `useSearch.ts:3` | No find/search in Text/Code/RTF/ODT. | Extend `useSearch` to these formats. | M |
| DAT-16 | MEDIUM ⇩was HIGH | testing | CONF | `src/viewers/__tests__/` (only 2 unrelated test files exist) | Zero coverage for every parsing/rendering component in this subsystem. | Add fixture-based render tests per viewer. | L |
| DAT-17 | MEDIUM | fidelity | UNV | `RtfViewer.tsx:56`; `OdtViewer.tsx:44` | RTF/ODT status-bar always hardcodes `pages:1`. | Derive an approximate page count. | S |
| DAT-18 | MEDIUM | fidelity | UNV | `OdtViewer.tsx:29`; odf-kit `parser.js:1364` | Tracked changes always resolved as accepted by default. | Pass `trackedChanges:'changes'` or surface a toggle. | S |
| DAT-19 | MEDIUM | architecture | UNV | `package.json:34`; `OdtViewer.tsx:21-24` | ODT fidelity rests on a pre-1.0 dependency with no local fixtures. | Pin the exact version; add a fixture corpus. | M |
| DAT-20 | MEDIUM | performance | UNV | `RtfViewer.tsx:41-49`; `OdtViewer.tsx:29` | No timeout/size guard against pathological input. | Wrap conversion in a Worker or timeout race. | M |
| DAT-21 | LOW | ux | UNV | `ViewerErrorBoundary.tsx:31-34`; `ViewerRouter.tsx:12-23` | "Try again" can't actually retry a deterministic failure. | Bust the lazy-import cache on retry. | S |
| DAT-22 | LOW | tech-debt | UNV | `main.cjs:63-71,258-262`; `useFileHandler.ts:10-12` | Shared text reader misleadingly named `readMarkdownFile()`. | Rename to `readTextFile`. | S |

## 11. Export Pipeline, Theming, Styling & Accessibility (`UX`) — Health 3/10

Two CRITICAL, confirmed bugs mean "Export to PDF" is broken for essentially every non-markdown format; a third CRITICAL bug (Ctrl+P double-fire) can corrupt output on the flagship DOCX editor; Code/RTF/ODT/Unknown viewers have literally no CSS at all. Accessibility is inconsistently applied — the DOCX toolbar has ~20 unlabeled icon buttons and a core text color fails contrast in most themes.

| ID | Sev | Category | Verify | Location | Impact | Recommendation | Eff |
|---|---|---|---|---|---|---|---|
| UX-01 | CRITICAL | bug | CONF | `App.tsx:236-239`; `export.ts:276-341`; `page-view.css:16-23` | Export to PDF rasterizes only the currently on-screen viewport for every non-markdown, non-PDF format. | Reuse each viewer's own full-content path instead of screenshotting. | L |
| UX-02 | CRITICAL | bug | CONF | `SlideDeck.tsx:157-238`; `PptxViewer.tsx:180-188`; `OdpViewer.tsx:180-188` | Same defect as SLD-01 from the export-pipeline angle. | Same fix as SLD-01. | M |
| UX-04 | CRITICAL | bug | CONF | `CodeViewer.tsx:116-127`; `RtfViewer.tsx:76`; `OdtViewer.tsx:79`; `UnknownViewer.tsx:9-16` | Same defect as DAT-02, confirmed exhaustively via grep. | Same fix as DAT-02. | S |
| UX-03 | HIGH ⇩was CRITICAL | bug | CONF | `DocxViewer.tsx:812-847,1184-1197,1321`; `useUniversalShortcuts.ts:49-53` | Ctrl+P in DOCX fires both the native print dialog AND the broken export simultaneously. | `stopPropagation()` in DocxViewer's Ctrl+P branch. | S |
| UX-05 | HIGH | fidelity | CONF | `export.ts:1-16,228-260,616-653`; `MarkdownRenderer.tsx:1-10,59-60` | HTML/DOCX export uses a separate parser than the live preview — math/Mermaid export as raw syntax text. | Serialize the already-rendered DOM instead of re-parsing with `marked`. | M |
| UX-07 | HIGH | bug | CONF | `useFontSize.ts:27`; `index.css:36,603`; `Toolbar.tsx:136-161` | Same defect as SHELL-14 — font-size controls dead everywhere but markdown. | Guard the toolbar group with the `isMarkdown` pattern. | S |
| UX-08 | HIGH | bug | CONF | `find-replace.css:11-13,41-43`; `page-view.css:7,11` | Orphaned CSS vars break DOCX Find & Replace and page backdrop in every non-light theme. | Rename to the app's real tokens. | S |
| UX-09 | HIGH | a11y | CONF | `ExportMenu.tsx:68-77`; `ThemeMenu.tsx:38-46`; `Toolbar.tsx:145-256`; `PdfViewer.tsx:515-604` | ~20+ icon-only buttons rely solely on `title`, no `aria-label`. | Add `aria-label` mirroring existing `title` text. | S |
| UX-10 | HIGH | a11y | CONF | `index.css:21,68,111,195` | `--text-tertiary` fails WCAG AA contrast (as low as 2.89:1) in 4 of 5 themes, 26 usages. | Darken/lighten per theme to reach ≥4.5:1. | S |
| UX-06 | MEDIUM | fidelity | UNV | `export.ts:454-501`; `index.css:700-706` | DOCX export silently drops task-list checkbox state. | Branch on `item.task`/`item.checked`. | S |
| UX-11 | MEDIUM | ux | UNV | `export.ts:22-34`; `App.tsx:52-62`; `DocxViewer.tsx:1165` | Export never uses Electron's native save dialog. | Route through `saveBinaryFile`/`saveFile` when in Electron. | M |
| UX-12 | MEDIUM | feature-gap | UNV | `ExportMenu.tsx:49-64` | Export menu offers only a generic, often-broken "Export to PDF." | Add format-specific items (CSV export, real per-slide export). | M |
| UX-13 | MEDIUM | a11y | UNV | `ShortcutsModal.tsx:36-44` | Modal has `aria-modal` but no focus trap/initial focus/restoration. | Trap Tab; move focus on open; restore on close. | S |
| UX-14 | MEDIUM | a11y | UNV | `ExportMenu.tsx:20-92`; `ThemeMenu.tsx:11-70`; `index.css:1343-1348` | Dropdowns declare `role=menu` but implement none of the keyboard pattern. | Implement arrow-key nav, or drop the ARIA role. | M |
| UX-15 | MEDIUM | a11y | UNV | `SearchOverlay.tsx:60-64`; `index.css:896-905` | Search match count has no `aria-live`; input has no visible focus indicator. | Add `aria-live`; add a visible `:focus` style. | S |
| UX-16 | MEDIUM | ux | UNV | `DropZone.tsx:10-18`; `useFileHandler.ts:63-93` | Drop-zone copy claims markdown-only despite full format detection. | Update DropZone copy. | S |
| UX-17 | MEDIUM | ux | UNV | `ViewerErrorBoundary.tsx:36-48`; `ViewerLoading.tsx:8-14`; `UnknownViewer.tsx:9-16` | Generic fallback UIs are completely unstyled. | Add a shared themed stylesheet. | S |
| UX-18 | LOW | ux | UNV | `App.tsx:255-258` | Export failures use a native, theme-ignoring `alert()`. | Replace with a themed toast. | S |
| UX-19 | LOW | tech-debt | UNV | `index.css:354-363` | Dead `.docx-wrapper` CSS selectors match nothing. | Delete the dead rules. | S |
| UX-20 | LOW | tech-debt | UNV | `export.ts:46-196`; `index.css:630-635` | Duplicated markdown export CSS already drifted from live preview. | Generate from the same source, or cross-reference. | S |
| UX-21 | LOW | a11y | UNV | `index.css:241,1206-1231,1367-1370`; `App.tsx:172` | No `prefers-reduced-motion` support anywhere. | Wrap animations in a media guard. | S |
| UX-22 | LOW | tech-debt | UNV | `package.json` | Three raster libs installed, only one imported. | Remove the three unused dependencies. | S |
| UX-23 | LOW | tech-debt | UNV | `useTheme.ts:20`; `electron.d.ts:35` | Same defect as SHELL-23. | Remove the cast. | S |
| UX-24 | LOW | docs | UNV | `useUniversalShortcuts.ts:49-53`; `ShortcutsModal.tsx:46-54` | Ctrl+P is a live shortcut, completely undocumented. | Add a shortcuts-modal entry. | S |

## 12. Tests, Tooling, Dependencies & Project Hygiene (`QA`) — Health 3/10

No version control, no CI, both lint and tests currently red. The single riskiest engine (DOCX layout/render) runs with 0% real coverage due to a fixable asset-resolution bug. 9 of 15 viewers and the entire shell/IPC layer have zero tests. TypeScript strict mode is entirely off. Two directly-used production dependencies carry known, largely-unpatched security advisories. The project's own planning docs overstate its health with no mechanism keeping that claim honest.

| ID | Sev | Category | Verify | Location | Impact | Recommendation | Eff |
|---|---|---|---|---|---|---|---|
| QA-22 | HIGH | tooling | CONF | (project root — no `.git`) | No version control at all. | `git init`; commit baseline. | S |
| QA-23 | HIGH | tooling | CONF | (no CI config) | No CI/CD pipeline of any kind. | Minimal `lint+tsc+test+build` pipeline. | M |
| QA-01 | HIGH | testing | CONF | `families.ts:7-28`; `vitest.config.ts` | Same root cause as DXP-10, at the tooling level. | Same fix as DXP-10. | S |
| QA-05 | HIGH | testing | CONF | `App.tsx`(392); `main.cjs`(455); `preload.cjs`(43) | Shell + IPC layer has zero unit tests. | RTL + mocked-IPC test suites. | L |
| QA-06 | HIGH | testing | CONF | `PdfViewer.tsx`(613), `PptxViewer.tsx`(536), etc. | 9 of 15 viewers have zero test coverage. | Prioritize hand-parsed formats first. | L |
| QA-07 | HIGH | testing | CONF | `MarkdownRenderer.tsx`, `Mermaid.tsx`, `RawEditor.tsx`, `useToc.ts`, `ExportMenu.tsx`, `export.ts` | The "perfect" feature has the least verification of anything in the app. | Characterization/snapshot tests to freeze current output. | M |
| QA-10 | HIGH | tooling | CONF | `tsconfig.app.json:2-23`; `tsconfig.node.json:2-22` | TypeScript strict family entirely disabled. | Ratchet `strictNullChecks` in incrementally, riskiest code first. | XL |
| QA-12 | HIGH ⇩was CRITICAL | security | PART | `package.json:36,51`; `dompurify:23` | `pdfjs-dist`/`xlsx` in vulnerable ranges, both exercised on untrusted files. | Upgrade pdfjs-dist; pair with DAT-01. | L |
| QA-02 | MEDIUM | bug | UNV | `comments.ts:83-89`; `comments.test.ts:79-119` | Same defect as DXP-05/DXE-10. | Fix the precedence bug. | S |
| QA-03 | MEDIUM | tooling | UNV | `vitest.config.ts:1-14` | Coverage report silently suppressed when any test fails. | Add `reportOnFailure:true`. | S |
| QA-04 | MEDIUM | testing | UNV | `vitest.config.ts:1-14` | No coverage thresholds configured anywhere. | Add thresholds once trustworthy; ratchet up. | M |
| QA-08 | MEDIUM | testing | UNV | `smoke.spec.ts:1-62`; `playwright.config.ts` | E2E suite only checks mount + zero console errors. | Extend with real scenario specs. | L |
| QA-09 | MEDIUM | testing | UNV | `src/docx/`; `tests/e2e/fixtures/` (1 sample.docx) | No DOCX round-trip fidelity corpus or visual regression testing. | Build a 10-20 file real-document corpus. | XL |
| QA-11 | MEDIUM | tooling | UNV | `useSpellCheck.test.tsx:26-27`; `canvasMetrics.ts:249`; `DocxViewer.tsx:1243,1255,1322,1408,1450` | ESLint red: 6 errors + 2 warnings, including a real perf-risk violation. | Fix each per ESLint's own guidance. | S |
| QA-13 | MEDIUM | security | CONF | `package.json:68` (electron ^35.1.2) | Electron ~9 majors behind, ~35 advisories. | Stage a major-version upgrade as its own epic. | XL |
| QA-14 | MEDIUM | tech-debt | UNV | `package.json:21,37-38,32` | glide-data-grid's peer deps violated by installed versions. | Pin via `overrides`; add viewer tests. | M |
| QA-17 | MEDIUM | tooling | UNV | `check-bundle.mjs:1-91`; `atlas-phase0.json` | Bundle regression gate orphaned and stale. | Wire into `postbuild`; refresh baseline. | S |
| QA-18 | MEDIUM ⇩was HIGH | docs | PART | `atlas-phase1.md:10-12,29,48,55,68` | "COMPLETE, 151/151 tests" repeatedly asserted while actually red. | Cross-check plan claims against a fresh run. | S |
| QA-19 | MEDIUM | docs | UNV | `atlas-phase2-docx.md:5,44-47`; `useSpellCheck.ts:1-30` | Phase 2's locked library commitments silently abandoned, undocumented. | Write a retroactive "Actual State" addendum. | M |
| QA-21 | MEDIUM | docs | UNV | `README.md:39,61,95-96` | README lists nonexistent dependencies and a wrong shortcut table. | Correct against the real source of truth. | S |
| QA-25 | MEDIUM | testing | UNV | `setup.ts:1`; `useFileHandler.test.ts:32-44` | Zero shared test mocks — every file hand-rolls its own. | Extract a shared typed `createMockElectronAPI()`. | M |
| QA-26 | MEDIUM | testing | UNV | `useFileHandler.ts:73,82,91,134`; `useFileHandler.test.ts:67-268` | Non-Electron browser-mode crash path completely untested. | Add a test asserting graceful failure. | S |
| QA-15 | LOW | tech-debt | UNV | `package.json` | Four dead dependencies inflate install size/audit surface. | Remove all four. | S |
| QA-16 | LOW | tech-debt | UNV | `node_modules/@emnapi/runtime` | Extraneous package signals a drifted install. | Run `npm ci`; confirm it disappears. | S |
| QA-20 | LOW | docs | UNV | `.sisyphus/baselines/` (only phase0 present) | Phase 1's own closing artifact was never produced. | Produce retroactively or drop the requirement. | S |
| QA-24 | LOW | docs | UNV | `package.json:4`; no `CHANGELOG.md` | No changelog across 9 undocumented patch releases. | Adopt `CHANGELOG.md` once git exists. | S |
| QA-27 | LOW | testing | UNV | `useRecentFiles.ts` (no dedicated test) | Only incidental coverage via another hook's test. | Add a dedicated test. | S |
| QA-28 | LOW | tooling | UNV | `package.json` (no `engines`); no `.nvmrc` | No toolchain pinning despite a bleeding-edge stack. | Add `engines`/`.nvmrc`. | S |

## 13. Live Runtime QA — Build + Drive the Real Electron App (`RUN`) — Health 3/10

Every one of 13 fixture formats loads and renders with zero console errors on the happy path — but the moment the app is probed past that path, the shell around every non-markdown viewer (and, for file-switching, markdown too) breaks down. All 17 findings here were reproduced live in a running Electron instance.

| ID | Sev | Category | Verify | Location | Impact | Recommendation | Eff |
|---|---|---|---|---|---|---|---|
| RUN-01 | CRITICAL | data-loss | RUNTIME | `App.tsx:86-90,339-346`; `ViewerRouter.tsx:12-20` | Runtime-reproduced: typed into markdown, opened a DOCX via the OS "Open with" path — no dialog, edits vanished. | One cross-format confirm-discard gate before any open proceeds. | M |
| RUN-04 | CRITICAL | bug | RUNTIME | `App.tsx:135-142`; `WelcomeScreen.tsx:53-55` | Runtime-reproduced: the Welcome screen's own headlined drag-and-drop feature is completely dead. | `webUtils.getPathForFile` preload bridge. | S |
| RUN-05 | CRITICAL | fidelity | RUNTIME | `main.cjs:63-71,258-262` | Runtime-reproduced with generated fixtures: accented characters replaced by U+FFFD glyphs. | Same fix as LOAD-05/DAT-03. | M |
| RUN-02 | HIGH | bug | RUNTIME | `useUniversalShortcuts.ts:70-73`; `DocxViewer.tsx:812-900` | Runtime-reproduced: Ctrl+B in DOCX visibly toggles the sidebar closed. | Same fix as SHELL-08. | S |
| RUN-03 | HIGH | bug | RUNTIME | `DocxViewer.tsx:836-840`; `App.tsx:102-112` | Runtime-reproduced: Ctrl+S in DOCX does nothing at all. | Same fix as DXE-07/SHELL-10. | S |
| RUN-07 | HIGH | data-loss | RUNTIME | `export.ts:276-341`; `App.tsx:236-239`; `SpreadsheetViewer.tsx:12-16` | Code-verified against a live 394,501-row CSV: export always rasterizes only the painted canvas region. | Dedicated tabular export off parsed rows. | L |
| RUN-08 | HIGH | feature-gap | RUNTIME | `useSearch.ts:115-125`; `App.tsx:203` | Runtime-reproduced: Ctrl+F on pdf/xlsx/csv/code produces no UI and suppresses native find via `preventDefault()`. | Extend search; stop `preventDefault()` when disabled. | M |
| RUN-09 | HIGH | bug | RUNTIME | `ViewerRouter.tsx:22-30`; `ViewerErrorBoundary.tsx:36-48` | Code-traced live: no `key` tied to `file.path` — a crash keeps showing "Viewer crashed" after a different valid file is opened. | Same fix as LOAD-08. | S |
| RUN-10 | HIGH | data-loss | RUNTIME | `DocxViewer.tsx:534,1157-1182`; `main.cjs:291-316` | Code-traced live: DocxViewer's Save always writes directly over the original path with no backup. | Same fix as DXS-12. | S |
| RUN-06 | MEDIUM | bug | RUNTIME | `SpreadsheetViewer.tsx:45-68` | Runtime-reproduced: a 0-byte `.xlsx` renders as a clean-looking empty spreadsheet, no error. | Add a post-parse sanity check routed through `setError`. | S |
| RUN-11 | MEDIUM | testing | RUNTIME | `smoke.spec.ts:54`; `App.tsx:13`; `registry.ts:16-29` | The only e2e smoke test's markdown case reliably fails on a clean run (cold dev-server compile, not a product bug). | Bump the timeout to match the statusbar assertion. | S |
| RUN-12 | MEDIUM | tooling | RUNTIME | `main.cjs:29,145-150`; `package.json:18` | `electron:preview` doesn't actually preview the production build. | Key the load path off `dist/index.html` existing. | S |
| RUN-13 | MEDIUM | feature-gap | RUNTIME | `ExportMenu.tsx:49-64`; `WelcomeScreen.tsx:100` | Runtime-confirmed: Welcome screen overpromises export capability. | Scope the copy accurately, or extend the menu. | L |
| RUN-14 | LOW | ux | RUNTIME | `DocxViewer.tsx`; `PdfViewer.tsx` | Raw library exceptions shown verbatim to end users. | Wrap in friendly, format-specific messages. | S |
| RUN-15 | MEDIUM | performance | RUNTIME | `dist/assets/rtf.js-*.js` (2.24MB min); `check-bundle.mjs:28-46` | `rtf.js` is the single largest asset in the build, for the least-used format. | Investigate a lighter converter; extend the bundle gate. | M |
| RUN-16 | MEDIUM | bug | RUNTIME | `App.tsx:66`; `useFileHandler.ts:57`; `useRecentFiles.ts:33-46` | Same duplicate-hook defect as SHELL-06/LOAD-16, confirmed live. | Same fix — single instance, passed down. | S |
| RUN-17 | MEDIUM | ux | RUNTIME | `useFileHandler.ts:22-26,51-52,64-123`; `App.tsx:69-76` | Same defect as SHELL-07/LOAD-06, confirmed live. | Same fix — destructure and surface both. | S |

---

## 14. Owner-Reported DOCX Editor Issues (`USR`) — reported 2026-09-15 on Atlas 3.1.0 (`487d9fc`)

Reported by the owner while editing a real document (`C:\Users\Youssef\Downloads\Youssef_Melki_Common_App_Essay.docx` — personal content: copy it to a temp folder for reproduction, never commit it). **Status: open, not yet fixed** (owner asked to record them without launching the fix). `Verify` = `OWNER` (reported by the owner, reproduction pending) or `SCREEN` (visible in the owner's screenshots). Working correctly per the owner: Ctrl+Z, Ctrl+C, Ctrl+V.

**Testing gap exposed:** 2527 unit tests and the e2e smoke suite passed while typing into a DOCX is broken — there is no e2e test that types, selects, and uses shortcuts inside the real DOCX editor. Fixing these requires Playwright scenarios on a real document, not only unit tests. Hypothesis to check first: several symptoms (typing, Ctrl+Y pasting, Ctrl+A, selection loss) point at the keyboard/`beforeinput` path between the wave-2 shortcut dispatcher and `DocxViewer`'s contentEditable handlers after the wave-3 merges.

| ID | Sev | Category | Verify | Symptom (owner's words, translated) | Expected behavior | Investigation / fix direction | Eff |
|---|---|---|---|---|---|---|---|
| USR-01 | CRITICAL | bug | OWNER | **Cannot type in the DOCX** ("on ne peut pas écrire dans le docx"). | Typing inserts text at the caret, like Word. | Trace keydown → `beforeinput` → `handleBeforeInput` → command → re-layout; check the shortcut dispatcher and track-changes recording aren't swallowing input; add a Playwright typing test on a real document. | M |
| USR-02 | HIGH | bug | OWNER | **Ctrl+Y pastes unrelated content** instead of redoing. | Ctrl+Y (and Ctrl+Shift+Z) = redo; never inserts content. | Check Ctrl+Y binding in the editor/dispatcher and whether it reaches a paste/clipboard or stale composite-command path. | S |
| USR-03 | HIGH | bug | OWNER | **Underline cannot be removed** once applied, not even with Ctrl+Z. | Ctrl+U/button toggles underline off; undo restores the previous state. | Toggle logic for `u` run property (value `none` vs absent) and the inverse command recorded in History. | S |
| USR-04 | HIGH | bug | OWNER | **Aligning the selected first line also re-aligns the following paragraph.** | Alignment applies only to paragraphs intersecting the selection. | Selection → paragraph-range resolution (selection end at offset 0 of the next paragraph must not include it); DOM↔model position mapping. | S |
| USR-05 | HIGH | bug / ux | OWNER + SCREEN | **Caret placement is inconsistent**: clicking in the empty space after a line doesn't put the caret at the end of that line. **Selection shows odd small boxes at spaces** (highlight broken into segments per word/space). | Click to the right of a line → caret at end of that line; click below the last line → end of paragraph; selection is one continuous highlight. | Hit-testing in `Cursor.ts` (`domPointToPosition`) for points outside glyph boxes; per-item absolutely-positioned spans leave gaps/overlaps — render selection from line boxes or make inter-word spaces part of the text spans. | M |
| USR-06 | HIGH | bug | OWNER | **Ctrl+F counts matches but doesn't go to them.** | Enter/next/previous scroll to the match, select it, and show the current index. | Find & Replace result → `positionToDomRange` + scrollIntoView on the virtualized/paginated page; select the match in the model. | S |
| USR-07 | HIGH | bug | OWNER | **Ctrl+A does not select all.** | Ctrl+A selects the whole document body. | Editor select-all command registered with the dispatcher; model selection spanning first→last paragraph. | S |
| USR-08 | HIGH | bug | OWNER | **Highlight (surlignage) does not work.** | Highlight button applies/removes `w:highlight` (or shading) with the chosen color. | Toolbar highlight command → run props → render + serializer. | S |
| USR-09 | MEDIUM | ux | OWNER | **After Ctrl+B / Ctrl+I the selection disappears**, so the user can't immediately toggle back. | Formatting commands keep the same selection highlighted. | Restore the model selection to the DOM after the re-layout that follows a formatting command. | S |
| USR-10 | MEDIUM | bug | OWNER + SCREEN | **After some edits, the first sentence shows no font and no size in the toolbar.** | Toolbar shows the effective font/size from the style cascade (e.g. Times New Roman, 12). | Toolbar state reads direct run props only, or loses them after run splitting; derive from resolved (cascaded) props. | S |
| USR-11 | MEDIUM | feature-gap | OWNER | **Far too few fonts** in the font picker. | Word-like list: bundled substitutes + all fonts installed on the system (+ document/theme fonts), searchable, rendered in their own face. | Enumerate system fonts (Electron `queryLocalFonts`/main-process font listing), merge with theme/document fonts, searchable combobox; keep metric-compatible substitutes for layout. | M |
| USR-12 | MEDIUM | ux | OWNER + SCREEN | **Toolbar and buttons look bad**: Save / Save As / Print / Update Fields / Update TOC are oversized, wrap onto two lines and overflow off-screen ("Update TOC" cut off); "Size" dropdown truncated; the page area gets a horizontal scrollbar. | Compact, consistent Word-like ribbon: grouped icon buttons with tooltips, overflow menu, no wrapping or horizontal scroll at common window sizes. | Redesign the DOCX toolbar/topbar layout (ribbon groups, icon+label sizes, responsive overflow), move document actions into a menu. | M |
| USR-13 | MEDIUM | bug | SCREEN | **Page count disagrees**: the editor shows "Pages: 2" while the status bar shows "1 pages" (also "1 pages" grammar). | One source of truth for page count; correct pluralization. | Stats published to the status bar vs paginator result. | S |
| USR-14 | LOW | bug | SCREEN | **"No outline available"** in the sidebar for this DOCX (to verify: may be correct if the document has no heading styles). | Outline lists heading-styled paragraphs (and optionally Title). | Check nav items derivation uses resolved style outline levels. | S |

### 14b. Owner-reported issues in other formats (reported 2026-09-15, same build)

| ID | Sev | Category | Verify | Symptom (owner's words, translated) | Expected behavior | Investigation / fix direction | Eff |
|---|---|---|---|---|---|---|---|
| USR-15 | HIGH | bug | OWNER + SCREEN | **PowerPoint: text is rendered badly** — on `sample.pptx` the title "Atlas PPTX fixture" is clipped (tops/bottoms of glyphs cut off); thumbnail text is unreadably tiny. | Text boxes show full glyphs with correct line height, font and size; thumbnails are a faithful scaled-down slide. | Text-box height / line-height / `overflow:hidden` in the slide renderer and EMU→px font scaling; check with real-world decks, not only the fixture. | S |
| USR-16 | HIGH | feature-gap | OWNER + SCREEN | **PowerPoint "has literally nothing"**: only zoom and fullscreen; no editing or real presentation features. | A usable presentation app: edit text in place, add/delete/duplicate/reorder slides, move/resize shapes and images, insert text box/image/shape, basic formatting, speaker notes editing, slideshow with presenter view, save back to .pptx (and .odp). | New slide editing layer on top of the SlideData model + a PPTX/ODP writer that preserves untouched parts (same passthrough approach as the DOCX serializer). | XL |
| USR-17 | CRITICAL | bug | OWNER | **Excel: nothing can be edited, including inside tables**, although spreadsheet editing shipped in 3.1.0 (wave 3 `sheets`). | Double-click/Enter/typing edits a cell; formulas; insert/delete rows/columns; editing works inside Excel tables (ListObjects) and merged areas; save. | Reproduce in the packaged app on real .xlsx files (with and without Excel tables): check the editing mode is actually enabled/reachable (grid `onCellEdited`, readonly flags left from the old P1.12 "no fake editing" fix, keyboard dispatcher swallowing keys), then add an e2e editing scenario. Parse/preserve Excel table definitions (`xl/tables/*.xml`). | M |
| USR-18 | HIGH | feature-gap | OWNER | **Code files (e.g. .js) are "not done at all"**: read-only highlighting, no features. | A real code editor: editing with undo/redo, save, line numbers, find/replace, go to line, code folding, bracket matching, auto-indent, multi-cursor, word wrap toggle, minimap optional, language detection. | Replace the Shiki read-only viewer with a lazily-loaded CodeMirror 6 (or Monaco) editor wired to the ViewerContext session contract (dirty/save) and the shortcut dispatcher; keep Shiki themes or map themes. Needs new npm dependencies (worktree with its own install). | L |
| USR-19 | MEDIUM | feature-gap | OWNER | **No code execution** ("exécution etc."). | Explicit "Run" for supported languages (JS/TS via Node, Python if installed…) with an output panel, stop button, and clear warnings. | Security-sensitive: running code from an opened file must be an explicit user action, never automatic; run in a separate child process from the main process (no renderer access, no Node integration in the renderer), with timeout/kill, working-directory choice, and a first-run consent dialog. Design review before implementation. | L |


### 14c. Wave-4 resolution status (2026-09-17)

Fixed in wave 4 (branch `wave4/code-editor`, stacked on `wave4/docx-editor-ux` → `wave4/sheets-editing-fix` → `wave4/slides-editor`), each with real-app Playwright coverage, because the 2 527-test unit suite had passed while typing into a DOCX was completely broken:

| ID | Status | Where | Evidence |
|---|---|---|---|
| USR-01…USR-09, USR-13 | Fixed | `4c39700` | Native `beforeinput` listener (React's synthetic `onBeforeInput` is built on legacy `textInput` and carries no `inputType`), pointer hit-testing in `docx/editor/Cursor.ts`, toolbar toggles, CSS Custom Highlight for Find, page-count stats — `tests/e2e/docx-editor.spec.ts` (7 scenarios) |
| USR-10, USR-11, USR-12 | Fixed | `ef1965f` | Searchable `FontPicker` fed by installed Windows fonts (`electron/lib/systemFonts.cjs` + `fonts:list` IPC) and a compact document toolbar |
| USR-14 | Verified, no defect | — | The reported document has no heading styles, so an empty outline is correct |
| USR-17 | Fixed | `49924b6` | Root cause: `index.html` had no `#portal` element, which glide-data-grid requires to mount ANY cell editor — every spreadsheet/CSV cell was uneditable in 3.1.0. Also fixes edits dropped when Enter follows typing within one frame, adds an Excel-like blank margin past the data, and reads/keeps Excel tables (`viewers/spreadsheet/spreadsheetTables.ts`) — `tests/e2e/spreadsheet-editor.spec.ts` |
| USR-15 | Fixed | `34f2b52` | `a:bodyPr` insets/anchor/wrap resolved through the layout/master chain, PowerPoint's ~1.2 line spacing, and text no longer clipped by a tight box; wider thumbnail rail |
| USR-16 | Fixed | `39b4622` | XML-passthrough PPTX editing (`viewers/slides/pptx/editing/`): edit text in place, move/resize/delete shapes, insert text boxes, add/duplicate/delete/reorder slides, speaker notes, presenter view, Save/Save As — `tests/e2e/pptx-editor.spec.ts` |
| USR-18, USR-19 | Fixed | `9a89f36` | CodeMirror 6 editor (find/replace, go to line, folding, multi-cursor, wrap, save) and an explicit Run gated by a native confirmation tied to the file's content (size+mtime, re-confirmed whenever it changes), running the file on disk in a shell-less child process with a 60 s limit, output cap and Stop |

Still open after wave 4:

- **ODP editing** — `.odp` decks stay read-only; the editing layer is PPTX-only (the ODF writer is a separate piece of work).
- ~~**Spreadsheet styling on save**~~ — fixed after the wave-4 review: an .xlsx/.xlsm is now saved THROUGH the file it was opened from (`viewers/spreadsheet/xlsxPassthrough.ts`), so untouched cells keep their exact XML (type, style, number format, cached formula value) and styles/charts/filters/pivots pass through; the stale calc chain is dropped and the workbook is marked "recalculate on load". The fresh-workbook writer remains the fallback for CSV, format changes, and sheet add/delete, and references outside the model (conditional formatting, data validation, defined names) are not re-anchored when rows/columns move.
- ~~**Rotated shapes**~~ — fixed: hit testing, the selection box and handle resizing all work in the shape's own rotated space (`viewers/shared/slideGeometry.ts`).
- **No Office verification** — the grafted OOXML was validated by structure and round-trip through SheetJS/the parsers; nobody opened the output in Microsoft Office, which is the only real proof of "no repair prompt".

---

## Refuted During Verification

**None.** The original audit's verification pass produced zero refuted findings across all 294 items (`"refuted": []` in the source audit data), and this finalization's critic-review pass (below) likewise refuted none of the critic's factual/sequencing/coverage claims — every one was independently confirmed against the live codebase or the plan's own text. 25 findings had their severity adjusted during the original verification pass (21 downgraded, 4 upgraded) — each is marked inline above with `⇩was X` / `⇧was X`; the verifier's rationale for each adjustment is embedded in that row's Impact/Recommendation summary and in the full audit record.

---

## Corrections and Additions from Critic Review

This register was finalized after an independent critic review of the improvement plan built from it. Every critic claim was checked against the live codebase before being acted on (see Section 11 of the companion `atlas-phase3-improvement.md` for the full narrative). Two things happened here: two existing rows were corrected in place, and five new findings were added. Nothing was refuted or removed.

**Corrections applied (2):**

1. **SHELL-13** — the Impact column named the export screenshot library as `html2canvas`. Verified via `grep` that `src/utils/export.ts:16` actually imports from `html2canvas-pro` (a different, forked npm package); the bare `html2canvas` package in `package.json` has zero import sites anywhere in `src/`. Corrected in place; `Verify` remains `CONF` since the underlying finding (viewport-only screenshot export) is unchanged, only the library name was wrong.
2. **SHELL-19** — claimed "six independent `window.addEventListener('keydown')` sites" but named only four (`useUniversalShortcuts`, `useFontSize`, `useSearch`, `ThemeMenu`), implicitly counting DocxViewer's handler as a fifth "window-level" site to approximate six. Verified via `grep` that the actual six `window`-level sites are those four plus `ExportMenu.tsx:36` and `ShortcutsModal.tsx:17` (both previously unlisted), and that DocxViewer's combo-key handling (`DocxViewer.tsx:1321`) is in fact `onKeyDown={handleKeyDownEvent}` — a React prop on its contentEditable div, not a `window` listener. Corrected in place; `Verify` changed from `UNV` to `PART` to reflect that the core architectural claim held but one framing detail (the count and its composition) needed adjustment.

**New findings added (5):**

| ID | Sev | Reason added |
|---|---|---|
| ELEC-24 | CRITICAL | Confirmed via direct code read (`main.cjs:264-289`) that the `save-file` handler — markdown's own save path — shares the identical non-atomic `fs.writeFileSync` pattern already flagged as CRITICAL (DXS-12) for `save-binary-file`, but was never flagged for this handler in the original 294. |
| DXP-20 | MEDIUM | Confirmed via grep that all 11 DOCX parser modules use `fast-xml-parser` with no depth/size guard beyond D21's zip-level check. Scope corrected from the critic's broader "DOCX/PPTX/ODP" framing to DOCX-only, since PPTX/ODP were confirmed (via code read) to use the browser's native `DOMParser` instead. |
| DXP-21 | MEDIUM | Confirmed via code read that `canvasMetrics.ts` has no `devicePixelRatio` handling, unlike the already-planned PDF-18 fix for the identical gap in the PDF viewer. |
| ELEC-25 | MEDIUM | Confirmed via file search that no `longPathAware` manifest opt-in exists under `build/`, and no UNC/long-path handling is designed into the planned IPC allowlist or atomic-write pattern. |
| ELEC-26 | MEDIUM | Confirmed by inspection that no code anywhere handles concurrent access to a `.docx` file by Microsoft Word (lock collisions); a best-effort messaging fix was scoped rather than proactive lock detection. |

**Deliberately not added as register rows** (addressed instead as plan-level Risk Register rows and `DEFER-` placeholders in `atlas-phase3-improvement.md`, since they are strategic/process gaps rather than discrete, located code defects this register's methodology captures): internationalization/locale-aware formatting, the per-machine/UAC installer tradeoff, real pixel-diff visual regression tooling, and the absence of a numeric performance budget anywhere in the plan.

**Totals after this pass:** 39 CRITICAL (was 38), 92 HIGH (unchanged), 126 MEDIUM (was 122), 42 LOW (unchanged) = **299 total** (was 294).

---

## Confirmation

- **B) Plan document**: written successfully to `C:\Users\Youssef\AppData\Local\Temp\claude\C--Users-Youssef-Documents-Projects-md-reader\e0db2117-46d2-4378-9b5e-544fb0ed11f3\scratchpad\atlas-phase3-improvement.md`, incorporating this register's corrections and additions (see its Section 11).
- **A) Findings register**: this document, written to `C:\Users\Youssef\AppData\Local\Temp\claude\C--Users-Youssef-Documents-Projects-md-reader\e0db2117-46d2-4378-9b5e-544fb0ed11f3\scratchpad\findings-register.md` — previously absent from disk (confirmed by direct file-system check before this pass began), now created as part of this finalization.
- **CRITICAL/HIGH findings left unmapped in the traceability matrix: none.** All 39 CRITICAL and all 92 HIGH findings (131 total, including the newly-added ELEC-24) are mapped to a closing task in the roadmap, including the findings explicitly routed to named `DEFER-` placeholders in Section 8 of the plan rather than silently dropped, per the plan's own traceability requirement.
