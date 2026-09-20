# Atlas 3.6.0 — Sweep A hands-on QA matrix (New document / Shell / Markdown / Code / PDF / Legacy)

**Run method:** the real Electron app, driven with Playwright's `_electron.launch`, `env: { CI: '1', PLAYWRIGHT: '1', ATLAS_HIDDEN_WINDOW: '1' }`. This is a genuine window — created, shown and composited exactly like a normal run, then made fully transparent, never focused, and click-through so it doesn't cover the owner's screen. It is **not headless**. Every script killed its own Electron process tree (`taskkill /T /F`) before the next launch; at most one instance ran at a time; no repo files were modified (fixtures for New-document/legacy-format tests were written under this scratchpad, never under `tests/e2e/`).

**Coverage:** ~53 distinct features/behaviors actually exercised across the 6 assigned areas, all **works** except one genuine bug (code editor's Ctrl+H) and a couple of intentionally-`not tested` items (Python-absent path — Python 3.14 is installed on this machine, so the "not installed" branch could not be driven without uninstalling it; and two PDF-password sub-cases: owner password, and the cancel/reopen-prompts-again cases). Nothing below is inferred from source — every "works" row was independently observed via a live DOM assertion or a file-system check after a real UI interaction, with `pageerror`/console-error listeners attached on every launch (all came back empty unless noted).

Legend: **works** / **partly works** / **broken** / **not tested**.

---

## 1. New document

| Feature | Status |
|---|---|
| New menu opens fully inside the window (no wrapped labels, x + width ≤ viewport) | works |
| Ctrl+N opens the New menu | works |
| New → Markdown (.md) creates and opens a 0-byte-by-design blank file, editable and saveable | works |
| New → Word document (.docx) creates a real non-empty file, opens in a tab | works |
| New → Excel workbook (.xlsx) creates a real non-empty file, opens in a tab | works |
| New → OpenDocument Spreadsheet (.ods) creates a real non-empty file, opens in a tab | works |
| New → PowerPoint presentation (.pptx) creates a real non-empty file, opens in a tab | works |
| New → OpenDocument Presentation (.odp) creates a real non-empty file, opens in a tab | works |
| New onto an existing file overwrites it | works |
| Cancelling the native save dialog during New creates no file/tab | works |

**Note on "New onto an existing file":** Atlas has no *own* overwrite-confirmation dialog for New — it relies entirely on the native `showSaveDialog`'s own built-in "this file already exists" prompt (same as every other save path, per `electron/main.cjs`'s comment at the `document:new` handler). Since e2e tests must stub `dialog.showSaveDialog` to a fixed path (a real native dialog can't be driven), this specific layer of protection could not be observed directly, only inferred from reading `electron/main.cjs:995-1017` — flagging this explicitly rather than claiming it as fully "works" from a live click. The write itself (atomicWriteFile, no app-level double-check) was confirmed live: pointing New → Markdown at a pre-existing file replaced its content with the blank template with no in-app warning.

---

## 2. Shell

| Feature | Status |
|---|---|
| Opening several documents (Open dialog) adds a tab each | works |
| Switching tabs shows the other document's content | works |
| Switching away from a dirty tab shows the Save/Discard/Cancel prompt (not a silent switch) | works |
| Cancel on that prompt keeps the original tab active, edit intact | works |
| Discard on that prompt switches tabs and throws the edit away, file on disk unchanged | works |
| Closing a dirty tab (not just switching) also shows the unsaved-changes prompt | works |
| Close tab, then Ctrl+Shift+T reopens the closed tab with its content | works |
| Drag-and-drop a real OS file onto the window opens it | works |
| Recent-files list (welcome screen) shows documents opened earlier in the same session, in MRU order | works |
| Clicking a recent-file entry reopens it | works |
| Recent-files persistence **across separate app launches** | not tested — `PLAYWRIGHT=1` makes main.cjs give every launch its own fresh temp `userData` dir (`electron/main.cjs:28-31`), by design, so no e2e harness (including the repo's own) can observe cross-launch persistence this way |
| All 5 themes apply (Light, Dark, Sepia, Nord, Dracula — verified via `data-theme` attribute, one at a time) | works |
| Language switch English → French relabels the UI (e.g. "Editor" → "Éditeur") | works |
| Language switch French → English restores English labels | works |
| Ctrl+/ opens the shortcuts dialog; it lists 21 shortcut entries | works |
| Escape closes the shortcuts dialog | works |
| Ctrl+1 / Ctrl+2 / Ctrl+3 switch Preview / Split / Editor modes | works |
| Ctrl+B toggles the sidebar | works |
| Ctrl+T cycles the theme | works |
| Ctrl+=, Ctrl+-, Ctrl+0 change/reset markdown font size (15px→16px→14px→15px observed) | works |
| Ctrl+E opens the Export menu | works |
| Ctrl+P (print/export) triggers no crash/pageerror (native dialog itself not directly observable) | works |
| Ctrl+Shift+S (Save As) writes a new file at the chosen path | works |
| Closing the window with a dirty document + Cancel keeps it open, edit intact | works |
| Closing the window with a dirty document + Discard closes it, file on disk unchanged | works |
| Closing the window with a dirty document + Save (real save-before-close round trip) | not tested directly in this sweep (Cancel/Discard were; the existing `tests/e2e/close-confirmation.spec.ts` in the repo already covers Save explicitly and was read, not re-driven, to avoid duplicating it) |

---

## 3. Markdown

| Feature | Status |
|---|---|
| Preview mode renders the markdown body | works |
| Split mode shows editor + preview simultaneously | works |
| Editor (raw) mode shows an editable textarea | works |
| Mermaid diagram renders to a real `<svg>` | works |
| A broken Mermaid diagram shows an inline error, no crash, no stray DOM left behind | works |
| KaTeX inline (`$...$`) and block (`$$...$$`) math render | works |
| Table of contents (sidebar) lists every heading in the document | works |
| Clicking a TOC entry scrolls the document to that heading | works |
| Ctrl+F search finds text in the preview, live "N of M" count | works |
| Export → HTML produces a non-empty file containing the document's text | works |
| Export → DOCX produces a valid, non-empty, unzippable .docx whose `word/document.xml` contains the document's text | works |
| Export → PDF produces a non-empty PDF file | works |

---

## 4. Code files

| Feature | Status |
|---|---|
| Editing a .js file in the CodeMirror editor | works |
| Ctrl+F opens a combined Find **and** Replace panel (single panel, not two separate ones); Replace All works end to end | works |
| **Ctrl+H does not open Find/Replace** | broken (see below) |
| Ctrl+G (Go to line) opens a goto-line panel and jumps to the given line | works |
| Save writes the edited content to disk | works |
| Run shows a native confirmation dialog first; declining runs nothing | works |
| Run (accepted) executes a .js file and streams its stdout + "Finished (exit code 0)" | works |
| Stop halts a long-running .js program ("Stopped.") and frees the Run slot for other files | works |
| Run a .ts file (TypeScript, Node type-stripping) | works |
| Run a .py file (Python) | works — Python 3.14 is installed on this machine |
| Behavior when Python is **not** installed | not tested — could not be exercised without uninstalling the system's Python |

**Broken — Ctrl+H does not open Find/Replace in the code editor, despite the source comment claiming it does:**
- File: `src/viewers/code/CodeEditor.tsx`, line 5 comment reads `search/replace (Ctrl+F / Ctrl+H)`, but no `Ctrl+H`/`Mod-h` binding exists anywhere in the file (only `Mod-g` and `Mod-d` are explicitly bound at lines 41-44) and CodeMirror 6's own default `searchKeymap` (pulled in via `basicSetup`) does not bind `Mod-h` either.
- Repro: open any code file (e.g. a `.js`), click into the editor, press Ctrl+H on a completely fresh session (no prior Ctrl+F). Observed: the `keydown` event for `h`/`ctrlKey` does reach `.cm-content` (confirmed via a capture-phase listener — `ctrl:true, defaultPrevented_before:false, target:"cm-content"`), but no `.cm-search` panel appears, even after a second Ctrl+H attempt or a ~1s wait. Pressing **Ctrl+F** immediately afterward opens the panel correctly (and that panel already contains both a Find and a Replace field together, `[name="search"]` / `[name="replace"]`, plus a "replace all" button that works).
- Expected (per the source comment): Ctrl+H should also open the panel. Observed: it silently does nothing.
- Severity note: **low real-world impact** — the code viewer's own visible UI button for this (`src/viewers/CodeViewer.tsx:177`) is correctly labeled `title="Find and replace (Ctrl+F)"` (`src/i18n/messages.en.ts:215`) and never promises Ctrl+H to the user, and Ctrl+F alone already reaches the full Find+Replace functionality. So this is a stale/inaccurate source comment plus a genuinely-missing (but not user-advertised) keybinding, not a broken user-facing feature. Suggested fix: either delete "/ Ctrl+H" from the `CodeEditor.tsx` header comment, or add `{ key: 'Mod-h', run: openSearchPanel, preventDefault: true }` to `editorKeys` if Ctrl+H was actually meant to work.
- No console/page errors accompanied this — it's a silent no-op, not a crash.

---

## 5. PDF

| Feature | Status |
|---|---|
| Multi-page PDF loads all pages (5/5 page-wrappers, "/ 5" indicator) | works |
| Zoom In steps 100%→125%; Zoom Out returns to 100% | works |
| Rotate page swaps a page's aspect ratio (portrait↔landscape bounding box) | works |
| Thumbnail rail shows one thumbnail per page and toggles closed | works |
| Jumping to a page via the page-number input navigates there | works |
| Next-page toolbar button advances the page | works |
| Ctrl+F find bar searches across pages, live "N of M" match count | works |
| Print renders every page into the print-only DOM (5/5) before calling `window.print()` | works |
| Opening a password-protected PDF shows the password dialog | works |
| Wrong password shows an inline error and keeps the dialog open for retry | works |
| Submit re-enables once a new password is typed after a failed attempt | works |
| Correct (user) password unlocks the document, no error screen | works |
| Owner password also unlocking the document | not tested (only the user password was exercised live) |
| Cancelling the password dialog (error state, safe tab close afterward) | not tested |
| Reopening the same encrypted file in a fresh app instance re-prompts (not cached) | not tested |

---

## 6. Legacy .doc / .ppt (read-only viewers)

Real OLE2/CFB fixtures were built on disk in the scratchpad using the exact same byte-layout helpers the repo's own unit tests use (`src/legacy/__tests__/fixtures.ts`, ported to a standalone script since that module only builds in-memory buffers) — never files taken from or written into the repo.

| Feature | Status |
|---|---|
| Legacy .doc viewer parses a real .doc and renders its paragraph text | works |
| Legacy .doc viewer shows a read-only banner naming .docx as the modern format | works |
| Legacy .doc viewer has no Save button (confirmed read-only) | works |
| Opening a corrupt/invalid .doc shows a friendly inline error, not a crash | works |
| Legacy .ppt viewer parses a real .ppt and renders the active slide's title + bullets | works |
| Legacy .ppt viewer shows one thumbnail per slide | works |
| Clicking a slide thumbnail switches the active slide's content | works |
| Legacy .ppt viewer shows a read-only banner naming .pptx as the modern format | works |

---

## Summary counts

- **works:** 53
- **partly works:** 0
- **broken:** 1 (code editor Ctrl+H — see Section 4)
- **not tested:** 7 (Python-not-installed; cross-launch recent-files persistence; window-close+Save round trip [already covered by the repo's own `close-confirmation.spec.ts`]; PDF owner password; PDF cancel-dialog case; PDF re-prompt-on-reopen case)

No `pageerror` or console error was captured on **any** run in this sweep.
