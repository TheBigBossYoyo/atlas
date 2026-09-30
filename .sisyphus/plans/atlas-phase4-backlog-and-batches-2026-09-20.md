# Atlas Phase 4 — verified backlog and batch execution plan

**Date:** 2026-09-20 · **Baseline:** `main` at 3.6.0 + two fixes committed today
(`06ba9bc` New-menu geometry, `87355fb` invisible-window e2e mode).

This file is the deliverable of a three-part investigation: a hands-on sweep of the
real app (Part 1), an analysis pass (Part 2), and the parallel execution plan the next
sessions run (Part 3).

## How to read this

Every item below was **re-verified against the code by the session lead**, not accepted
from an agent's report. Where an agent's claim did not survive that check, it is not
here. Items that could not be proven by reading are quarantined in §4 and are explicitly
**not** scheduled until a run confirms them.

Severity means user impact, not code ugliness:
- **critical** — silent data loss, or a security hole reachable without user error.
- **high** — content or formatting is lost/corrupted, or a security control is defeated.
- **medium** — visible wrongness, or a real accessibility/UX failure.
- **low** — cosmetic, or a gap with a plausible workaround.

---

## 1. Fixed today (already on `main`)

| ID | What | Commit |
|---|---|---|
| FIX-1 | The **New** menu opened 76px off the left edge of the window, and 4 of its 6 labels wrapped to two lines. `.dropdown__menu` is right-aligned, correct for the three dropdowns in `.toolbar__right`, wrong for the one at the far left. Left-aligned it with a `--start` modifier; `white-space: nowrap` lets every dropdown size to its content. Measured before: `x=-75.96`. After: all four dropdowns fully inside the window, no wrapped labels. | `06ba9bc` |
| FIX-2 | `ATLAS_HIDDEN_WINDOW=1`, a test-only invisible-window mode for driving the real app. **Not headless** — `show:false` stops Chromium compositing and `requestAnimationFrame`, so Playwright's actionability checks hang and every click times out (measured: 30s timeout on the first click). Instead the window is shown, `setOpacity(0)`, `showInactive()`, click-through. Full 79-spec suite passes with the flag off. | `87355fb` |

Both bugs were invisible to the existing suites for the same reason the overnight run's
bugs were: Playwright still clicks an element whose visible remainder is on-screen, and
jsdom has no layout at all. FIX-1's regression test asserts the rendered box and fails on
the old CSS with exactly the measured `-75.96`.

---

## 2. Verified backlog

### 2.1 Critical and high — data loss

**SAVE-1 · high · effort M** — Save As on a non-markdown document, resolving after you
switch tabs, hijacks the tab you switched *to*.
`handleViewerSavedPath` (`src/App.tsx:561`) guards only `savedPath === file.path`, then
unconditionally `renameSession(...)` + `showSessionFile(moved)` → `adoptFile`.
`savedPathRef` is kept pointed at the *current* tab by an effect
(`src/viewers/shared/ViewerContext.tsx:121-126`), so a save that resolves late calls the
new tab's callback with the old tab's path. The new tab is renamed to that path and
re-read from disk over whatever is showing.
**This exact bug was already fixed on the markdown path and never ported.** Compare
`applySavedPathToInactiveTab` (`src/App.tsx:392`), whose own comment says it must
"never touch the tab that's showing now: no `adoptFile` (which would yank the view back
to the just-saved document)", and the `stillShowing` guards at `src/App.tsx:427` and
`:473`. None of the three viewer call sites has an equivalent check
(`DocxViewer.tsx:2097`, `useSlideEditorCore.ts:159`, `useSpreadsheetEditor.ts:263`).
*Files:* `src/App.tsx`, `src/viewers/shared/ViewerContext.tsx`, `src/viewers/DocxViewer.tsx`,
`src/viewers/slides/shared/useSlideEditorCore.ts`, `src/viewers/spreadsheet/useSpreadsheetEditor.ts`.
*Test gap:* `ViewerContext.test.tsx:416` only asserts the callback fires; no test changes
`filePath` between save start and resolution.

**DOCX-1 · high · effort S** — Theme-referenced fonts are dropped on every save.
Parser reads `w:asciiTheme`/`hAnsiTheme`/`cstheme`/`eastAsiaTheme`
(`src/docx/parser/document.ts:1826-1839`, `src/docx/parser/styles.ts:604-607`) and the
model carries them (`src/docx/model/styles.ts:167-177`). Both writers emit only the
literal attributes and no theme attribute at all:
`buildFontSetXml` (`src/docx/serializer/stylesWriter.ts:508-520`) and
`buildFontSetElement` (`src/docx/serializer/documentWriter.ts:1302-1315`, reused by the
header/footer/footnote/endnote/comment writers).
Word has specified Normal's font by theme reference since Office 2007, so this hits most
real documents; where no literal `w:ascii` accompanies it, the whole `w:rFonts` element
disappears and the font silently reverts.

**SHEET-2 · high · effort S** — Excel table filters and sorts are wiped on every save,
even when nothing about the table changed.
`rewriteTableXml` (`src/viewers/spreadsheet/spreadsheetTables.ts:256-258`) unconditionally
runs `removeAll(root, 'sortState')` and `removeAll(root, 'filterColumn')`, and
`xlsxPassthrough.ts:1148-1151` calls it for every table on every sheet on every save with
no "did this table change" guard. Edit one unrelated cell, save: the AutoFilter resets to
show-everything, dropdown arrows still present, no warning. The code comment ("Column
indexes may have moved, so filter/sort state cannot be trusted") is the right instinct
applied far too broadly.

**SHEET-3 · high · effort S-M** — A whole-column or whole-row conditional format, data
validation, hyperlink or autoFilter is **deleted** by any row/column insert or delete
anywhere on that sheet.
`decodeCell` (`src/viewers/spreadsheet/spreadsheetTables.ts:53`) uses
`/^\$?([A-Z]+)\$?(\d+)$/i`, which cannot match a bare `A` or `1`, so `decodeRange("A:A")`
returns `null`. `remapSqref` (`src/viewers/spreadsheet/spreadsheetRangeShift.ts:110-120`)
treats that as "malformed piece in the source file — drop rather than propagate garbage",
returns `null` when every piece drops, and `updateShiftedRanges` then removes the owning
element outright. But `A:A` is perfectly valid OOXML. The sibling `formulaRefs.ts` already
handles whole-row/column ranges and has a regression test for it; the capability never
reached the sqref path, and no test covers a whole-column `conditionalFormatting`.

**SHEET-4 · high · effort S-M** — An xlsx save silently falls back from the lossless
passthrough writer to the lossy fresh-SheetJS writer, and still reports success.
`writeWorkbookThroughOriginal` returns `null` on several internal failures; the caller
does `throughOriginal ?? writeWorkbookBytesWithTables(...)`
(`src/viewers/spreadsheet/useSpreadsheetEditor.ts:89-93`). `docs/KNOWN_LIMITATIONS.md`
promises styles/charts/filters/tables survive an xlsx save. When the fallback fires they
do not, and nothing tells the user — there is no spreadsheet equivalent of the DOCX
`lossySaveWarnings` detector.
*Trigger rate in the wild is unmeasured; the mechanism is verified.*

**DOCX-2 · high · effort S (warn) / M (preserve)** — Content controls and shape fallbacks
inside **footnotes, endnotes and comments** are stripped on save with no warning.
`parseBlocksFromXmlFragment` (`src/docx/parser/partBody.ts:60-65`) reads
`document.sections` and discards `document.wrappers`, so the passthrough regions are gone
before any writer sees them; no notes/comments writer references `wrapperRegions`. And the
detector that would at least *tell* the user explicitly excludes those three parts
(`src/docx/fidelity/lossySaveWarnings.ts:56-67`). `docs/KNOWN_LIMITATIONS.md` documents
this stripping for table cells/headers/footers only.

**DOCX-3 · high · effort S-M** — Paragraph-mark run properties (`w:pPr/w:rPr`) are never
parsed, so they are stripped from every paragraph.
`parseParaProps` (`src/docx/parser/document.ts:1591-1644`) enumerates `w:pStyle` through
`w:sectPr` and never `w:rPr`; `ParaProps` has no field; the writer says so itself
(`src/docx/serializer/documentWriter.ts:1162`: "Atlas doesn't model a paragraph-mark
`rPr`"). This is frequently the only thing setting a blank line's font size, so spacing
shifts through the document after a save.

### 2.2 Critical and high — security

**SEC-1 · high · effort S** — The path allowlist can be bypassed for any file that exists
on disk, giving a compromised renderer arbitrary file read and **silent** arbitrary file
overwrite.
`preload.cjs:103` exposes `registerDroppedPath` to the renderer. `registerDroppedPath`
(`electron/main.cjs:363-369`) checks main-frame, non-empty string, and `fs.existsSync` —
then `pathAllowlist.add(filePath)`. It never verifies the path came from a real
`webUtils.getPathForFile` drop. That same allowlist is the only gate on
`file:readBinaryByPath` (`electron/main.cjs:844-861`) and on `save-binary-file`'s
`existingPath` (`:941-944`), where a hit **skips the save dialog** and goes straight to
`atomicWriteFile`. `isFromMainFrame` is no defence: under this codebase's own threat model
the hostile script runs in the main frame.
The team already identified this shape — `electron/lib/recentFilesStore.cjs`'s header calls
drag-drop registration unverifiable and excludes it from the persisted store — but the
mitigation was applied only to recent-files persistence, not to the read/write allowlist.
*Fix direction:* split write-eligibility out of the general allowlist. Only `trustPath()`
sources (dialog results, argv/open-file, recent-store-verified) may satisfy
`existingPath` in `save-file`/`save-binary-file`.

**SHEET-1 · high · effort S** — The zip-bomb guard was never wired into the code path that
actually opens a spreadsheet.
`loadWorkbookZip` (`src/viewers/spreadsheet/spreadsheetZipBudget.ts`) rejects on declared
uncompressed size before inflating, and the passthrough/table/pane readers all use it. But
`parseWorkbookBuffer` (`src/viewers/shared/spreadsheetGrid.ts:222`) — reached from the
viewer, the Worker and the PDF export — calls `XLSX.read(new Uint8Array(buffer), ...)`
directly with no budget. `electron/lib/fileSizeGuard.cjs` caps compressed on-disk size
(200 MiB), not the decompression ratio.
*Verified: no guard runs on that path. The memory-exhaustion consequence is a structural
inference, not a measured number.*

**SEC-2 · low · effort XS** — The hidden `BrowserWindow` in
`electron/lib/printToPdf.cjs:125-134` never gets `applyNavigationGuards`. Mitigated by
`javascript: false` and the export sanitizer's `FORBID_TAGS` (which includes `meta`,
`base`, `iframe`), and it does inherit the default session's CSP hook. Defence-in-depth
only; no exploit path found.

**SEC-3 · medium · effort S** — No encryption detection for legacy `.doc`/`.ppt`.
`src/legacy/doc/fib.ts:24-29` reads `flags1` but only tests `F_WHICH_TABLE_STREAM`; the
`fEncrypted` bit (`0x0100`) is never read, and there is no
`CryptSession10Container` check in `src/legacy/ppt/slides.ts`. An encrypted file is parsed
as plaintext. PDF, by contrast, has a proper password flow.
*The "no check exists" half is verified; the exact failure mode (confusing error vs.
mojibake shown as content) is not — it needs an encrypted fixture.*

### 2.3 Medium — fidelity

| ID | What breaks for the user | Sev | Effort | Files |
|---|---|---|---|---|
| FID-1 | Spreadsheets render as a flat, unstyled grid: no bold, no font/fill colour, no conditional-format colour scales, data bars or icon sets, no rich-text run formatting. `SheetGrid.rows` is `string[][]` (`spreadsheetGrid.ts:53`) — there is no data path for style to travel, and every cell is `kind: 'text'` (`useSpreadsheetGrid.ts:151-166`). The style XML *is* preserved on disk by the passthrough; it is simply never read for display. | medium | L | `spreadsheetGrid.ts`, `xlsxCellFormat.ts`, `useSpreadsheetGrid.ts` |
| FID-2 | Hidden rows and hidden columns are shown fully expanded. `sheetToGrid` (`spreadsheetGrid.ts:193-207`) reads `wpx`/`hpx` but never `ColInfo.hidden`/`RowInfo.hidden`. `KNOWN_LIMITATIONS.md` documents hidden *sheets* only. | medium | S | `spreadsheetGrid.ts`, `useSpreadsheetGrid.ts`, `SpreadsheetViewer.tsx` |
| FID-3 | Cell comments and threaded comments are completely invisible — no marker, no tooltip. They are referenced only in `xlsxPassthrough.ts` for zip housekeeping on sheet delete. | medium | M | new parser + `SpreadsheetViewer.tsx` |
| FID-4 | A PDF bookmark or internal link lands at the top of the page instead of the spot it points at. `resolveDestinationPage` (`src/viewers/pdf/outline.ts:26-48`) keeps only `dest[0]` and discards `/XYZ`, `/FitH`, `/FitR` offsets and zoom. | medium | M | `pdf/outline.ts`, `pdf/annotationOverlay.ts`, `pdf/PdfViewer.tsx` |
| FID-5 | A reviewer's sticky-note comments on a PDF cannot be read. `classifyAnnotation` (`src/viewers/pdf/annotations.ts:91-116`) recognises only `Link` and `Widget`; `RawPdfAnnotation` has no `contents` field at all. The icon may paint via the appearance stream, but there is no way to open it. | medium | M | `pdf/annotations.ts`, `pdf/annotationOverlay.ts` |
| FID-6 | PDFs with interactive layers (CAD exports, multi-language overlays) cannot have layers toggled. Initial visibility is correct — pdf.js applies the document default when `optionalContentConfigPromise` is omitted — there is just no UI. | low | M | `pdf/PdfPage.tsx`, `pdf/PdfViewer.tsx`, `pdf/PdfToolbar.tsx` |
| DOCX-4 | Every image is written with `wp:docPr id="1"` (`documentWriter.ts:778`), so a document with two or more images has duplicate IDs after one save. Word tolerates it; it is an OOXML uniqueness violation the package validator does not catch. | medium | XS | `documentWriter.ts` |
| DOCX-5 | Overlapping floating images lose their stacking order on disk: `relativeHeight` is hardcoded `'0'`, as are `distT/B/L/R`, `simplePos`, `locked` (`documentWriter.ts:869-880`). `KNOWN_LIMITATIONS.md`'s "correct z-order" claim is true only of Atlas's own renderer. | medium | S | `documentWriter.ts`, `parser/document.ts` |
| DOCX-6 | A section overriding footnote/endnote numbering (front matter in roman numerals) reverts to the document default on save. `parseSectionProps` never reads `w:footnotePr`/`w:endnotePr`. | medium | M | `model/document.ts`, `parser/document.ts`, `documentWriter.ts` |
| DOCX-7 | A hyperlink placed on a picture itself (`a:hlinkClick`, e.g. a clickable logo) is stripped on first save. Zero matches for `hlinkClick` anywhere in `src/`. | medium | M | `parser/document.ts`, `documentWriter.ts` |
| DOCX-8 | Picture bullets (`w:numPicBullet`) are lost, because `numbering.xml` is regenerated from a model that has no field for them. | medium | M | `parser/numbering.ts`, `serializer/numberingWriter.ts` |
| DOCX-9 | Table caption/description (`w:tblCaption`, `w:tblDescription`) — screen-reader metadata — is dropped on every save. Notable given the app's own WCAG investment. | low | XS-S | `parser/document.ts`, `documentWriter.ts` |
| DOCX-10 | A floating/positioned table (`w:tblpPr`) loses its position on **open**, not just save. | low | M | `parser/document.ts`, `model/document.ts`, `documentWriter.ts` |
| DOCX-11 | Per-cell rotated/vertical text (`w:tcPr/w:textDirection`) is lost, although the identical attribute at section level round-trips correctly. | low | XS-S | `parser/document.ts`, `documentWriter.ts` |

### 2.4 Medium — UX and accessibility

**A11Y-1 · medium · effort XS** — `useFocusTrap` was extracted specifically to stop the
Tab-trap pattern being copy-pasted, but its two original sources were never migrated:
`src/components/ShortcutsModal.tsx:36-63` and
`src/components/UnsavedChangesDialog.tsx:55-84` still hand-roll it. They have **already
diverged** — the hook and the dialog filter out `disabled` elements, the modal does not.
Latent today (one focusable control), silently wrong the moment a disabled control is added.

**A11Y-2 · medium · effort XS** — `docs/KNOWN_LIMITATIONS.md:341-345` states that
`PresenterView` traps Tab via the shared hook. It does not — `PresenterView.tsx:39-66`
only captures and restores focus. When `requestFullscreen()` is refused (and the refusal
is swallowed at line 60), Tab walks straight out of presenter view into the document behind.

**DRAFT-1 · medium · effort S** — A draft the user explicitly discarded can come back as a
false "restore unsaved work?" offer. `clearDraft()` is called on a successful markdown save
(`src/App.tsx:432`, `:478`) and from the recovery banner's own Discard (`:659`), but never
from `handleUnsavedDialogDiscard` (`:509-514`) — the Discard used by Ctrl+W, tab switch,
Open and quit. `useAutosave` then stops touching the draft once the document is no longer
markdown, so it survives indefinitely.

**SESS-1 · low-medium · effort XS-S** — `renameSession`
(`src/session/documentSessions.ts:174-183`) has no de-duplication, unlike `openSession`
(`:56-68`) which merges into an existing entry. Save As onto a path that is already open in
another tab produces two sessions with the same `id`, breaking the identity invariant the
file's own header documents.
*Reachability depends on the OS dialog permitting that choice — not driven end-to-end.*

### 2.5 Test coverage — the gaps that would hide a regression

These matter more than the raw count (3324 unit, 79 e2e). Each is a place where a real
regression would pass the suite silently — the exact failure mode of the overnight run's bugs.

**TEST-1 · effort M** — The DOCX round-trip corpus always edits paragraph `[0]`.
`roundtrip.corpus.test.ts:43-47` hardcodes `FIRST_PARAGRAPH_START` and reuses it for all 22
fixtures, and paragraph 0 is a plain "Fixture: …" intro line in every one. So the
edit-then-save-then-reparse mode never touches a tracked-change region, a content control,
a table cell or a nested list item — precisely the constructs most likely to break under an
editor command.

**TEST-2 · effort S** — The corpus fingerprint counts five tag names (`w:p`, `w:tbl`,
`w:drawing`, `w:sectPr`, `w:tblGrid`) and inspects no attribute values
(`corpusRoundtripHelpers.ts:97-122`). A regression that flattened every `w:ilvl` to 0, or
dropped all tracked-change wrapping, leaves all five counts unchanged.

**TEST-3 · effort XS** — `tests/e2e/export.spec.ts:77-123` asserts only the PDF **page
count** for DOCX and PPTX export. The XLSX case in the same file already extracts per-page
text via pdfjs. A bug duplicating page 1 onto page 2 still yields `numPages === 2`.

**TEST-4 · effort M** — `lossySaveWarnings` — the one general safety net for "did the writer
silently drop something" — can only see a tag go from present to entirely absent. `walk()`
skips attributes (`:122-143`) and the comparison is `savedCounts.has(tag)` (`:190-193`), so
12 `w:tab` becoming 1, or a `w:sz` value changing, produces no signal. DOCX-1 above is
exactly this shape and the detector cannot see it.

**TEST-5 · effort S** — No test asserts `localStorage['atlas-draft']` after the
unsaved-changes dialog's Discard (see DRAFT-1), and none covers `renameSession` onto an
already-open path (see SESS-1).

### 2.6 Known items — reconfirmed

| Item | Status |
|---|---|
| Nothing Atlas writes has been opened in real Microsoft Office or LibreOffice | **Still true.** Neither is installed. `scripts/validate-office-file.mjs` is a structural substitute, and DOCX-4's duplicate `docPr` ids show it does not catch semantic violations. |
| Closed spreadsheet tab retains ~24 MB | **Refuted 2026-09-28.** Re-measured with CDP `Runtime.getHeapUsage` after repeated `collectGarbage`: 92-94% released, ~0.3-0.4 MB retained per open/close cycle. The original figure was most likely `performance.memory.usedJSHeapSize`, which Chromium quantizes and caches for ~20 minutes — it reads a flat number through a 19.5 MB workbook opening AND closing. Entry rewritten around the measurement; pinned by `tests/e2e/spreadsheet-memory-release.spec.ts`. |
| `.doc`/`.ppt` read-only, text only | Still true. SEC-3 adds that encrypted files are not detected. |
| No split view | Still true. |
| Edited content control / shape unwrapped on save, user is told | True for `document.xml`, styles, numbering, headers and footers. **DOCX-2 above narrows this**: in footnotes, endnotes and comments it is unwrapped and the user is *not* told. |
| Code signing needs a paid certificate | Still true; `docs/RELEASE.md` has the options, nothing implemented. |

Register maintenance: `.sisyphus/plans/atlas-phase3-findings-register.md:514`'s four
"review leftovers" (header/footer commit-on-blur, D23 line cache, markdown save-in-flight
tab switch, unbudgeted JSZip readers) are **all fixed in current code** with regression
tests. The register is stale there and should be corrected.

### 2.7 Second analysis pass — additional verified items

A second pass over fidelity/performance/UX independently re-found SAVE-1 and SHEET-2 (good
corroboration: two passes, same file:line evidence). These are the items it added that
survived re-verification by the session lead.

**DOCX-12 · critical · effort M-L** — A whole class of character effects is permanently
stripped the first time Atlas saves.
`RunProps` (`src/docx/model/styles.ts:305-331`) has no field for `w:outline`, `w:emboss`,
`w:imprint`, `w:em` (emphasis marks — routine in CJK documents), `w:effect`,
`w:eastAsianLayout`, `w:fitText`, run-level `w:bdr`, or `w:w` (manual character scaling).
`parseRunProps` reads a fixed list that excludes all of them and there is **no catch-all
passthrough for unknown run properties**, so `buildRunPropertiesNode`
(`src/docx/serializer/documentWriter.ts:1125-1158`) cannot re-emit what was never modelled.
*Lead verification: grepped `src/docx/model/styles.ts`, `parser/document.ts` and
`serializer/documentWriter.ts` for each of those element names — zero hits. (`shadow` and
`outlineLvl` do appear, but as a table-border property and a paragraph outline level, not
character effects.)*
This is the strongest argument for a general **unknown-run-property passthrough**, which
would close DOCX-12 and inoculate against the next dozen of its kind.

**DOCX-13 · high · effort S-M** — Text marked Hidden in Word is always displayed, and takes
up layout space. `w:vanish`/`w:webHidden` are parsed and preserved on save
(`parser/document.ts:1559-1560`), but `runStyleToCss` (`src/docx/render/style.ts:127-215`)
never consults them and layout measures hidden runs like any other.

**DIRTY-1 · high · effort S** — Undo a DOCX back to exactly its saved content and it stays
marked dirty forever; the close-confirmation prompt still fires.
Dirty is a **reference** comparison against the saved snapshot — the code says so itself
(`src/viewers/DocxViewer.tsx:960-965`: "report dirty whenever documentModel diverges (by
reference — every edit replaces it immutably) from the last-saved-or-loaded snapshot"). But
`History.undo`/`redo` rebuild the graph through the inverse command, producing an object
that is deep-equal and never reference-equal. `History.test.ts` asserts `toEqual`, never
`toBe`, which quietly confirms it.
*Code logic verified; the on-screen symptom is being checked by the hands-on sweep.*

**SHEET-5 · high · effort S-M** — A formula that depends on another formula cell later in
the sheet shows a stale value.
`recalculateSheet` (`src/viewers/spreadsheet/spreadsheetDocument.ts:205-235`) is a single
row-major pass with no dependency ordering, and its `lookup` reads
`(rows ?? sheet.rows)[row]?.[col]` — i.e. the value from *before* this pass reached that
cell. So `A1 = "=C5+1"` where `C5` is itself a formula reads C5's pre-recalculation text.
It self-heals only if some later unrelated edit happens to trigger a pass in a luckier order.
*Verified by reading the loop.*

**A11Y-3 · high · effort S** — Closing a tab with the keyboard drops focus to `<body>`, so
the next Tab starts from the top of the window. `TabBar.tsx:126-134`'s close button does a
plain `onClick` with no focus handling. `src/hooks/useRestoreFocusOnClose.ts` exists for
exactly this and is already wired into `ExportMenu`, `ThemeMenu`, `NewDocumentMenu` and
`LanguageMenu` — just never into `TabBar`. Closing a tab is a far more frequent action than
closing a dropdown.

**A11Y-4 · medium · effort XS** — The PDF find bar has no focus trap, so Tab leaks into the
document behind it. Same one-line pattern as A11Y-3. *(Note: a related PDF find-bar bug —
it stayed open and swallowed the next shortcut — was one of the "CI flakes" that turned out
to be a real product bug. This area deserves the attention.)*

**I18N-1 · high · effort XS — this is the English-mode thread, pulled** —
`src/components/RawEditor.tsx` never got the i18n pass at all: `<span>Markdown Source</span>`
and `placeholder="Type or paste markdown here..."` are hardcoded English, so they stay
English in the French UI. The textarea also has **no accessible name** — no `<label>`, no
`aria-label`, no `aria-labelledby`. Its only name is the placeholder, which disappears the
moment there is text, so a screen-reader user tabbing into a non-empty markdown editor hears
nothing. This is one of the most-used surfaces in the app.

**TEST-7 · high · effort S — nothing that persists across a relaunch is tested at all** —
`electron/main.cjs:28-32` mints a fresh temp `userData` directory for every launch when
`PLAYWRIGHT=1`, and no spec pins a stable one (grepped all of `tests/e2e/*.spec.ts`: zero
hits for a profile override). That was the right call — it stops tests writing into the
owner's real store and fixes a singleton-lock race — but it means **everything living in
`userData` is structurally untestable end to end**: the recent-files list, window state
(bounds, maximised), and the persisted `recentFilesStore`.
That last one is not a convenience feature: it is the ground truth behind
`recent:request-open`'s security check (`electron/main.cjs:382-389`), the control that was
added to fix a real localStorage-based allowlist bypass. It has no cross-launch test.
*Fix direction:* an env override for a caller-supplied profile directory, used by a small
number of specs that relaunch into the same profile deliberately.
*Found by the hands-on sweep, which listed recent-files persistence as untestable and gave
this as the reason.*

**TEST-6 · high · effort S — why I18N-1 shipped, and why more will** —
`src/i18n/__tests__/noHardcodedStrings.test.ts` guards an **opt-in allowlist**
(`CONVERTED_FILES`, ~35 paths). `RawEditor.tsx` appears in it zero times. The guard
therefore cannot catch a component nobody remembered to add — a new component is unguarded
**by default**, which is backwards for a lint-style test. Invert it: scan every component
and viewer, and require an explicit, commented opt-out for the surfaces genuinely out of
scope. Until then, "the i18n test passes" means considerably less than it appears to.

### 2.8 Found by driving the real app — the worst of the lot

These came from the hands-on sweep (§5) and were then verified in code by the session lead.
**They are more urgent than anything the static analysis found**, because they are wrong
every time you use the feature, not in a race or an edge case.

**DOCX-14 · critical · effort S — every colour Atlas applies is invalid OOXML.**
Set text colour from the toolbar and the saved file contains
`<w:color w:val="#ff0000"/>`. `ST_HexColor` is `auto` or six hex digits — **never** with a
leading `#`. Word's reaction to an invalid attribute value is the "unreadable content"
repair path, the same class the schema-order wave was chasing.
*Verified, and wider than the sweep reported:* `hexColor()`
(`src/docx/model/styles.ts:22`) is `value as HexColor` — a bare cast with no validation or
normalisation, so it is type-safety theatre. The `#` enters from two places: the toolbar
colour picker (`src/docx/editor/toolbarAdapter.ts:341`, an `<input type="color">` value) and
a **hardcoded** `TABLE_BORDER_ON` (`toolbarAdapter.ts:264`, `hexColor('#000000')`), so
toolbar-inserted table borders are invalid too. Both are written verbatim
(`documentWriter.ts:1144` and `:1452`). The parser stores the file's raw value
(`parseColor`, `document.ts:2808-2814`), so colours that came from the file round-trip
correctly — only newly applied ones are broken, which is why no round-trip test sees it.
*Fix:* normalise and validate inside `hexColor()` itself, so no call site can reintroduce it.

**DOCX-15 · critical · effort S — a new bullet or numbered list is invalid OOXML.**
`createListNumberingEntry` (`src/docx/editor/insertList.ts:86`) sets
``abstractNumId = `atlas-list-${numIdStr}` ``. `w:abstractNumId`'s value is
`ST_DecimalNumber` — an integer. Shared with rich paste's own numbering minting, so it
affects pasted lists too.

**DOCX-16 · critical · effort M — Insert Table produces an unusable table and silently
discards what you type into it.** From the toolbar, the inserted table renders at zero width
and cannot be clicked; text typed at that point is dropped entirely rather than going
anywhere. *This is live, in-session work loss, not a save-path issue.*
*Suspected:* `buildEmptyTable` in `src/docx/editor/commands.ts` / `src/docx/layout/layoutTable.ts`.
Runtime-observed; the precise cause is for the fixing agent to isolate.

**DOCX-17 · high · effort M — the caret cannot be placed in any table cell**, including a
correctly rendered pre-existing table. Clicking a cell, or Tab/ArrowDown into one, lands the
caret in the paragraph *after* the table. Cell merge is unreachable as a direct consequence.
Tables are effectively read-only in the editor.

**SHEET-6 · high · effort M** — Text-returning formulas are never evaluated. `=A1&"!"` or
`=CONCATENATE(...)` are displayed *and saved* as literal formula text. Numeric formulas
(`=A2+B2`) evaluate correctly.

**SHEET-7 · medium · effort M** — There is no UI to apply a number format or a cell style
(only preservation of what the file already had), and no UI at all to reorder sheet tabs —
no drag, no context menu, no move buttons. The save-side support for sheet reorder exists
and is tested; the control to invoke it does not.

**SLIDE-1 · medium · effort S** — A shape's rotation renders and persists correctly, but
there is no control to set or change it — only the eight resize handles exist.

**UX-FR · low · effort XS** — In DOCX Find & Replace, Escape closes the panel only when
focus is in the Find field; from Replace or Replace All it does nothing.

**CODE-1 · low · effort XS** — `Ctrl+H` does nothing in the code editor, though
`src/viewers/code/CodeEditor.tsx:5` claims "search/replace (Ctrl+F / Ctrl+H)" and no
`Mod-h` binding exists. It is **not** advertised in the shortcuts dialog, docs or README, so
impact is low — but it *does* work in the DOCX editor (`DocxViewer.tsx:1627`), so the two
editors disagree. Either bind it or correct the comment.

**TEST-8 · high · effort S — the CI package validator gave a clean pass on both invalid
files above.** `scripts/validate-office-file.mjs` checks OPC/OOXML/ODF package structure but
not **attribute datatypes**, so `w:val="#ff0000"` and `abstractNumId="atlas-list-2"` sail
through. This matters beyond these two bugs: the validator is the stated substitute for not
having real Office to test against, so its blind spots define what "structurally valid"
actually means here. Teach it the common simple types (`ST_HexColor`, `ST_DecimalNumber`,
`ST_OnOff`, `ST_TwipsMeasure`) and it would have caught both.

---

## 2.9 Execution status — batches 1 and 2 are merged and pushed

**Batch 1** (`952a61e`) — DOCX-14, DOCX-15, DOCX-16, DOCX-17, TEST-8, SEC-1 all landed.
Gate: tsc clean, eslint clean, 3363 unit tests, bundle gate passed, 86 e2e passed.

**Batch 3** (`a0dcd2b`) — A11Y-1..4, I18N-1, TEST-6, DIRTY-1, DRAFT-1, SESS-1, SHEET-4 and
SHEET-6 all landed. Gate: tsc clean, eslint clean, 3528 unit tests, bundle gate passed,
91 e2e passed.

### CI flake investigation (`3f2d134`) — both causes found, CI green

**Resolved · was a real product defect, not a flaky test** — the close-confirmation IPC
listener (`src/App.tsx`) was subscribed in an effect depending on `[saveFile]`, and
`saveFile` closes over `localMarkdown`, so the app **tore down and re-registered a real
Electron IPC listener on every keystroke**. A close request landing in the window before
React ran the passive effect that re-subscribes ran the *stale* closure, where
`isMarkdownDocument` was still false, routing to `viewerSaveRef.current()` → `false`. Now
reads `saveFileRef.current()` from a mount-once listener. The regression test asserts the
listener is registered exactly once across an open plus two edits (fails pre-fix with
`expected 1, got 4`).

**Resolved · genuine budgeting error** — `tests/e2e/export.spec.ts` nested a 30 s
`waitForViewer` and a 60 s `waitForFile` inside playwright.config.ts's 60 s **outer**
per-test timeout. Each inner budget had been raised independently over time without anyone
noticing they share one outer clock, so neither could ever be honoured. Now
`test.setTimeout(120_000)`, matching the pattern `memory-retention.spec.ts` and
`perf.spec.ts` already use.

**TEST-10 · medium · RESOLVED** (verified 2026-09-28) — the cluster is healthy. Each member
passes alone under `--maxWorkers=1` (`App.shellSession` 37, `App.dirtyState` 1, `App.export`
5) and the three pass together 3 runs out of 3; the full 3839-test suite and four CI runs
have been green since. Resolved by accumulated work rather than one fix: `App.shellSession`'s
case was a real product bug (already recorded), and `App.dirtyState` got the suite-wide
`asyncUtilTimeout` raised to 5s plus a 30s budget on the one assertion that blocks on a
`React.lazy()` chunk. Honest caveat: the resolution is "the budgets are now adequate", not
"the coupling was found and eliminated" — if this area starts flaking again, look for
coupling rather than raising budgets a third time.

Original entry: four tests in the App-shell area now behave differently under
CPU contention: they pass in isolation and fail in a loaded full-suite or CI run.
`App.shellSession.test.tsx`'s close-confirmation case (now fixed — and it turned out to be
a **product** bug), `App.dirtyState.characterization.test.tsx` (passes in the suite, fails
alone — the opposite direction), `App.export.test.tsx`, and whatever couples them. Given
the first one investigated was a real defect, these should not be assumed to be noise.
Worth one focused investigation of the cluster rather than four separate chases.

### Released as 3.7.0 (`2026-09-21`)

Batches 1-3, a fix batch for everything the hands-on sweep found, and part of
batch 4. Gate at release: tsc clean, eslint clean, 3581 unit tests, bundle gate
passed, 114 e2e passed, CI green.

**The sweep was the whole point.** Two hands-on sweeps found 7 defects the green
suite could not see; a re-verification found an 8th. Three of those were
*asserted as correct* by existing tests (the Excel table filter, the Ctrl+W
swallow, and the hyperlink dialog, whose test mocked `window.prompt` so jsdom
passed while real Electron threw). Fixing a test that pins a bug in place is not
weakening it.

### CI-only failures observed tonight — a pattern worth watching

Three distinct failures, all on the `windows-latest` runner, all budget- or
timeout-shaped, all passing locally or on retry:
- `export.spec.ts` twice: first the outer per-test clock (a genuine budgeting
  error, fixed), then `waitForFile`'s own 60s poll. Not reproduced locally
  despite CPU contention, software rendering and 45 sequential launches;
  baseline export is 180-350ms.
- `perf.spec.ts`'s cold-start long-task budget once: a 592ms task against the
  300ms CI-adjusted budget. **Passed on re-run of the same commit**, so variance.

The export handler's only failure signal was `logMainEvent`, writing to a file
nothing in the harness reads, so a genuine `PrintToPdfError` was
indistinguishable from slowness. Main-process stderr is now piped into the
export tests. If it recurs, the log will name the cause.

Open question for a future session: whether the CI runner is simply slower than
several of these budgets assume, or whether something genuinely stalls there.

**PROFILE-LEAK-1 · fixed** — every `PLAYWRIGHT=1` launch minted a `userData`
directory and nothing deleted one: 2,876 directories / 26 GB on the development
machine, growing ~110 per full suite run. Now pruned on launch, bounded and
best-effort. Found while investigating the export failure.

**INSERT-TEXT-THROWS-1 · medium · FIXED** (verified against the code 2026-09-28) —
`commands.ts` now has `chunkParagraphForInsert`, which partitions a paragraph into inert
markers (bookmarks, comment ranges/references, note references, opaque nodes like a bare
`w:oMath` — kept byte-identical and never a caret target) and the editable runs between
them, so typing elsewhere in the same paragraph works and the marker is untouched. The
fix comment in `commands.ts` names this finding. The 2026-09-25 night-run CHANGELOG entry
("typing in a paragraph that contains a comment, a footnote reference or an equation
works") is the user-facing record.

~~`applyInsertText` throws for any paragraph containing a comment range/reference…~~

### 2026-09-24 — the CI red on `95a28a8` was F6 regressing on a slow CPU

**F6b · high (silent data loss) · fixed** — the 3.7.0 F6 fix held only on a fast
machine. With the renderer CPU-throttled through CDP
(`Emulation.setCPUThrottlingRate`), click a cell, type "HELLO", save: the bytes
held `"ELLO"` in 3 of 4 trials at 4-8x, and `"Name"` (the cell untouched, the
exact CI failure) at 0ms/char. The unthrottled 48/48 probe could not see it.
Cause, from a keydown/focusin trace: glide-data-grid's pointer-down handler
`preventDefault()`s the browser's focus move and focuses the grid one
`requestAnimationFrame` later; until then keys go to `<body>` and are lost.
The F6 seed buffer never saw them, because the grid never did.

Tried and rejected: focusing the canvas on mouse-down. That frame is also
when the click's selection takes effect: keys delivered earlier hit a grid
with no selection (click, ArrowRight, ArrowDown x2, "42" put 42 in A1;
caught by `spreadsheet-editor.spec.ts`). The fix (`src/viewers/shared/keyHold.ts`)
holds keystrokes from a grid click until the grid owns focus, then replays
them in order, one per task, with a 250ms fallback if glide's frame never
comes. Proof: the new throttled e2e tests fail on the unpatched build
(`"Name"`, `"ELLO"`, `"ELLO,value"`) and pass after; a 24-trial throttled probe
went from 3/4 wrong to 24/24 correct.

Not proven: why the CI runner's frame was late even with a 1000ms pause after
the click. CI kept no e2e output, so the workflow now uploads `test-results/`
when e2e fails. The spec also now waits on the file's mtime instead of a fixed
1500ms sleep, and prints every saved cell on failure, so "never saved", "edit
lost" and "wrong cell" fail differently.

**F6b, second half · fixed** — CI on `4200064` failed the new 8x/0ms test
with `"Name"` again, but a second save 3s later held `"HELLO"`: **Ctrl+S
overtook the edit.** Two paths, both traced at 24x: the save chord bypassed
keys `KeyHold` was still holding (chords were deliberately not held), or it
ran after every key replayed but before the overlay mounted and applied
`commitOnMount` (that one predates 3.7.0's fix too). `KeyHold` now holds
chords as well while waiting for focus, holds them while an edit is in flight
(the F6 pending seed), and queues everything behind anything held. A seed is
dropped as soon as glide shows it opened no editor (read-only grid or cell),
so it can't hold Ctrl+C/Ctrl+S; 5s is only a last-resort bound (a 2s bound
lost to a 2.45s overlay mount at 24x). New e2e test: 24x, Ctrl+S straight
after Enter. It failed 3/3 on `4200064`'s build and passed 10/10 after.

**TEST-11 · fixed — real, test-visible, very unlikely for a user** —
`App.shellSession.test.tsx` › "a background Save As that resolves onto the
currently-active tab's path…" (CI `9987bfd`: editor `# A (background save…)`
under toolbar `b.md`). Cause: `fileIdentityKeyRef`, which `saveFile`/`saveFileAs`
read to decide "is the saved tab still showing", was updated in a passive
`useEffect`. After a tab switch commits, a save resolving before that effect
runs reads the old tab, takes the "still showing" branch, and adopts the old
tab's content under the new tab's path. Reproduced 3/4 runs (40 repeats each)
under 3 CPU burners; 0/2 alone. Fixed with `useLayoutEffect`: 0 failures in
6 x 40 loaded repeats. A person can hardly hit this window, because Save As is a
modal dialog.

**REF-EFFECT-1 · OPEN, likely lead for TEST-9/10** — the same "ref mirrors
state through a passive effect, read by async code" pattern remains in
`App.tsx`: `saveRef`, `saveAsRef`, `exportContentRef`, `saveFileRef`, and
`dirtyGuardStateRef` (read by the unsaved-changes guard; a stale "not dirty"
right after a commit is how a close could skip the prompt). Reproduce each
under load (3 burners + `{ repeats: 40 }`, see TEST-11) before changing it.

**SHORTCUT-FIELD-1 · OPEN, UX, needs an owner decision** — found by the
installer smoke test: Ctrl+O does nothing while the cursor is in a Word
document's text (it works with focus elsewhere). `isPlainFieldTarget` counts
every contentEditable as a field, and `useUniversalShortcuts` skips
o/n/e/t/b/1/2/3// in fields. Ctrl+O/N/T have no native meaning in a text
field to protect, and the Markdown editor is contentEditable too. Same
class as FIELD-01 (Ctrl+W). Ctrl+B must stay gated.

### 2026-09-25 night run — sweep + fixes (full log: `..\atlas-night\NIGHT-LOG.md` and the three sweep reports there)

Three QA agents drove a frozen build (`76bee8d`) against saved bytes; four fix agents worked in
separate worktrees with disjoint files; the coordinator re-verified every claim on `main`
(tsc -b, eslint, targeted vitest, build, affected e2e) before merging.

**Fixed and merged:** REF-EFFECT-1 (`c86d716`), SHORTCUT-FIELD-1 (`481a548`), text encoding/BOM/EOL
round trip for markdown/code/CSV incl. UTF-16 and cp1252 (SHELL-1/2, SHEET-6/7/8: `fe585fd`, `672645b`),
SHEET-1/3/4/5/9 (`17f0a1a`), SHELL-3/5 + FIXTURE-1 (`35a8ec3`), SHELL-6 crash draft overwritten on
relaunch (`89303cd`), DOCX-1 vertical caret, DOCX-2 margin-click dead caret, INSERT-TEXT-THROWS-1
(docx merge). Every fix has a test shown to fail before and pass after.

**Found by CI's smaller window (after the merge):** DOCX-2 and DOCX-1's page crossing both failed on
e2e-windows and reproduced locally just by forcing a 1024x768 window (local default is 1202x802). Both
were real product bugs: `align-items: center` on a page wider than its scroll container made its left edge
unreachable (fixed: `safe center`); ArrowDown couldn't cross into a page virtualization hadn't mounted yet
(fixed: pin the page and retry the move in the same commit). Both specs now run at 1024x768 and 1400x900.
Lesson: geometry-sensitive e2e specs should pin their window size and cover a small one.

**CI-environment flake fixed in the tests:** `toBeFocused()` also requires the window to have OS focus,
which a fresh window on the runner sometimes lacks ("inactive"); replaced with a DOM-focus check
(`tests/e2e/helpers/domFocus.ts`) where the saved file is asserted afterwards anyway.

**Rejected finding:** SHEET-2 (frozen-pane click off by one) was the sweep harness's own bug; regression
tests now pin the correct mapping.

**Tests that pinned bugs, corrected:** the three `toThrow()` INSERT-TEXT-THROWS-1 assertions;
`CsvViewer.editing.test.tsx` asserting the CSV final newline was dropped.

**Still open:**
- **DOCX-3** — after Save As, focus is lost (typing does nothing until you click). Cause:
  `ViewerRouter.tsx:63` keys `ViewerErrorBoundary` by `file.path`, so the rename remounts the viewer
  (content is correct: `showSessionFile` re-reads the saved file). Likely also loses undo history and
  scroll (inferred, not verified). Fix = keep viewers mounted across a rename, which needs the per-viewer
  state audit that file's own comment describes. Pinned by a `test.fail()` e2e test.
- **CI-only flakes to watch:** the 24x "Ctrl+S straight after typing" test failed once (`35a8ec3`) and
  passed on later runs; the in-flight bound was raised 5s→30s (`0c48ffb`), not proven to be the cause.
  The two markdown Ctrl+2 specs now self-diagnose.
- Windows-1252 save that can't be represented falls back to UTF-8 and reports `encodingFallback`; no UI
  surfaces it yet (needs an i18n string).
- PageUp/PageDown move by about one screen (no paginator is wired to the editor), not an exact page.
- Not swept: legacy `.doc`/`.ppt` (no fixtures), markdown preview rendering details, recent files across relaunch.

### 2026-09-27 — 3.8.0 release verification

Installer: `release\Atlas-Setup-3.8.0.exe`, 136 MB, unsigned, SHA256
`3b83b808bbd9001fb6d7ff604e9d9d324de944a9af86327630255dc4fe108a17`.

**Embedded exe metadata verified against the real artifact** — the check
`docs/RELEASE.md` said "still requires the lead's next real `electron-builder --win`
run and inspecting the resulting exe's Properties -> Details tab". Read off
`release\win-unpacked\Atlas.exe`'s `VersionInfo`: ProductName / FileDescription /
CompanyName all `Atlas`, FileVersion `3.8.0`, ProductVersion `3.8.0.0`,
LegalCopyright `Copyright (c) 2026 Atlas`. ELEC-10 is now confirmed against a built
binary rather than only against `electron-builder.yml` and
`releaseMetadata.test.ts`.

**Step 8 (packaged installer smoke test) — DONE for 3.8.0, including the
upgrade path.** The owner accepted the UAC prompt; everything else was driven
programmatically. The installer was run silently (`/S`) so UAC was the only
interaction needed.

**This also closes P5.3's installer-upgrade-path item**, which had never been
verified against a real prior install. 3.7.0 was installed beforehand, so this was a
genuine upgrade in place, not a fresh install:

| Check | Before | After | Result |
|---|---|---|---|
| Installed exe version | 3.7.0 | 3.8.0 (ProductVersion `3.8.0.0`) | upgraded |
| Control Panel entries | one, `Atlas 3.7.0` | one, `Atlas 3.8.0` | **no duplicate entry** |
| Start Menu shortcut | `Atlas.lnk`, 24/09 | same path, re-stamped 27/09, target `C:\Program Files\Atlas\Atlas.exe` | survived |
| `recent-files.json` | 4210 bytes, 50 entries, sha256 `CDDB0768…` | **byte-identical, same sha256** | preserved |
| `window-state.json` | 1542x913 @168,56 maximized | identical | preserved |
| `Local Storage` (theme, locale, drafts) | 10 files, 111 712 bytes | identical | preserved |
| Installer exit code | — | 0 | clean |

Post-install functional checks on the **installed** build:
- Launched from the Start Menu shortcut: 4 processes, window title `Atlas`.
- Opened a `.docx` by argv (what a double-click does): window title
  `sample.docx — Atlas`.
- File associations: all 5 `Atlas.*` ProgIDs present, and the legacy
  `.doc`/`.xls`/`.ppt` ProgIDs resolve to `C:\Program Files\Atlas\Atlas.exe`.
  Default-Apps integration intact — `HKLM\SOFTWARE\RegisteredApplications` -> `Atlas`
  -> `SOFTWARE\Atlas\Capabilities`, 43 extensions.
- **F6c verified fixed in the shipped artifact, not just on `main`**: the repro run
  against `C:\Program Files\Atlas\Atlas.exe` over CDP at 8x CPU throttle passed
  **5/5** (A2 `0.5`, B2 `45365` both saved). Before the fix the same scenario failed
  5/6 at that rate. Playwright's `_electron.launch()` cannot drive a packaged app
  here, so the app was started with `--remote-debugging-port` and driven via
  `chromium.connectOverCDP` — worth knowing for future packaged-build testing.

**Not done, and deliberately:** uninstall verification. It would remove the owner's
working 3.8.0 install, and the upgrade over 3.7.0 already exercised the
uninstall-and-replace path NSIS runs internally.

**PROGID-1 · minor · FIXED in source 2026-09-28; NEEDS AN INSTALL TEST BEFORE THE
NEXT RELEASE.** `electron-builder` uses each `fileAssociations` entry's `name:` as
the ProgID, and Atlas's names were human labels, so an install wrote ~30
unqualified keys under `HKLM\SOFTWARE\Classes` — `Word Document`,
`Excel 97-2003 Workbook`, `Source Code` — each with an open command pointing at
Atlas.exe. No collision with Microsoft Office (`Word.Document.12`-style ids), and
`build/installer.nsh` separately registers the qualified `Atlas.*` set the
Capabilities block uses. But squatting generic names in a global hive means any
other app following the same electron-builder pattern collides, and Atlas's
uninstaller would then delete keys it does not own.

The manifest field is now `associationProgId` (it was `associationName`, which is
what invited the mistake — electron-builder never treated it as a label) and holds
`Atlas.WordDocument`, `Atlas.Excel972003Workbook`, `Atlas.SourceCode` and so on.
`associationDescription` is untouched and is still what Explorer shows, so no
visible label changes. `Macro-Enabled` becomes `Macro` in the ProgIDs only, because
`Atlas.PowerPointMacroEnabledPresentation` is 40 characters and Microsoft's limit
is 39 — silently, not with an error.

**Corrected risk assessment.** The earlier note here called this "the same
association-identity risk as ELEC-22's legacy `appId`". That was wrong, and worth
saying plainly: `appId` keys upgrade-in-place detection and the Windows uninstall
entry, so changing it breaks the upgrade path itself. A ProgID only names a file
type. An in-place upgrade runs the previous uninstaller first (NSIS's
uninstall-and-replace path), which removes the old generic keys before the new
installer writes the qualified ones — the ordinary flow, not a migration. The
residual risk is orphaned keys if that uninstall step is skipped, which is
untidiness rather than a broken association.

**A latent bug found by the test written for this**, not by inspection: `.txt` and
`.log` shared the ProgID `Plain Text` while carrying different descriptions
("Plain Text" and "Plain Text Log"), so Explorer labelled both with whichever
extension the installer happened to write last. `.log` now has
`Atlas.PlainTextLog`. The test asserts a one-to-one ProgID/description mapping for
exactly this reason, plus the `Atlas.` prefix, the 39-character limit and the
character set.

**How far verification actually got.** In source: 95 manifest entries, the
regenerated `electron-builder.yml` read by hand, and 4 new tests. In the build:
`npm run electron:build` succeeds with the new names, which rules out
electron-builder rejecting any of them — that is the whole of what the build
proves. A string-level check of the produced installer was attempted and is NOT
possible: NSIS keeps its script strings in an LZMA-compressed header, so neither
`grep` (ASCII or UTF-16LE) nor `7z x` finds them, and 7-Zip no longer decompiles
`[NSIS].nsi`. So the registry behaviour is unverified: NOT by installing. Before the next release,
install over an existing 3.8.0 and confirm: double-clicking a `.docx`/`.xlsx`/
`.pdf` still opens Atlas, `HKLM\SOFTWARE\Classes` has the `Atlas.*` keys and no
longer has `Word Document`/`Source Code`/etc., and the Capabilities block still
resolves. That needs an elevated install on a real machine, which is why it is
written down here rather than assumed.

Also verified, as a narrower packaged-asset check before the install: the unpacked
build renders five formats with no console errors —
`release\win-unpacked\Atlas.exe` was launched directly (not `electron:preview`,
which skips NSIS and runs from source) with a real file of five formats, asserting
the format's own rendered DOM appeared and that the renderer logged no console
errors:

| Fixture | Asserted | Result |
|---|---|---|
| `sample.docx` | `.docx-page` renders, text correct | 1 page, no console errors |
| `sample.pdf` | `.pdf-viewer canvas` | 1 canvas, no console errors |
| `sample-multisheet.xlsx` | `.spreadsheet-viewer__grid canvas` | 2 canvases, no console errors |
| `sample.pptx` | `.pptx-viewer` | rendered, no console errors |
| `sample.odt` | `.odt-viewer__body` | rendered, no console errors |

This covers pdfjs's worker, the spreadsheet worker, glide's canvas, the bundled
fonts and the lazy per-format chunks actually resolving from the packaged asar —
the things that break in a package and not in dev. It does not cover installation,
file associations, the Start Menu entry, or uninstall.

### 2026-09-24 — packaged installer smoke test (`docs/RELEASE.md` step 8), 3.7.0

Owner installed `Atlas-Setup-3.7.0.exe` (UAC); the rest was driven against
`C:\Program Files\Atlas\Atlas.exe` with Playwright (`executablePath`) and
checked against saved bytes. Results:
- PASS: the packaged build runs (`isPackaged`, 3.7.0). Start Menu shortcut → `Atlas.exe`,
  window opens.
- PASS: `.docx`/`.xlsx`/`.pdf`/`.pptx`/`.md`/`.csv` registered under
  `OpenWithProgids` (`Atlas.Document` etc., command `"…\Atlas.exe" "%1"`).
  Running that exact command opened the file (title `assoc-test.docx — Atlas`).
  Atlas is not the *default* handler for any of them, so an Explorer
  double-click opens Atlas only via "Open with" or after the user picks it.
- PASS: `.docx` opened from the command line, edited, saved; the text is in
  `word/document.xml`. `.xlsx` via Ctrl+O, A1 edited, saved, read back.
  `.pdf` via Ctrl+O rendered. `validate-office-file.mjs`: both saved files clean.
- PASS: version resource (Atlas, 3.7.0 / 3.7.0.0, copyright) and icon matches
  `build/icon.png`.
- FAIL: Ctrl+O with the cursor in the `.docx` body (SHORTCUT-FIELD-1).
- Ctrl+O's native dialog was stubbed (`dialog.showOpenDialog`), not clicked.
  Step 6 (uninstall) is not done: it removes the owner's install.

### New items found while executing batch 3

**QUIT-DRAFT-1 · medium · RESOLVED** (verified against the code 2026-09-26) — `electron/main.cjs` now has `notifyRendererDiscardThenClose`, which sends the discard notification and races the renderer's acknowledgement against a bounded `DISCARD_ACK_TIMEOUT_MS` before destroying the window; its comment names QUIT-DRAFT-1 and explains why a fire-and-forget send was rejected. Original entry below, kept for the record.

~~a discard at *quit* still leaves the autosave draft behind.~~
`electron/main.cjs`'s `handleWindowCloseRequest` shows its own native message box and, on
Discard, calls `mainWindow.destroy()` with no IPC round-trip to the renderer — unlike the
Save path, which does round-trip. So DRAFT-1's fix (clearing the draft in
`handleUnsavedDialogDiscard`) covers Ctrl+W, the toolbar Close, tab switches and File >
Open, but **not** Alt+F4 / the window X / File > Quit. Needs a symmetric discard
notification in `electron/main.cjs` + `electron/preload.cjs`.

**CTRLW-FOCUS-1 · low · FIXED** (2026-09-28) — rather than make every caller announce
intent, `TabBar`'s focus effect now also INFERS the re-home: when the session list shrank and
the tab that had focus is gone, it focuses the tab that took its place. That covers Ctrl+W,
the toolbar Close and anything added later, with no cooperation from the caller.

Guarded on focus having been inside the tab bar at the time — Ctrl+W pressed while the caret
is in the document must not yank focus up to a close button, which would be worse than the
bug. Two tests, and the first was checked against a neutralised fix to be sure it bites.
**Left deliberately unfixed:** where focus should go when Ctrl+W is pressed from inside the
document (presumably the newly-active document) is a ViewerRouter question, not a TabBar one.
— `src/App.tsx`'s global Ctrl+W calls `closeSessionById` directly,
bypassing `TabBar`'s new `handleClose`, so that one route still does not move focus to the
neighbouring tab. Fixable only from `App.tsx`.

**I18N-TEXT-1 · medium · FIXED** (2026-09-26) — the guard now walks the real TypeScript
AST for `JsxText` nodes across all 83 surface files, with a five-entry allowlist for text
that is identical in every locale (`/`, `%`, `●`, `A`, `Atlas`). Regex could not do this
job; TypeScript was already a devDependency, so no new tooling.

**It was not a hypothetical gap.** On the day the check was written it found **twelve
user-facing strings shipping in English in every locale**: `viewer.odp.expectedBinary`,
`viewer.odp.renderFailed`, `viewer.odp.loading`, `viewer.odt.renderFailed`,
`viewer.odt.trackedChangesBanner`, `viewer.pdf.renderFailed`, `viewer.pdf.preparingPrint`,
`viewer.pptx.expectedBinary`, `viewer.pptx.renderFailed`, `viewer.pptx.loading`,
`viewer.rtf.renderFailed` and `docx.render.drawingPlaceholder` — all now translated in both
catalogues. A French user saw every one of them in English.

Worth recording why they were missed: an earlier measurement reported "zero literal text
children" but had only looked inside `<button>`/`<label>`/`<option>`/`<th>`/`<summary>`.
Every one of these twelve lives in a `<div>` or `<span>`. Verified the guard actually bites
by reintroducing one literal (it failed with the file, line and text) and restoring it.

**FROZENROWS-I18N-1 · low · RESOLVED** (2026-09-26) — now `t('spreadsheet.frozenRowCellAria', { row, col })`, with the French string added and the guard's opt-out entry removed, so the file is scanned like every other one. Covered by `FrozenRowsStrip.test.tsx` (including a French-locale assertion).

~~`src/viewers/spreadsheet/FrozenRowsStrip.tsx` has a hardcoded~~
`Frozen row {r}, column {c}` aria-label. Opted out of the guard with a TODO during batch 3
because a sibling agent owned the file. The **only** genuine i18n gap the inverted guard
found across all 83 files.

**REDO-REPLAY-1 · medium · FIXED** (2026-09-28) — `redo` no longer replays anything. An
undo already holds the document a redo has to produce — it is the document being undone —
so the redo entry records it, along with the state the undo left behind (`fromDocument`) and
the original inverse the undo popped. Redo returns the recorded document (reference-equal to
the pre-undo object, asserted with `toBe`), re-pushes the original inverse rather than an
inverse-of-the-inverse, and derives only the caret position by replaying the command against
`fromDocument` — the pair it was recorded against — inside a try/catch, so a throw costs
the caret position instead of the whole redo.

Handed a document History never produced, redo now drops the stack and returns null rather
than replaying the snapshot over an untracked change. That branch should be unreachable
(`push`, `bumpRevision` and a coalesced keystroke all drop the stack), which is exactly why
it is asserted rather than assumed.

Cost: two document graphs per redo entry, both structurally shared with what the app already
holds, alive only while the user is inside an undo chain. `undo` is unchanged and still
replays — a push can happen against an intermediate document (the delete-then-insert
composite in `Input.ts`), so there is no single state to snapshot there.

5 tests in `editor/__tests__/History.test.ts`.

**SEED-24X-1 · OPEN QUESTION, 2026-09-30** — `spreadsheet-keystroke-seed.spec.ts`'s
24x-throttled "Ctrl+S pressed straight after typing saves the typed text" failed once
("the file was never saved") in the SECOND full-suite run after MATRIX-FLAKE-1's fix.

It is very likely the flake already recorded in this file's CI-only-flakes note — the
same test failed once at `35a8ec3` and passed on later runs, and the in-flight bound
was raised 5s→30s at `0c48ffb` without that being proven to be the cause. It also
passed in the first full run after the fix and in a 37/37 targeted spreadsheet run.

**But it sits in the code path MATRIX-FLAKE-1's fix touched, so "pre-existing" is an
assumption, not a measurement.** The one plausible mechanism by which that fix could
reach this test has since been removed: `KeyHold.drain`'s new overlay-focus gate now
exempts chords, because a Ctrl+S is a window shortcut rather than text for the cell
editor, so making it wait for overlay focus (or moving focus on its behalf) only added
delay and risk to the F6b path that holds it behind an in-flight edit.

**Unmeasured.** The run that would have settled it — that single test, `--repeat-each=5`
— was stopped by the harness when the machine ran low on memory, twice. Do not treat
this as resolved until it has run: `npx playwright test tests/e2e/spreadsheet-keystroke-seed.spec.ts -g "24x" --repeat-each=5`.
Compare against the pre-fix baseline before concluding anything.

**MATRIX-FLAKE-1 · FIXED 2026-09-30** — the cell editor now claims DOM focus for
itself (`TextCellEditor`'s focus guard in `SpreadsheetDataEditor.tsx`).

The overlay mounts, takes the first typed character, and then never gets DOM focus,
so everything typed after it goes to a `<td>` of glide's accessibility table and is
lost. This module's own header had already measured the window — 200-300ms between
the seeding keystroke and the textarea gaining focus, because opening the overlay
makes glide re-render the whole grid. `TextCellEntry` asks for React's `autoFocus`,
which applies once, on mount; if that re-render leaves focus on a `<td>`, nothing
re-asserts it. The guard claims focus and STOPS AT THE FIRST SUCCESS — one that kept
re-focusing would fight the click-away commit path, where focus legitimately leaves
a still-mounted overlay — and only intervenes when focus is somewhere never
legitimate mid-edit (nothing, `<body>`, or a table cell).

**THE FIRST FIX ATTEMPT WAS IN THE WRONG LAYER**, and that is the part worth
remembering. It changed `KeyHold`'s replay path — but the lost keys are typed AFTER
the overlay is open, when `awaitingFocus` is null and the queue is empty, so
`KeyHold` is not holding them and never sees them. The run after it came back with
TWO failures instead of one. This is the fourth fix in this family to be attempted
inside `KeyHold` and the fourth to be wrong about it; `spreadsheet-second-cell-edit.spec.ts`'s
header records the first three. **If you are looking at a lost-keystroke bug in this
grid, establish which element actually has focus before changing anything.**

Two `KeyHold` changes were kept anyway, because both are correct on their own terms:
`release` no longer focuses the canvas out from under a mounted overlay, and `drain`
no longer replays into an overlay that has not got focus yet.

Evidence: 5 new unit tests in `keyHold.test.ts`, 3 of which fail when the fix is
reverted (checked by reverting, not assumed); 37/37 spreadsheet e2e; 166/166 full
e2e twice consecutively, against a baseline of 2 failures in 4 runs. Those tests
also had to stub `document.execCommand`, which jsdom lacks — so `replayKey`'s
textarea branch had no coverage at all before this.

Original diagnosis follows, kept because the diagnostic output is what cracked it.

**MATRIX-FLAKE-1 · mechanism identified 2026-09-29** The diagnostics added on 2026-09-28 fired on the next
occurrence and settled it in one line:

```
the edit overlay never closed after committing two at (1,1) with Enter:
focus=TD[glide-cell-2-1]  overlays=1  overlayValue="t"
```

Read against the guide written for exactly this: focus is on a `<td>` of glide's
accessibility table (**the F6c signature**), the overlay is still open, and it holds
only `"t"` when the test typed `"two"`. So the overlay mounted and received the FIRST
character, then never got DOM focus, and `w`, `o` and the Enter all went to that
`<td>` and were lost.

That is the same failure shape as F6c, in the same path, and F6c's fix does not cover
it. F6c was "the overlay never opened because `KeyHold` was never armed"; this is "the
overlay opened and was left without focus". `holdKeysOnMouseDown` skips arming when an
overlay in `#portal` already has focus — but here an overlay EXISTS without having
focus, which that check reads as "nothing to do".

Frequency: 2 of 4 full 166-test suite runs on 2026-09-29. Never reproducible in
isolation (8/8 with `--repeat-each=4` at 1x and 8x, 14/14 at file scope), so it needs
whole-suite load. Do not chase it with a longer timeout — the overlay never closes,
so no timeout makes it pass, and this session already shipped one 15s band-aid whose
real cause turned out to be a bad wait condition.

Next step for whoever picks this up: the question is which of `SpreadsheetDataEditor`'s
focus paths can leave a mounted `#portal` textarea unfocused. Start at
`holdKeysOnMouseDown`'s early return and `KeyHold.release`'s `heldOutsideGrid` branch —
an overlay that exists but is not `document.activeElement` satisfies neither "the canvas
has focus" nor "focus is held outside the grid", so nothing takes ownership of it.

Original 2026-09-28 entry follows.

**MATRIX-FLAKE-1 · originally recorded 2026-09-28** — `spreadsheet-edit-matrix.spec.ts`'s
UNTHROTTLED "a percent, an ISO date and a formula all survive consecutive edits" case
timed out once waiting for the edit overlay to close after committing with Enter
(20s). Measured since:

- 1 failure in 2 full 166-test suite runs (the first full run passed it)
- 8/8 passing in isolation with `--repeat-each=4`, both 1x and 8x
- 14/14 passing at file scope
- the 8x-throttled twin of the same case passed in the run that failed

So it is not reproducible on demand, which is exactly the shape F6, F6b and F6c each
had before they were pinned down — and this path has been a real bug 3 times out of 3,
so it is NOT being written off as noise. What is different here is that it fails only
under whole-suite load, where ~160 Electron launches have already come and gone.

Deliberately NOT papered over with a longer timeout: this session already shipped a
15s band-aid on a different load-dependent failure, and the real cause turned out to
be a bad wait condition. Instead both waits in `editCell` now throw with the focus
state, the overlay count and the overlay's current value, mirroring
`spreadsheet-second-cell-edit.spec.ts`'s `diag` — the next occurrence will say
whether focus sat on a `<td>` of glide's accessibility table (the F6c signature),
whether the overlay held the typed text, or whether nothing had focus at all.

Next step when it recurs: read the message. If focus is on the canvas and the overlay
holds the full text, it is a commit-path bug; if focus is elsewhere, it is the F6c
family again; if the overlay is empty, the keystrokes never arrived.

**BUNDLE-BASELINE-1 · low · RESOLVED 2026-09-29 (baseline refreshed at the owner's request)** — the stored bundle baseline
(`.sisyphus/baselines/atlas-phase3-bundle.json`) was captured on 2026-09-20 and the
build is now +7.96% total / +7.51% gzip against it. The gate passes (threshold 25%),
and none of that growth is from this session: axe-core is a devDependency and
`@lezer/highlight` was already bundled transitively. It is eight days of legitimate
feature work, but it means a real regression now has less of the 25% window to show
up in. Refreshing the baseline (`node scripts/capture-bundle-baseline.mjs`) is the
intended move for accumulated intentional growth — deliberately NOT done here,
because it also erases the drift record, which is the owner's call rather than a
drive-by. Worth knowing alongside it: the absolute gzip ceiling is 5 MiB and the
build is at 3.99 MiB, so about 1.25 MiB of headroom remains.

**Refreshed 2026-09-29.** New baseline: 267 chunks, 12523.29 KB total,
3994.88 KB gzip, captured at package version 3.8.0. The drift record the refresh
erases is preserved here deliberately — the OLD baseline was 2026-09-20, 11598.59
KB / 3715.42 KB gzip, so nine days of feature work cost +7.97% total and +7.52%
gzip. The 25% window now starts from today's size again. The absolute ceilings did
not move (16 MiB total / 5 MiB gzip / 3 MiB per chunk), so the ~1.25 MiB of gzip
headroom is unchanged and is the figure that actually bounds growth.

**CHARTSHEET-1 · low · CLOSED AS A FEATURE REQUEST (2026-09-28)** — one reason the xlsx
passthrough writer bails to the lossy path is a chartsheet or dialogsheet (no
`<sheetData>`). Reviewed again and deliberately not scheduled: nothing is silently
wrong here. The file still opens, the sheets Atlas understands still render, and the
user is told before a save takes the lossy path. Supporting chartsheets means
modelling and round-tripping chart XML — a feature of comparable size to the
spreadsheet viewer itself, not a fix. It stays recorded so the warning's cause is
known, and should be reopened only as a scoped feature, not carried as a defect.

---

**Batch 2** (`a947adc`) — DOCX-1, DOCX-12, DOCX-2, SAVE-1, SHEET-1, SHEET-2, SHEET-3,
SHEET-5 all landed. Gate: tsc clean, eslint clean, 3434 unit tests, bundle gate passed,
89 e2e passed.

### New items found while executing (verified, not yet scheduled)

> **Reconciled 2026-09-26.** Five entries in this section were already fixed in the
> code and are now marked RESOLVED with the evidence inline: QUIT-DRAFT-1, TEST-9,
> FIXTURE-1, ARROW-VERT-1 and FROZENROWS-I18N-1. FIXTURE-1's fix was even recorded
> in the batch log higher up this very file (`35a8ec3`) without this entry being
> updated, which is the drift mechanism in miniature: fixes land and get logged
> chronologically, but the item list is never re-read against the code.
>
> That cost real time — two separate sessions re-investigated ARROW-VERT-1 and
> TEST-9 from scratch. **If you fix something listed here, mark it here**, or
> re-verify this section against the code before planning off it. Everything left
> below was checked on 2026-09-26 and is genuinely open, except where noted.

**ZIP64-1 · medium · FIXED** (2026-09-26) — now fails CLOSED wherever there is positive
evidence of Zip64: a sentinel in the EOCD entry count, in the central-directory offset, or
in an entry's declared size all raise `SpreadsheetZipBombError` instead of skipping the
budget. Refusing costs nothing real — Zip64 only exists for an archive over 4GB or with
more than 65,535 entries, both orders of magnitude past the 200 MiB/500 MiB budgets, so
such a file would be refused the moment its sizes could be read anyway. The one case that
still returns quietly is "no plain EOCD found at all", which is NOT evidence of Zip64 (it is
what a non-zip or truncated file looks like) and where failing closed would report a merely
corrupt workbook as "too large". Four new tests in `spreadsheetZipBudget.test.ts` cover all
three refusals plus that deliberate exemption.

~~the new synchronous zip budget… deliberately skips Zip64 archives… It fails open there.~~

**F6c · HIGH · SILENT DATA LOSS · FIXED** (found and fixed 2026-09-27) — **the second cell edit of a
session is lost, and the Ctrl+S after it silently does not save.** Reachable by hand: click a
cell and type at once, Enter, click another cell and type at once, Enter, Ctrl+S. The second
cell keeps its old value and the file is never written — `mtime` unchanged 20s later. Nothing
is shown to the user.

Frequency (this laptop, `Emulation.setCPUThrottlingRate`, n=6 per rate): **5/6 at 8x, 1/6
unthrottled.** It reproduces without throttling, so a slow machine is not a precondition.

Mechanism, from `document.activeElement` sampled at each step: committing an edit leaves
focus on a `<td>` of glide-data-grid's own accessibility table instead of
`canvas[data-testid="data-grid-canvas"]`. `KeyHold` then replays its held keystrokes to
`document.activeElement` — that `<td>` — so they do nothing. The overlay opens (glide's `<td>`
handler sees the key) but never receives the text, so Enter cannot commit it; and because
`KeyHold` holds chords while an edit is in flight, the Ctrl+S queues behind an edit that can
never commit and is never replayed. That is why the file is not written rather than written
with stale text.

**THE FIX** — two halves in two files, and they only work together. Everywhere the question
"does the grid have focus?" is asked, it must be asked about the CANVAS
(`canvas[data-testid="data-grid-canvas"]`, now `GRID_CANVAS_SELECTOR`/`gridCanvas()` in
`keyHold.ts`) and not about the grid container:
  1. `SpreadsheetDataEditor.holdKeysOnMouseDown` skipped arming `KeyHold` whenever
     `event.currentTarget.contains(document.activeElement)` — which glide's `<td>` satisfies.
     It now skips only when the canvas itself, or an overlay in `#portal`, already has focus.
  2. `KeyHold.release` required `!grid.contains(document.activeElement)` before focusing the
     canvas, so it did nothing while focus was on that `<td>`. It now focuses the canvas
     unless focus is held outside the grid.

**Half 1 is why the first three attempts failed.** They all changed `KeyHold`, which could not
help: `KeyHold.start` was never called at all. Recorded so nobody re-treads them — (a) canvas
check in `release` only: 3/3 -> 1/3 at 8x; (b) treating intra-grid focus moves as
not-moved-away: no change; (c) re-asserting canvas focus before every replay: still 5/6 at 8x.

Verified by control experiment, not assumption: both halves -> **6/6 pass at 8x**; reverting
only half 1 and keeping half 2 -> **6/6 fail**. Before the fix: 5/6 fail at 8x, 1/6
unthrottled.

Regression test: `tests/e2e/spreadsheet-second-cell-edit.spec.ts`, throttled 8x on purpose
(1-in-6 at full speed would not reliably catch a regression). This was also the cause of
`spreadsheet-grid-fixes.spec.ts`'s SHEET-4/5 test going red intermittently — that was true
signal, correctly not dismissed as flake.

**SHEET45-FLAKE-1 · low · FIXED** (2026-09-26) — `spreadsheet-grid-fixes.spec.ts`'s
`typeAt` helper failed once in a full e2e run (`#portal textarea` still present 5s after
Enter) while passing 4/4 in isolation. Investigated rather than retried, because this path
has a 2-for-2 record of load-dependent failures being real (F6, F6b). A CPU-throttled probe
split the question: asserting the SCREEN failed 3/3 at 8x, but asserting the SAVED BYTES
passed 6/6 at both 1x and 8x. So the data path is correct and only the helper's intermediate
screen checks were racing glide-data-grid's overlay lifecycle — exactly what
`spreadsheet-keystroke-seed.spec.ts`'s header warns about, and what `helpers/domFocus.ts`
already recorded this spec flaking on twice. `typeAt` now waits for the overlay to exist
before asking whether it is focused, and gives all three waits a 20s budget (overlay mount
was measured at 2.45s at 24x, so the 5s default was marginal). The end-of-test byte
assertions were always the real verification and are unchanged.

**RPR-STYLES-1 · medium · FIXED** (2026-09-28) — `parser/styles.ts` now captures every
unmodeled `w:rPr` child into the same `RunProps.rPrUnknown` the document path uses, and
`serializer/stylesWriter.ts` splices each one back ahead of the modeled sibling it sat
before, so the `CT_RPr` sequence stays schema-ordered. The non-order-preserving XML shape
turned out not to be the obstacle it looked like: `fast-xml-parser` inserts keys in document
order and none of these names are integer-like, so `Object.keys` IS document order. The
passthrough mechanism (placeholder element per fragment, substituted into the finished
string) lives in the new `serializer/rawPassthrough.ts`.

Two things the entry did not anticipate, both found while implementing:
- `writeStylesXml` hardcoded `xmlns:w` plus `xmlns:r` and emitted no other root attribute.
  Harmless while nothing it emitted used another prefix — but re-emitting a `w14:`/`w15:`
  run property under a root that never declares the prefix is not well-formed XML and Word
  rejects the whole file. Both writers now carry the source root's own attributes through.
  Fixing the data loss without this would have shipped a worse bug.
- `numbering.xml` shares `parseRunPropsNode` and `buildRunPropertiesXml`, so it inherited
  the capture and with it the obligation to restore the placeholders; without that it would
  have written a literal `<atlas-raw-0/>` into the file.

13 tests in `serializer/__tests__/stylesRPrPassthrough.test.ts`, all asserting the
serialized XML — the model holding a fragment says nothing about the file getting it back.

~~the general unknown-`w:rPr`-child passthrough covers
`document.xml` (and, through `partWriterSupport.ts`, headers/footers/notes/comments) but
**not `styles.xml`**. `src/docx/parser/styles.ts` uses a non-order-preserving XML shape, so
an equivalent passthrough there is a separate, larger change. Unknown run properties in
`w:docDefaults` and named styles are still dropped on save. Theme fonts *are* fixed there.

**TEST-9 · medium · RESOLVED** (verified 2026-09-26) — ran the exact command this entry names, `npx vitest run --maxWorkers=1 src/__tests__/App.dirtyState.characterization.test.tsx`: it passes. The file's `waitFor` for `.markdown-body` was reworked when `MarkdownRenderer` became `React.lazy()` (PERF-01), and the suite-wide `asyncUtilTimeout` was raised from Testing Library's 1s default to 5s in `src/__tests__/setup.ts`. Note the failure direction had ALSO inverted at one point — it failed inside the full suite and passed alone — which is what the `asyncUtilTimeout` change fixed.

~~`src/__tests__/App.dirtyState.characterization.test.tsx` **passes in~~
the full suite and fails when run alone** (`npx vitest run --maxWorkers=1 <that file>`),
confirmed on clean `main` at `952a61e`. It depends on state or timing from other tests, so
it cannot be trusted as a regression signal, and it silently costs nothing when it breaks.

**PARTWRITER-1 · resolved during batch 2, recorded for the register** —
`partWriterSupport.ts`'s `buildBlockNodes` never consulted wrapper regions at the block
level at all; only `documentWriter.ts`'s private copy did. A `w:sdt` wrapping a whole
paragraph or table in a footnote could never have round-tripped, independently of DOCX-2's
threading gap. Found and fixed by the DOCX-2 agent.

**FIXTURE-1 · low · RESOLVED** — fixed in `35a8ec3` (already recorded in the batch log above at "SHELL-3/5 + FIXTURE-1"; this entry was simply never reconciled). Verified 2026-09-26: `generate.mjs` has a reordering helper whose header cites FIXTURE-1, and every ODF fixture writes `mimetype` first with `{ compression: 'STORE' }`.

~~`tests/e2e/fixtures/generate.mjs:407` produces a non-conforming~~
`.ods`: `mimetype` is not the first zip entry (`odf-mimetype-not-first`, reproduced on
`main`). The save-path fix for this landed in an earlier wave; the *fixture generator* never
did. So every test asserting "Atlas opens `.ods`" proves it against a file no real tool
would produce.

**INSERT-TABLE-CURSOR-1 · low · OPEN, and blocked on something else than this entry
assumed — attempted and reverted 2026-09-28.** The premise ("cells are now addressable, so
it is a small follow-up") is only half right. Cell PATHS do resolve —
`PageView.renderPageTableRow` attaches one per cell — but paths are attached to rendered
LINES, and a freshly inserted table's cells are empty, so there is no line and therefore no
DOM node for a caret to anchor to. Measured: with the cursor changed to
`[section, ...prefix, index + 1, 0, 0, 0]`, `syncSelectionToDom` still cannot resolve it, the
DOM selection never moves, and typed text lands on path `2` (the paragraph after the table)
exactly as before — diagnostic from the real app read
`hits: ["path=2 inTable=false"], tableText: ""`, the empty `tableText` being the tell.

So the real blocker is **empty-cell caret anchoring**, not path resolution: a cell with no
content renders no line box. Fixing this needs either a zero-width/placeholder line box per
empty cell paragraph in `layoutTable.ts`, or a `positionToDomRange` fallback that anchors to
the cell element itself when the cell has no lines. Both are larger than "small", and the
DOCX-16 failure mode if it is got wrong is that every keystroke after Insert Table vanishes
silently — which is why `docx-insert-table.spec.ts` asserts the landing paragraph explicitly.
Reverted to the DOCX-16 behaviour (park on the paragraph after the table), which that test
still passes.

Original entry: batch 1's Insert Table fix parks the cursor on the
addressable paragraph *after* the table, because cells were unreachable at the time. Cells
are now addressable, so moving it into the first cell (Word's actual behaviour) is a small
follow-up plus a test update.

**ARROW-VERT-1 · medium · RESOLVED** (verified against the code 2026-09-26) — both halves of this entry's root cause are gone. `DocxViewer.tsx` intercepts `VERTICAL_MOVE_KEYS` (ArrowUp/ArrowDown/PageUp/PageDown) and **always** `preventDefault()`s before anything else, so the keys never fall through to native caret movement; `Input.ts`'s null return is now documented as a deliberate handover to that interception, which resolves the move via `Cursor.ts`'s line geometry. Covered end to end by `tests/e2e/docx-vertical-caret.spec.ts`, and a deferred-retry path for a target page virtualization has not mounted yet lives in `src/docx/render/useDocxVerticalCaret.ts` (DOCX-SMALL-WINDOW-2). The CHANGELOG's 2026-09-25 night-run section records the user-visible fix.

~~ArrowUp/ArrowDown misbehaviour is now~~
**confirmed by running the app**: in a plain three-paragraph document with **no tables**,
ArrowDown from paragraph 0 jumps to paragraph 2, and a second press does not move. Separate
root cause from the table-caret bug: `Input.ts:696-700` returns `null` for vertical movement
and `DocxViewer.tsx:1690-1710` does not `preventDefault`, so the keys fall through to native
caret movement on the absolutely-positioned-per-line DOM.

---

## 3. Batch execution plan

### Rules that apply to every batch

**Collision map — these files may have exactly one owner per batch:**
`src/App.tsx` · `src/viewers/DocxViewer.tsx` · `electron/main.cjs` ·
`src/docx/serializer/documentWriter.ts` · `src/viewers/spreadsheet/xlsxPassthrough.ts` ·
`src/i18n/messages.en.ts` + `messages.fr.ts` · `CHANGELOG.md` ·
`docs/KNOWN_LIMITATIONS.md`.

**Must stay sequential — never two at once, in the same batch or across overlapping work:**
- the DOCX serializer (`src/docx/serializer/**`),
- the spreadsheet passthrough (`xlsxPassthrough.ts` + `spreadsheetRangeShift.ts` + `spreadsheetTables.ts`),
- the main-process IPC surface (`electron/main.cjs` + `electron/preload.cjs`).

**i18n and CHANGELOG are written by the session lead at merge time**, not by agents — they
are pure collision surface and every agent would otherwise touch them.

**Verification gate between batches** (in this order, never concurrently — this laptop has
hard-powered-off under full CPU load):
```
npx tsc -b
npx eslint src electron tests scripts
npx vitest run --maxWorkers=3
npm run build
npx playwright test
git checkout -- tests/e2e/fixtures/
```
then push and watch CI. A red gate blocks the next batch. If an e2e test flakes only on CI,
suspect the product before the test — that has been a real bug twice.

**Every agent:** works in its own worktree; removes the `node_modules` junction with
`cmd /c rmdir <path>\node_modules` before `git worktree remove` (never `rm -rf` through a
junction); never runs `npm install` in a worktree; runs `npx vite build` before any Electron
probe; kills Electron with `taskkill /PID <pid> /T /F`.

---

### Batch 1 — "Word may refuse the file, and typing is being lost" (5 agents)

Re-ordered after the hands-on sweep. Everything here is wrong **every time you use the
feature**, so it outranks the race conditions and save-path losses that follow.

| Agent | Objective | Owns | Must NOT touch | Acceptance | Tests to add |
|---|---|---|---|---|---|
| **1A** | DOCX-14 + DOCX-15: make invalid attribute values unrepresentable. Normalise and validate inside `hexColor()` itself so no call site can reintroduce a `#`; mint `abstractNumId` as an integer. | `src/docx/model/styles.ts`, `src/docx/editor/toolbarAdapter.ts`, `src/docx/editor/insertList.ts`, `src/docx/editor/pasteBlocks.ts` | `src/docx/serializer/**`, `src/docx/parser/**` | Applying a colour writes `w:val="ff0000"`; a toolbar-inserted table border likewise; a new list writes an integer `abstractNumId`; a colour read from a file still round-trips unchanged. | unit: `hexColor('#ff0000')` normalises, `hexColor('zzz')` rejects; unit: toolbar colour + table border + new list produce schema-valid values; round-trip: a file colour is untouched |
| **1B** | DOCX-16: Insert Table must produce a usable table, and typing into it must never be silently discarded. | `src/docx/editor/commands.ts`, `src/docx/layout/layoutTable.ts` | `src/docx/editor/Input.ts`, `DocxViewer.tsx` | A toolbar-inserted table renders at a real width, is clickable, and accepts typed text that survives save/reopen. | unit: `buildEmptyTable` geometry; e2e: insert table, type, save, reopen, assert the text is in `document.xml` |
| **1C** | DOCX-17: the caret must land in the table cell that was clicked, and Tab/arrows must move between cells. | `src/docx/editor/Input.ts`, `src/viewers/DocxViewer.tsx` | `src/docx/editor/commands.ts`, `src/docx/layout/**`, `src/App.tsx` | Clicking a cell in a pre-existing table places the caret there; Tab moves to the next cell; cell merge becomes reachable. | unit: hit-testing into a cell; e2e: click a cell, type, assert it landed in that cell |
| **1D** | TEST-8: teach the package validator attribute datatypes. | `scripts/validate-office-file.mjs` | everything under `src/` | The validator **fails** on today's pre-fix output for DOCX-14 and DOCX-15, and passes after 1A. Covers at least `ST_HexColor`, `ST_DecimalNumber`, `ST_OnOff`, `ST_TwipsMeasure`. | fixture-based: one invalid file per datatype |
| **1E** | SEC-1: split write-eligibility out of the path allowlist so a renderer-registered path can never satisfy `existingPath`. | `electron/main.cjs`, `electron/lib/pathAllowlist.cjs` | anything under `src/` | A path added only via `path:register-dropped` is refused by `save-file`/`save-binary-file` (dialog shown instead); real drag-drop open still works; real Save still works. | unit: allowlist trust tiers; e2e: drag-drop open then Save still writes |

*1B and 1C both sit in the DOCX editor and are deliberately split along the
commands/layout vs. input/hit-testing seam. If 1C cannot avoid `commands.ts`, run it in
batch 2 instead — do not let both edit it.*

---

### Batch 2 — silent loss on save (5 agents)

| Agent | Objective | Owns | Must NOT touch | Acceptance | Tests to add |
|---|---|---|---|---|---|
| **2A** | DOCX-1 + **DOCX-12**: emit theme font attributes, and add a general **unknown-run-property passthrough** so an unmodelled `w:rPr` child survives a round-trip. Then model the common effects (`w:outline`, `w:emboss`, `w:imprint`, `w:em`, `w:effect`, run `w:bdr`, `w:w`) for rendering. | `src/docx/model/styles.ts`, `src/docx/serializer/stylesWriter.ts`, `src/docx/serializer/documentWriter.ts`, `src/docx/render/style.ts` | `src/docx/parser/partBody.ts`, the notes writers | A theme-font style round-trips with its theme reference; a run carrying any named effect round-trips byte-identical; **a newly invented `w:rPr` child also survives**, proving the passthrough is general and not a longer fixed list. | unit: theme-only and mixed literal+theme; unit: each effect; unit: an unknown `w:rPr` child survives; corpus fixture with emphasis marks |
| **2B** | SAVE-1: port the markdown `stillShowing` pattern to `handleViewerSavedPath` and the three viewer call sites. | `src/App.tsx`, `src/viewers/shared/ViewerContext.tsx`, `src/viewers/DocxViewer.tsx`, `src/viewers/slides/shared/useSlideEditorCore.ts`, `src/viewers/spreadsheet/useSpreadsheetEditor.ts` | `xlsxPassthrough.ts`, `spreadsheetGrid.ts`, `spreadsheetDocument.ts` | A Save As resolving after a tab switch updates the saved document's own tab and leaves the showing tab untouched. | unit: `reportSavedPath` after `filePath` changed; e2e: Save As then switch tab before the write resolves |
| **2C** | SHEET-2 + SHEET-3: stop wiping filter/sort on untouched tables; teach `decodeCell`/`remapSqref` whole-row and whole-column ranges. | `src/viewers/spreadsheet/spreadsheetTables.ts`, `src/viewers/spreadsheet/spreadsheetRangeShift.ts`, `src/viewers/spreadsheet/xlsxPassthrough.ts` | `useSpreadsheetEditor.ts`, `spreadsheetGrid.ts`, `spreadsheetDocument.ts` | Editing an unrelated cell and saving preserves a table's `sortState`/`filterColumn`; a `conditionalFormatting` on `A:A` survives a row insert, shifted correctly. | unit: `remapSqref("A:A", …)` and `("1:1", …)`; unit: unchanged table keeps filter state; e2e: save/reopen with a whole-column CF |
| **2D** | SHEET-1 + SHEET-5: route `parseWorkbookBuffer` through the existing zip budget; give `recalculateSheet` a dependency order with a cycle guard. | `src/viewers/shared/spreadsheetGrid.ts`, `src/viewers/spreadsheet/spreadsheetDocument.ts` | the passthrough files, `useSpreadsheetEditor.ts` | A workbook with an implausible declared uncompressed size is refused before inflation on the viewer, Worker and PDF-export paths; a formula depending on a later formula cell is fresh on the first pass; a circular reference gives a defined error, not a hang. | unit: budget rejection on a crafted central directory; unit: forward and backward formula dependencies; unit: cycle |
| **2E** | DOCX-2: carry `wrapperRegions` into footnotes/endnotes/comments, and add those three parts to the lossy-save detector. | `src/docx/parser/partBody.ts`, `src/docx/parser/document.ts`, `src/docx/fidelity/lossySaveWarnings.ts`, `src/docx/serializer/footnotesWriter.ts`, `endnotesWriter.ts`, `commentsWriter.ts` | `documentWriter.ts`, `stylesWriter.ts`, `src/docx/model/styles.ts` | An unedited `w:sdt` in a footnote round-trips byte-identical; if it cannot, the user is warned. | unit: footnote/endnote/comment wrapper round-trip; unit: detector covers the three parts |

*2A and 2E are the two DOCX agents and own disjoint files — 2A the model/serializer,
2E the parser and the notes writers. 2A owns `model/styles.ts`; 2E must not touch it.
2C and 2D are the two spreadsheet agents, split passthrough vs. viewer/document.*

**2A is the one to watch.** The general passthrough is worth more than the named
properties: it is the structural answer to "how did a whole class of formatting get silently
dropped", and it retires the next DOCX-12 before it is written.

---

### Batch 3 — trust, focus and the English UI (5 agents)

| Agent | Objective | Owns | Must NOT touch | Acceptance |
|---|---|---|---|---|
| **3A** | A11Y-1…A11Y-4: one focus-management pass. Migrate the two hand-rolled traps onto `useFocusTrap`; give `PresenterView` and the PDF find bar real traps; restore focus after a tab close. Correct the docs. | `src/components/ShortcutsModal.tsx`, `UnsavedChangesDialog.tsx`, `TabBar.tsx`, `src/viewers/shared/PresenterView.tsx`, `src/viewers/pdf/PdfFindBar.tsx`, `src/hooks/useFocusTrap.ts`, `docs/KNOWN_LIMITATIONS.md` | `src/App.tsx` — the tab-close fix must be self-contained in `TabBar.tsx` | Tab cycles within every overlay including with a disabled control present; presenter view traps Tab with fullscreen refused; closing a tab by keyboard leaves focus on a sensible neighbour, never `<body>`. |
| **3B** | I18N-1 + TEST-6: translate `RawEditor` and give its textarea a real accessible name; **invert the i18n guard from an opt-in allowlist to scan-everything plus an explicit, commented opt-out list.** | `src/components/RawEditor.tsx`, `src/i18n/messages.en.ts`, `messages.fr.ts`, `src/i18n/__tests__/noHardcodedStrings.test.ts` | everything else | `RawEditor` is translated and has an accessible name; **the inverted guard fails on today's `RawEditor` before the fix and passes after**; its opt-out list is explicit and justified per entry. |
| **3C** | DIRTY-1: dirty tracking survives undo back to the saved state. | `src/viewers/DocxViewer.tsx`, `src/docx/editor/History.ts` | `src/App.tsx`, the serializer | Save, edit, undo → clean, and closing does not prompt. Prefer a content/revision identity over reference equality; do **not** deep-compare a large document on every keystroke. |
| **3D** | DRAFT-1 + SESS-1: clear the draft on the guard dialog's Discard; de-duplicate `renameSession`. | `src/App.tsx`, `src/hooks/useAutosave.ts`, `src/session/documentSessions.ts` | the viewers | Discarding via Ctrl+W/tab switch/Open/quit leaves no draft; Save As onto an already-open path yields one session, not two. |
| **3E** | SHEET-4 + SHEET-6: make the passthrough→lossy fallback visible; evaluate text-returning formulas. | `src/viewers/spreadsheet/xlsxPassthrough.ts`, `useSpreadsheetEditor.ts`, `src/viewers/spreadsheet/formula/**` | `spreadsheetTables.ts`, `spreadsheetRangeShift.ts`, `spreadsheetGrid.ts` | A save that falls back says in plain language what may be lost; `=A1&"!"` and `=CONCATENATE(...)` evaluate and save as values. |

*3C and 3D both touch the shell but own disjoint files — 3C `DocxViewer.tsx`, 3D `App.tsx`.
3B is the only agent that may touch the i18n catalogues, which is why it owns them alone.*

---

### Batch 4 — the tests that stop all of this recurring (5 agents)
TEST-1 + TEST-2 + TEST-3 (corpus edit targets, attribute-level fingerprints, real export
assertions) · TEST-4 (make `lossySaveWarnings` see attribute values and count deltas,
regression-checked against DOCX-1's and DOCX-12's pre-fix shapes) · TEST-5 (draft and
session-collision coverage) · TEST-7 (a profile-directory override so recent files, window
state and the `recentFilesStore` security check are testable across a relaunch).

Scheduled **after** the fixes deliberately: each of these is written to fail against the
pre-fix behaviour, which is the only way to know the test is worth keeping.

### Batch 5 — remaining fidelity (5 agents)
DOCX-3 (paragraph-mark `rPr`) · DOCX-4 + DOCX-5 (`docPr` ids, anchor attributes) ·
DOCX-13 (hidden text) · FID-2 (hidden rows/columns) · FID-4 + FID-5 (PDF destinations and
annotation text) · SEC-2 + SEC-3 (printToPdf guards, legacy encryption detection) ·
SHEET-7 / SLIDE-1 / UX-FR / CODE-1 (missing controls and small UX fixes).

### Batch 6 — the large work, only if there is a reason
FID-1 (per-cell spreadsheet styling, L) · FID-3 (cell comments, M) · FID-6 (PDF layers, M) ·
DOCX-6…DOCX-11 (M each) · SEC-4 (`patch-package` or a watchdog for the vendored SheetJS CFB
cycle — see §4).

### Cost shape

| Batch | Agents | Why it is worth it |
|---|---|---|
| 1 | 5 | Files Word may refuse to open, and an editor that silently discards what you type into a table. Not optional. |
| 2 | 5 | Every remaining silent loss on save, plus the general passthrough that prevents the next one. |
| 3 | 5 | Dirty-state trust, the focus pass, and the English/French surface you asked about. |
| 4 | 5 | The tests without which batches 1–3 can silently regress. Written to fail pre-fix. |
| 5 | 5 | Fidelity polish and missing controls. Safe to defer. |
| 6 | ~5 | Large features. Defer until there is a reason. |

**Batches 1–4 (20 agents) are the ones I would actually run**, in that order, with the gate
between each. Batches 5–6 are a queue, not a commitment — re-scope them once 1–4 are merged
and CI is green.

---

## 4. NOT VERIFIED — needs a run before it is scheduled

None of these are in a batch above. Each needs evidence before it earns a slot.

- **DOCX vertical caret movement.** `Input.ts:696-700` returns `null` for
  ArrowUp/ArrowDown/PageUp/PageDown with a "deferred, needs paginator" TODO, and
  `DocxViewer.tsx:1690-1710` does not `preventDefault`, so the browser handles them on a DOM
  where every line is `position: absolute` (`page-view.css:38`) — the same DOM the code
  elsewhere says makes native caret placement unreliable. **Plausible and important**
  (Up/Down is a core editing interaction) but nobody has watched it misbehave. The hands-on
  sweep should settle this.
- **New cells written with no style attribute.** `xlsxPassthrough.ts:366` takes
  `originalCell?.getAttribute('s') ?? null`; typing into a cell that had no `<c>` element
  writes no `s`. OOXML's row/column cascade may resolve it correctly. Confirming needs real
  Excel, which is not installed.
- **SheetJS CFB cyclic sector chain.** `get_sector_list` in the vendored
  `node_modules/xlsx/xlsx.js` lacks the cycle check its sibling `make_sector_list` has, so a
  crafted `.doc`/`.ppt`/`.xls` could hang. Source asymmetry is clear; no crafted file was
  built to confirm, and the fix lives in a dependency.
- **Path allowlist and symlinks.** `normalizePath` never calls `fs.realpathSync`. Requires an
  attacker to have pre-planted a reparse point, which is a separate primitive.
- **SHEET-4 trigger rate.** The silent-fallback mechanism is verified; how often it fires on
  real workbooks is unmeasured.

---

## 5. Part 1 — hands-on feature matrix

Two agents drove the real application, each under `ATLAS_HIDDEN_WINDOW=1`. **Neither run was
headless** — a real window existed and was composited throughout; it was simply transparent,
unfocused and click-through. Full matrices, with repro steps for every non-green row:

- `.sisyphus/plans/atlas-phase4-sweep-A-matrix.md` — shell, new documents, markdown, code
  files, PDF, legacy `.doc`/`.ppt` viewers.
- `.sisyphus/plans/atlas-phase4-sweep-B-matrix.md` — DOCX editing, spreadsheets, slides.

| Area | works | partly | broken | not tested |
|---|---|---|---|---|
| Sweep A — shell / new doc / markdown / code / PDF / legacy | 53 | 0 | 1 | 7 |
| Sweep B — DOCX / spreadsheets / slides | 35 | 1 | 9 | 2 |
| **Total** | **88** | **1** | **10** | **9** |

Zero `pageerror` or console errors were captured on any sweep-A run.

### What this changes

**The shell is in good shape.** 53 of 54 exercised rows work, and the single failure
(CODE-1, `Ctrl+H` in the code editor) is not advertised to users anywhere. Tabs, drag and
drop, themes, the English/French switch, the documented shortcuts, the unsaved-changes
prompt, window-close-with-dirty-documents, markdown preview/split/editor, Mermaid, KaTeX,
the table of contents, all three exports, code editing and Run/Stop for JS/TS/Python, PDF
search/zoom/rotate/thumbnails/navigation/password/print, and both legacy viewers: all
confirmed working by actually using them.

**The document editors are not.** Nine broken rows, and the four worst
(DOCX-14…DOCX-17) are now the top of batch 1. Sweep B verified save→reopen→inspect-the-bytes
for everything it passed, which is why its "works" rows are worth more than a count: typing,
bold/italic/underline/strikethrough, font family and size, highlight, alignment, undo/redo
across a save, find and replace, plain *and* mixed field/image headers and footers, insert
image, insert hyperlink, zoom, spreadsheet cell editing, number-format preservation,
row/column insert and delete, sheet add/rename/delete, Excel tables, ODS, and the whole
slide surface (text editing, move, resize, insert text box, add/duplicate/delete/reorder,
notes, presenter view, ODP) all survived the round trip to disk.

**One caveat the sweep raised, and it is important:** `scripts/validate-office-file.mjs`
passed *clean* on both invalid-OOXML DOCX files. A clean validator run does not mean the
file is schema-valid — see TEST-8. Given the validator is the stated substitute for not
having real Office to test against, its blind spots define what "valid" currently means here.

### Not tested, with reasons (9)

Sweep A: the Python-absent Run path (Python 3.14 is installed on this machine);
recent-files persistence across relaunches (structurally impossible — see TEST-7);
window-close-with-Save (already covered by `close-confirmation.spec.ts`); three PDF-password
sub-cases (owner password, cancel, re-prompt on reopen). New-document overwrite onto an
existing file relies entirely on the native save dialog's own warning — there is no
Atlas-level confirmation (`electron/main.cjs:995-1017`) — and that native layer cannot be
driven, because the tests must stub `showSaveDialog`.

Sweep B: DOCX live spell-check misspelling and suggestion UI (needs a real OS dictionary and
a native context-menu event); DOCX page-navigation controls (only zoom was exercised).

These are honest gaps, not green rows. Nothing above claims a status that was not observed.
