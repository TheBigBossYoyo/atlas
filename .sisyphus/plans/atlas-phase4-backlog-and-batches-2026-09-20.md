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
| Closed spreadsheet tab retains ~24 MB | Documented in `docs/KNOWN_LIMITATIONS.md` with its measurement; not re-measured this session. |
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

**Batch 2** (`a947adc`) — DOCX-1, DOCX-12, DOCX-2, SAVE-1, SHEET-1, SHEET-2, SHEET-3,
SHEET-5 all landed. Gate: tsc clean, eslint clean, 3434 unit tests, bundle gate passed,
89 e2e passed.

### New items found while executing (verified, not yet scheduled)

**ZIP64-1 · medium** — the new synchronous zip budget in `spreadsheetZipBudget.ts`
(`checkWorkbookZipBudgetSync`) reads declared sizes off the zip central directory, but
**deliberately skips Zip64 archives** rather than risk misreading a 32-bit sentinel. It
fails open there. Honest, and no worse than before the guard existed — but an attacker
choosing Zip64 bypasses it entirely, which is not a property a security guard should have.

**RPR-STYLES-1 · medium** — the general unknown-`w:rPr`-child passthrough covers
`document.xml` (and, through `partWriterSupport.ts`, headers/footers/notes/comments) but
**not `styles.xml`**. `src/docx/parser/styles.ts` uses a non-order-preserving XML shape, so
an equivalent passthrough there is a separate, larger change. Unknown run properties in
`w:docDefaults` and named styles are still dropped on save. Theme fonts *are* fixed there.

**TEST-9 · medium** — `src/__tests__/App.dirtyState.characterization.test.tsx` **passes in
the full suite and fails when run alone** (`npx vitest run --maxWorkers=1 <that file>`),
confirmed on clean `main` at `952a61e`. It depends on state or timing from other tests, so
it cannot be trusted as a regression signal, and it silently costs nothing when it breaks.

**PARTWRITER-1 · resolved during batch 2, recorded for the register** —
`partWriterSupport.ts`'s `buildBlockNodes` never consulted wrapper regions at the block
level at all; only `documentWriter.ts`'s private copy did. A `w:sdt` wrapping a whole
paragraph or table in a footnote could never have round-tripped, independently of DOCX-2's
threading gap. Found and fixed by the DOCX-2 agent.

**FIXTURE-1 · low** — `tests/e2e/fixtures/generate.mjs:407` produces a non-conforming
`.ods`: `mimetype` is not the first zip entry (`odf-mimetype-not-first`, reproduced on
`main`). The save-path fix for this landed in an earlier wave; the *fixture generator* never
did. So every test asserting "Atlas opens `.ods`" proves it against a file no real tool
would produce.

**INSERT-TABLE-CURSOR-1 · low** — batch 1's Insert Table fix parks the cursor on the
addressable paragraph *after* the table, because cells were unreachable at the time. Cells
are now addressable, so moving it into the first cell (Word's actual behaviour) is a small
follow-up plus a test update.

**ARROW-VERT-1 · medium — promoted out of §4** — ArrowUp/ArrowDown misbehaviour is now
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
