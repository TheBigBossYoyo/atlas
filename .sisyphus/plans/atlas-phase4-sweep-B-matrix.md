# Atlas 3.6.0 — QA sweep B (DOCX editing, Spreadsheets, Slides)

**Run conditions:** Playwright `_electron.launch()` driving the real built app (`dist/`, unmodified) from standalone Node scripts, `CI=1 PLAYWRIGHT=1 ATLAS_HIDDEN_WINDOW=1`. This is a **visible-but-fully-transparent, click-through window** — not headless; it was created, shown and composited normally, then made invisible to the operator. Every script killed its own Electron process tree via `taskkill /PID <pid> /T /F` before the next one launched. English UI throughout. `node scripts/validate-office-file.mjs` was run against every saved file mentioned below (see the "office validator" note at the bottom).

**Coverage:** I personally drove and observed **47 feature rows** below: 23 in DOCX, 12 in spreadsheets, 12 in slides. Of those, **35 work, 9 are broken (including 2 features — sheet reordering and shape rotation — for which no UI control exists at all), 1 partly works, and 2 were not tested** (the live spell-check misspelling/suggestion UI, which needs a real OS dictionary and native context-menu event I couldn't reliably drive; and dedicated page-navigation controls, as opposed to zoom, which I did test). DOCX Save-As and legacy `.xls`/`.doc`/`.ppt` formats were not in my scope this run (see the bottom of this file). Everything in this file is a first-hand, driven observation from the real app; the two "no UI control exists" findings (sheet reorder, shape rotate) were confirmed both by reading the relevant component and by visually confirming live that no such control appears after selecting the object.

Scripts live in the scratchpad `sweepB/` folder alongside this file; fixtures were written under the OS temp dir (`atlas-*` prefixed), never under the repo, and nothing in `tests/e2e/fixtures/` was touched.

---

## Summary counts

| Status | Count |
|---|---|
| works | 35 |
| broken | 9 |
| partly works | 1 |
| not tested | 2 |
| **Total rows** | **47** |

---

## DOCX editing

| # | Feature | Status | Notes |
|---|---|---|---|
| 1 | Type text at caret | works | |
| 2 | Bold / Italic / Underline / Strikethrough | works | On screen + `<w:b/><w:i/><w:u/><w:strike/>` all present after save+reopen. |
| 3 | Font family | works | Toolbar font picker; `w:rFonts` correct on disk. |
| 4 | Font size | works | Disk `<w:sz w:val="56"/>` for 28pt selection is correct. |
| 5 | **Text/font colour** | **broken** | See detail below. |
| 6 | Highlight colour | works | Named `<w:highlight w:val="black"/>` etc. — valid enum, correct. |
| 7 | Alignment | works | Right-align serializes as `<w:jc w:val="end"/>`, a valid ECMA-376 value (logical "right" in LTR), not a bug. |
| 8 | **Bullets / numbered lists (newly inserted)** | **broken** | See detail below. |
| 9 | Undo/redo, including across a Save | works | Type → Save → (refocus the document) → Ctrl+Z reverts, resave drops it from disk; Ctrl+Y restores it. Note: immediately after clicking the Save **button**, Ctrl+Z/Ctrl+Y do nothing until the editor regains focus — ordinary focus behaviour, not a bug, but worth knowing if you're driving it by keyboard right after a mouse Save click. |
| 10 | Find | works | Match count, Enter-to-advance all correct. |
| 11 | Replace / Replace All | works | Both applied correctly and persisted to disk on save. |
| 12 | Find & Replace — Escape closes the panel | **partly works** | Escape only closes the dialog while focus is literally inside the Find input. After using Replace/Replace All (focus moves to that button), Escape is a no-op — `FindReplace.tsx`'s `handleFindKeyDown` is wired only to the Find field's own `onKeyDown`, nothing else in the dialog. Have to click the ✕. |
| 13 | Headers/footers — plain text | works | Edited via the "Header and footer" panel, saved, and reopened correctly (run-splitting means the edited text can land across two adjacent `<w:t>`s — expected, not corruption). |
| 14 | Headers/footers — mixed content (text + PAGE field; image + text) | works | Built a header with `"Page " + PAGE field + " of report"` and a footer with an image + text. Editing the text segment kept the `PAGE` field's `instrText` intact and interleaved correctly; the footer's `<w:drawing>` and its text both survived save+reopen. |
| 15 | **Tables — insert (via Insert ▸ Table ▸ NxM picker)** | **broken** | See detail below. |
| 16 | **Tables — edit cells** | **broken** | See detail below — affects even a correctly-rendered, pre-existing (parsed-from-file) table, not just newly inserted ones. |
| 17 | **Tables — merge cells** | **broken** | Unreachable as a consequence of #16 — merge requires the cursor to register as "inside the table," which never happens. |
| 18 | Insert image | works | Native picker mocked to return a PNG; image renders on screen, saves with `<w:drawing>`, a real `pic:pic`/`a:blip` element, an embedded `word/media/image1.png`, and a correct `_rels` entry; survives reopen. |
| 19 | Insert hyperlink | works | Selected a word, mocked `window.prompt` to supply a URL, hyperlink wraps the run (`<w:hyperlink r:id=...>`) and `word/_rels/document.xml.rels` gets a correct `TargetMode="External"` relationship. |
| 20 | Spell check — toggle | works | Review tab toggle flips the surface's `spellcheck` attribute. |
| 21 | Spell check — live misspelling detection/suggestions | not tested | Depends on Electron's native spellchecker firing a real OS-dictionary context-menu event; not reliably drivable through Playwright automation in this environment. |
| 22 | Zoom in/out | works | The page element's actual layout width scales correctly with each click (confirmed via `.docx-page` bounding box, not just the `%` label). |
| 23 | Page navigation controls | not tested | Only zoom was exercised. |

### Detail: text colour writes invalid OOXML

**Repro:** select a word → Home tab → Font Color → pick the red swatch → Save → unzip → `word/document.xml`.
**Observed:** `<w:color w:val="#ff0000"/>`
**Expected:** `<w:color w:val="FF0000"/>` — ECMA-376's `ST_HexColor` is a bare 6-hex-digit value (or the literal `"auto""`); a leading `#` is not part of the type.
**Responsible code:** `src/docx/model/styles.ts` — `export const hexColor = (value) => value as HexColor` is a pure identity cast, never strips the `#` the toolbar's colour swatches supply (`Toolbar.tsx`'s `DEFAULT_COLORS` are all `#rrggbb` strings). Flows through `src/docx/editor/toolbarAdapter.ts` (`case 'set-font-color'`) into `src/docx/serializer/documentWriter.ts:1144` unchanged.
**Validator:** `validate-office-file.mjs` did **not** flag this file — it doesn't check `ST_HexColor` syntax.

### Detail: newly-created bullet/numbered lists write a non-numeric `abstractNumId`

**Repro:** place the cursor in a paragraph → Home tab → Bullet List (or Numbered List) → Save → unzip → `word/numbering.xml`.
**Observed:** `<w:abstractNum w:abstractNumId="atlas-list-2">` … `<w:num w:numId="2"><w:abstractNumId w:val="atlas-list-2"/></w:num>`
**Expected:** `w:abstractNumId`/its `w:val` must be `ST_DecimalNumber` (an integer) per the wml schema — a string like `atlas-list-2` is not a valid value there.
**Responsible code:** `src/docx/editor/insertList.ts` line 86: `const abstractNumId = \`atlas-list-${numIdStr}\`` — generates the id as a template string instead of an integer.
**Validator:** also not flagged by `validate-office-file.mjs`. This is exactly the kind of "well-formed XML, wrong schema type" defect the prompt's "unreadable content" precedent describes — I did not have real Word available to confirm the repair-dialog behaviour, but the value is unambiguously outside the type's grammar.

### Detail: inserting a table produces an invisible, unusable, data-losing table

**Repro:** place the cursor at the end of a paragraph → Insert tab → Insert Table → pick 3×3 → (try to click into a cell, or just type immediately since the command places the cursor in the first cell).
**Observed:**
- The table's DOM exists (9 `<td>`s found) but renders with **width: 0** — `<div class="docx-page__table-wrapper" style="...width: 0px;">` and every `<col style="width: 0px;">`. Nothing is visible on screen at all (screenshot confirms only the original paragraph text shows).
- Because every cell has zero width, **no cell can be clicked** (Playwright's own actionability check times out — "element is not visible").
- Typing immediately after insertion (relying on the documented "cursor lands in the first cell" behaviour) **silently discards the keystrokes** — after Save, the typed text is nowhere in the document at all, not even misplaced elsewhere.
- The saved file's `<w:tbl>` itself is well-formed (`gridCol` widths are correctly `3000`/`3000`/`3000` twips) and passes `validate-office-file.mjs` — the defect is purely at render/edit time, not in what eventually reaches disk (until the user tries to type into it, at which point their input is lost).
**Isolated root cause:** a table already present when a file is *loaded* (parsed from XML, carrying a real `tblW`) renders with correct, non-zero widths (verified directly: `.docx-page__table` bounding box `width: 601px` for a loaded table vs `0px` for a freshly-inserted one). The one structural difference is that `buildEmptyTable()` in `src/docx/editor/commands.ts` (~line 401) constructs `{ kind: 'table', tblGrid, rows }` with **no `props`/`tblW` at all**, and `src/docx/layout/layoutTable.ts`'s width resolution (~lines 56–72: `resolveWidthToPoints(table.props?.tblW ?? style?.width, availableWidthPt) ?? clampNonNegative(availableWidthPt)`) ends up resolving to 0 for that shape, even though `tblGrid` itself has perfectly good column widths to fall back on.
**Best guess at fix location:** `src/docx/editor/commands.ts` (`buildEmptyTable`) and/or `src/docx/layout/layoutTable.ts`'s width fallback.

### Detail: clicking (or tabbing) into ANY table cell does not move the editing caret there

This is the more fundamental of the two table bugs and affects tables that otherwise render perfectly (i.e. it is **not** limited to the zero-width case above).
**Repro:** open a document with a real, correctly-parsed 2-row table (e.g. from a `docx` library-built fixture) between two paragraphs → click directly on a table cell's text run.
**Observed:** `window.getSelection()` afterward reports the caret sitting in the **paragraph after the table** ("After table.", not "R1C1"), confirmed by walking the selection's parent-element chain (`docx-run → docx-page__line → docx-page__column`, i.e. the ordinary flowed body content, not the table). Typing after such a click, or after `Home`+type, lands the new characters in the body paragraph, never in the cell. Keyboard navigation (`End` then `ArrowDown`, or `End` then `Tab`, from the paragraph immediately before the table) also fails to enter the table — the keystroke is either swallowed or stays in the same paragraph.
**Consequence:** a table's cell contents cannot be edited through any input method I could drive (mouse click, Tab, arrow keys). This also explains why "Edit Table"/merge is unreachable — it's gated on the cursor registering as inside a table cell, which never happens.
**Best guess at responsible area:** the DOCX editor's click→document-position hit-testing does not account for tables, which `PageView.tsx` renders as absolutely-positioned overlays (`position: absolute`, `top`/`left`) separate from the normal paragraph flow that the hit-testing presumably walks (see `renderPageTable` in `src/docx/render/PageView.tsx`). I did not chase the exact hit-testing function itself (likely in `src/docx/editor/Input.ts` or the caret-from-point logic in `DocxViewer.tsx`) — that's the next place I'd look.

---

## Spreadsheets (.xlsx, .ods)

| # | Feature | Status | Notes |
|---|---|---|---|
| 1 | Edit cells | works | |
| 2 | Formulas — arithmetic (`=A2+B2`) | works | Correct on screen ("9") and correct cached `<v>9</v>` on disk; `SheetJS` re-read agrees. |
| 3 | **Formulas — text-returning (`=A1&"!"`, `=CONCATENATE(A1,B1)`)** | **broken** | See detail below. |
| 4 | Number formats — preserved on save | works | Edited a cell that already had a currency number-format style (`s="4"`) — the edited cell **kept its style id**; an untouched neighbouring date-formatted cell also kept its style id, after an intervening row+column insert shifted both cells' addresses. |
| 5 | Number formats / cell styles — apply new ones via UI | broken | No toolbar control for this exists anywhere in `SpreadsheetEditToolbar.tsx`/`SpreadsheetViewer.tsx` (only undo/redo, row/col insert-delete, paste, save). Confirmed by reading every control in the toolbar component; Atlas can only preserve pre-existing formatting, not let the user apply new number formats or styles. |
| 6 | Insert/delete row and column | works | Grid stats (`rows × columns`) update correctly; verified structurally via the styled-workbook fixture too (conditional-formatting range and styles.xml passed through untouched). |
| 7 | Add sheet | works | |
| 8 | Rename sheet | works | Double-click tab → rename input → Enter. |
| 9 | Delete sheet | works | |
| 10 | **Reorder sheets** | **broken** | No drag-and-drop, context menu, or move-left/right control exists anywhere on the sheet tab strip (`SpreadsheetViewer.tsx`) — confirmed by reading the component and by visual inspection of a live run (tabs have no `draggable` attribute and no reorder affordance appears). |
| 11 | Excel Tables (`xl/tables/table1.xml`) | works | Edited a cell inside a table; on save the table definition, its `displayName`, and the sheet's `<tableParts>` reference all survive; the edited value round-trips correctly. |
| 12 | ODS — edit, save, reopen | works | Edit persisted into `content.xml`; `mimetype` stayed the required first, uncompressed zip entry; reopened without error. |

### Detail: string-returning formulas are not evaluated

**Repro:** in a blank cell, type `=A1&"!"` (where A1 contains "Hello"), or `=CONCATENATE(A1,B1)`.
**Observed on screen:** the grid cell displays the **literal formula text** — `=A1&"!"` and `=CONCATENATE(...` — not the computed string.
**Observed on disk after Save:** `<c r="C1" t="str"><f>A1&amp;"!"</f><v>=A1&amp;"!"</v></c>` — the cached `<v>` is the raw formula source, not "Hello!".
**Contrast:** the exact same flow with an arithmetic formula (`=A2+B2`) computes and displays correctly both on screen and on disk (`<v>9</v>`).
**Conclusion:** Atlas's formula engine evaluates numeric expressions but does not evaluate formulas whose result is a string (concatenation via `&`, `CONCATENATE`, and by extension almost certainly other text functions like `UPPER`/`LEFT`/`TEXT`, though I only drove these two). This is reproducible and affects both the live UI and the saved file's cached value.
**Best guess at responsible area:** the spreadsheet formula evaluator (I did not trace the exact file — the fixture-generation and passthrough logic live under `src/viewers/spreadsheet/`; the evaluator itself I did not locate in the time available).

---

## Slides (.pptx, .odp)

| # | Feature | Status | Notes |
|---|---|---|---|
| 1 | Edit text in place | works | Both `.pptx` and `.odp`. |
| 2 | Move a shape | works | Both formats; `.pptx` save no longer contains the shape's original `<a:off>`, `.odp`'s bounding box shifts and content.xml round-trips. |
| 3 | Resize a shape (drag handle) | works | Bounding box changes on screen; the resized `<a:ext>` no longer matches the original `1905000x381000`. |
| 4 | **Rotate a shape** | **broken** | A shape's *existing* rotation (source file `a:xfrm rot="600000"`, i.e. 10°) renders correctly (`transform: rotate(10deg)`), and the resize math correctly accounts for rotation — but there is no UI control to let the user change a shape's rotation. `src/viewers/shared/SlideEditCanvas.tsx`'s `Handle` type (`slideGeometry.ts`) only has the 8 resize handles (`n/s/e/w/ne/nw/se/sw`); confirmed live that selecting a shape shows exactly those 8 handles and nothing else. |
| 5 | Insert a text box | works | Toolbar "Insert text box" adds a new shape; double-clicking it opens the same text editor and typed content persists to disk. |
| 6 | Add slide | works | |
| 7 | Duplicate slide | works | |
| 8 | Delete slide | works | |
| 9 | Reorder slide (move up/down) | works | Confirmed on disk: the moved slide's `r:id` now appears before the sibling's in `ppt/presentation.xml`. |
| 10 | Speaker notes | works | Persists into `ppt/notesSlides/notesSlideN.xml`. |
| 11 | Presenter view — open/counter/Next/Exit | works | Opens, shows "Slide 1 / 2", advances to "Slide 2 / 2" on Next, exits cleanly. |
| 12 | ODP — full edit/move/save/reopen round trip | works | Same coverage as `.pptx` items 1–2, run against `tests/e2e/fixtures/sample.odp`; `mimetype` stays the first, uncompressed entry. |

All of the above `.pptx` and `.odp` saves passed `validate-office-file.mjs` with no violations.

---

## Office validator gap (informational, not a matrix row)

`node scripts/validate-office-file.mjs` reported **zero violations** for every file in this sweep, including the two genuinely-invalid DOCX files described above (`w:color w:val="#ff0000"` and `w:abstractNumId="atlas-list-2"`). The validator is useful for structural OPC/part-relationship problems but does not check attribute-level type grammar (`ST_HexColor`, `ST_DecimalNumber`), so it should not be read as clearing those two defects.

## Out of scope / not attempted

Legacy `.doc`/`.ppt`, PDF, markdown, code files, and anything about the app shell/new-document flow are the other agent's area and were not touched here, per the task split.
