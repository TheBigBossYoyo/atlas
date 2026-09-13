# Atlas Phase 2 — DOCX Editing Parity (Plan)

**Repo**: `C:\Users\Youssef\Documents\Projects\md-reader`
**Goal**: Hand-coded DOCX viewer + editor with feature parity vs Microsoft Word for all common authoring tasks.
**Constraint** (user-locked): Hand-coded only. **No SDK licenses. No bundled office binaries.** OSS dependencies are limited to general-purpose libs (jszip, fast-xml-parser). NO docx-preview, NO docx.js for serialization, NO mammoth — we own the parser, the editor, and the serializer end-to-end.
**Output**: `Atlas-Setup-3.0.0.exe`. App version bumps to `3.0.0`.

## Honest Scope Statement

This phase is **conservatively estimated at 14–22 multi-task sessions** (~3–6 months of session work) before reaching usable v1 editing. Microsoft Word is ~30 years and millions of LOC. Hand-coding "everything in MS Office" verbatim is not bounded — we draw a line at **"Common Word features used by ≥95% of documents"** and ship that as v1, then iterate.

### IN SCOPE (v1 — must ship)
- **Document model**: paragraphs, runs, runs properties (rPr), paragraph properties (pPr), sections (sectPr), tables (tbl/tr/tc), images (drawing), hyperlinks, bookmarks, comments, footnotes/endnotes
- **Run formatting**: bold, italic, underline, strike, subscript/superscript, font family, font size, color, highlight color, character spacing, all caps/small caps
- **Paragraph formatting**: alignment (L/R/C/justify), line spacing, space before/after, indent (left/right/first-line/hanging), keep-with-next, page break before, bidi, shading, borders (4 sides), tab stops
- **Lists**: ordered/unordered/multi-level via numbering.xml + abstractNum + lvlOverride
- **Styles**: paragraph styles, character styles, table styles, linked styles. Full styles.xml read+write+modify+create. Style inheritance chain.
- **Tables**: insert/delete row/col, merge/split cells, cell shading, cell borders, column widths (auto/fixed/pct), header row, table style application
- **Images**: insert from file, resize (preserve ratio), wrap (inline/square/tight), alt text, embedded as `word/media/`
- **Headers/footers**: per-section, first-page-different, odd/even-different, page numbers, date fields
- **Page setup**: paper size (A4/Letter/Legal/custom), orientation, margins (4-side), columns (1-3), section breaks (next page/continuous/even/odd)
- **Find & Replace** with regex + match case + whole word
- **Spell check**: hunspell dictionary (`.dic`/`.aff`) — bundle en_US, fr_FR (matches user locale)
- **Track Changes**: insertion/deletion runs (`<w:ins>`/`<w:del>`), accept/reject single & all
- **Comments**: insert, reply, resolve, delete (`comments.xml` + `commentsExtended.xml`)
- **Undo/Redo**: command pattern with bounded history (200 ops)
- **Paste from Word/web**: HTML→DOCX-AST converter that preserves common formatting
- **Export**: Save as DOCX (faithful), PDF (existing pipeline), HTML, Markdown (lossy)

### OUT OF SCOPE for Phase 2 (deferred to Phase 2.x patches)
- Equations (OMML) — VIEWER-only via MathML conversion in v1; editing in 2.1
- SmartArt — render as image placeholder in v1; editing in 2.2
- Charts (`chart1.xml`) — render as embedded image fallback; editing in 2.3
- Macros (`vbaProject.bin`) — preserved on roundtrip but never executed
- Mail merge fields beyond basic `MERGEFIELD` rendering
- Real-time collaboration (no CRDT in Phase 2)
- DRM / Information Rights Management
- Forms (legacy `FORMTEXT` etc.) — preserved on roundtrip, no editing UI
- Ink annotations
- 3D models embedded (`a:graphicData uri="3DModel"`) — placeholder image

### Locked Library Stack (Phase 2)
- `jszip` — already installed (Phase 1)
- `fast-xml-parser` — NEW. Permissive license (MIT). XML parse + build with attribute order preservation. We do NOT use docx-preview or any wrapper.
- `nspell` — NEW. MIT. Hunspell-compatible spellchecker in pure JS.
- Hunspell dictionaries: `dictionary-en` and `dictionary-fr` packages (BSD/LGPL data files — permitted; data not code).
- NO `docx`, NO `docx-preview`, NO `mammoth`. The existing `docx` dep used for export from markdown stays for the markdown→docx path only and is replaced in W2.E1 by our own serializer.

### Architecture Contracts (locked)
- New tree:
  ```
  src/docx/
    model/           # In-memory AST (TS types only, no behavior)
      document.ts    # Document, Section, Paragraph, Run, Table, ...
      styles.ts      # Style, ParaProps, RunProps, NumDef, ...
      ids.ts         # rId allocator, bookmark id allocator
    parser/          # OOXML → AST
      unzip.ts       # JSZip wrapper, returns Map<path, Uint8Array>
      relationships.ts  # _rels/*.rels parser
      contentTypes.ts   # [Content_Types].xml
      document.ts    # word/document.xml parser
      styles.ts      # word/styles.xml
      numbering.ts   # word/numbering.xml
      theme.ts       # word/theme/theme1.xml (font/color schemes)
      headerFooter.ts
      comments.ts
      footnotes.ts
      media.ts       # word/media/*
    serializer/      # AST → OOXML
      zip.ts         # JSZip writer with deterministic order
      relationships.ts
      document.ts
      styles.ts
      numbering.ts
      headerFooter.ts
      comments.ts
      footnotes.ts
    editor/          # Editing surface
      DocxEditor.tsx           # Main editable component
      Selection.ts             # Selection/range model (anchor + focus)
      Commands.ts              # Insert/Delete/Format command objects
      History.ts               # Undo/redo stack
      Toolbar.tsx              # Word-style ribbon (lite)
      paste/
        htmlToAst.ts           # HTML clipboard → AST fragment
        wordHtmlSanitizer.ts   # Strip Word's bloat
    layout/          # Print-accurate layout engine
      Paginator.ts             # Page break computation
      LineBreaker.ts           # Greedy line breaker w/ TeX-like fallback
      TableLayout.ts
      Cursor.ts                # Pixel-coord ↔ (paragraph, offset) mapping
    render/          # Layout → DOM
      PageView.tsx             # One A4/Letter page
      Renderer.tsx             # Maps AST + layout to React tree
    fonts/
      registry.ts              # Bundled font list + Web Font Loader fallback
      metrics.ts               # Cached glyph metrics for layout
    spell/
      hunspell.ts              # nspell wrapper
      worker.ts                # Web Worker for spellcheck (off main thread)
  ```
- Discriminated unions throughout. `assertNever` pattern from Phase 1 enforced.
- The viewer (`src/viewers/DocxViewer.tsx`) is now a THIN wrapper around the editor in read-only mode (`<DocxEditor readOnly />`). The Phase 1 plan's W2.DOCX (docx-preview based) is REPLACED by Phase 2's hand-coded path.
- All viewer libs lazy-loaded — `src/docx/**` is one lazy chunk.

---

## Wave A — Foundation (Parser + Model + Read-Only Render)

### A.1 — DOCX file model & unzip
- **Description**: Implement `src/docx/parser/unzip.ts` returning `Map<string, Uint8Array>`. Implement `relationships.ts` parsing `_rels/.rels` and `word/_rels/document.xml.rels`. Implement `contentTypes.ts` parsing `[Content_Types].xml`. Add round-trip test: read DOCX → write back via JSZip → verify identical file list.
- **Files**: `src/docx/parser/{unzip,relationships,contentTypes}.ts`, tests.
- **Agent**: Category `unspecified-high`, Skills `[]`.
- **Pass**: 5 sample DOCX files unzip + relist identically.

### A.2 — AST type system
- **Description**: `src/docx/model/document.ts` defines `Document`, `Section`, `Paragraph`, `Run`, `Table`, `TableRow`, `TableCell`, `Drawing`, `Hyperlink`, `Bookmark`, `Comment`, `Footnote`. `model/styles.ts` defines `Style`, `ParaProps`, `RunProps`, `NumberingDef`, `LvlDef`. ALL `readonly`, all named, all immutable. `assertNever` for kind switches.
- **Files**: `src/docx/model/*`, type tests.
- **Agent**: Category `ultrabrain`, Skills `[]`. (Type design must be airtight; later phases depend on it.)
- **Pass**: tsc clean; 100% of OOXML run/para attributes representable.

### A.3 — Document parser (paragraphs, runs, tables)
- **Description**: `parser/document.ts` parses `word/document.xml` into AST. Handle `<w:p>`, `<w:r>`, `<w:t>`, `<w:tab>`, `<w:br>`, `<w:tbl>`, `<w:tr>`, `<w:tc>`, `<w:hyperlink>`, `<w:bookmarkStart>`, `<w:bookmarkEnd>`, `<w:sectPr>`. Preserve unknown elements as `UnknownNode { xml: string }` for lossless roundtrip.
- **Agent**: Category `ultrabrain`, Skills `[]`.
- **Pass**: Parse 20 real-world DOCX files; assert no `UnknownNode` count >5% of total nodes.

### A.4 — Styles + numbering parser
- **Description**: `parser/{styles,numbering}.ts`. Build style inheritance graph (basedOn / linkedTo). Resolve effective rPr/pPr per run/para via cascade: doc default → style → direct. Parse abstractNum + num + lvlOverride for lists.
- **Agent**: Category `ultrabrain`, Skills `[]`.
- **Pass**: Cascade test — 5 documents with deep style chains resolve to expected effective props.

### A.5 — Headers, footers, footnotes, comments, theme parsers
- **Agent**: Category `unspecified-high`, Skills `[]`.

### A.6 — Renderer (read-only DOM)
- **Description**: `render/Renderer.tsx` maps AST → React tree. One `<section>` per Section, `<p>` per Paragraph with computed inline style from effective pPr, `<span>` per Run with rPr style. Tables → `<table>`. Hyperlinks → `<a>`. Page boundaries via CSS `break-before: page`. NOT yet the print-accurate paginator (that's W B.x).
- **Agent**: Category `visual-engineering`, Skills `[frontend-ui-ux]`.
- **Pass**: Visual diff vs MS Word screenshot of 5 fixtures — heading sizes, bold/italic, list indents, table borders all visible.

### A.7 — DocxViewer rewires to hand-coded path
- **Description**: Replace Phase 1 W2.DOCX docx-preview implementation with `<HandRolledDocxViewer file={file} />` using `src/docx/parser` + `src/docx/render`. Keep same NavItem (heading outline) + ViewerStats publishing contract.
- **Agent**: Category `unspecified-high`, Skills `[]`.

### Gate GA — Parser correctness
- `npm test -- docx` covers parser cases.
- 20-fixture roundtrip script: parse → serialize → re-parse → AST equality (ignoring whitespace and unknown-element ordering).
- Pass: 18/20 byte-level identical, all 20 AST-equal.

---

## Wave B — Layout & Pagination

### B.1 — Font metrics
- **Description**: Bundle 4 font families (Calibri-substitute = Carlito, Cambria-substitute = Caladea, Courier New-substitute = Courier Prime, Times New Roman-substitute = Tinos — all Apache-licensed Google Crosextra fonts). Build `fonts/metrics.ts` extracting per-glyph advance widths from TTF tables (cmap, hmtx, head). Cache in IndexedDB.
- **Agent**: Category `deep`, Skills `[]`.

### B.2 — Line breaker
- **Description**: Greedy line breaker matching Word's algorithm (no Knuth-Plass for v1). Inputs: run array with widths, available width. Output: line array. Handle hyphens, no-break spaces, tab stops. Justify by stretching inter-word spacing.
- **Agent**: Category `ultrabrain`, Skills `[]`.

### B.3 — Paginator
- **Description**: Stack lines into pages respecting page size, margins, headers/footers, columns. Handle keep-with-next, page-break-before, widow/orphan control. Output: `Page[]` with `LineBox[]` and `(paragraph, runIndex, charOffset)` ranges.
- **Agent**: Category `ultrabrain`, Skills `[]`.

### B.4 — TableLayout
- **Description**: Compute column widths (auto/fixed/pct), row heights based on cell content reflow, handle merged cells, repeat-header-rows across pages.
- **Agent**: Category `ultrabrain`, Skills `[]`.

### B.5 — PageView render
- **Description**: `render/PageView.tsx` renders a single `Page` as a fixed-size DOM sheet (e.g., 210×297mm at zoom factor). Print preview by stacking pages. Replaces the W A.6 flow renderer for print mode.
- **Agent**: Category `visual-engineering`, Skills `[frontend-ui-ux]`.

### Gate GB — Layout fidelity
- 10 fixtures with known page counts (computed in MS Word) — Atlas's paginator matches page count exactly on 9/10.

---

## Wave C — Editor Surface

### C.1 — Selection model
- **Description**: `editor/Selection.ts`. `Position = { paragraphPath: number[], runIndex: number, charOffset: number }`. Range = anchor + focus. Bidirectional. Hooks: `useSelection()`. Mouse handlers: click/drag/double-click (word) /triple-click (paragraph) → Position via `Cursor.ts` reverse mapping.
- **Agent**: Category `ultrabrain`, Skills `[]`.

### C.2 — Command system + History
- **Description**: `Commands.ts` defines `Command<T>` discriminated union: `InsertText`, `DeleteRange`, `InsertParagraphBreak`, `ApplyRunFormat`, `ApplyParaFormat`, `InsertTable`, `InsertImage`, `InsertHyperlink`, `MergeCells`, `SplitCells`, `InsertSection`, `ApplyStyle`, `InsertList`, `ChangeListLevel`, etc. (~30 commands). Each has `apply(doc): doc'` and `invert(doc): Command`. `History.ts` keeps stacks; coalesces consecutive `InsertText`.
- **Agent**: Category `ultrabrain`, Skills `[]`.

### C.3 — Keyboard input
- **Description**: `editor/Input.ts` translates DOM events to Commands. Map: typing→InsertText, Backspace/Delete→DeleteRange, Enter→InsertParagraphBreak, Shift+Enter→soft break, Tab→increase list level / next cell in table, arrow keys→Selection moves, Ctrl+B/I/U→ApplyRunFormat, Ctrl+Z/Y→Undo/Redo, Ctrl+X/C/V→cut/copy/paste.
- **Agent**: Category `unspecified-high`, Skills `[]`.

### C.4 — IME / contentEditable handling
- **Description**: For accents (é è ç) and CJK input, hook `compositionstart`/`compositionupdate`/`compositionend`. Avoid double-input bugs. Use `beforeinput` event where possible.
- **Agent**: Category `deep`, Skills `[]`.

### C.5 — Toolbar (Word-lite ribbon)
- **Description**: `editor/Toolbar.tsx` — Home tab equivalent: font picker, size picker, bold/italic/underline/strike/sub/sup, font color, highlight color, alignment 4-way, line spacing, bullet/number list, indent±, styles dropdown. Insert tab: table, image, hyperlink, header/footer, page-break, comment. Layout tab: margins, orientation, size, columns. Review tab: spell-check toggle, comments pane, track-changes toggle, accept/reject.
- **Agent**: Category `visual-engineering`, Skills `[frontend-ui-ux]`.

### C.6 — Find & Replace
- **Description**: Walk AST runs, regex match, highlight matches, "Replace" issues `DeleteRange` + `InsertText` commands.
- **Agent**: Category `unspecified-high`, Skills `[]`.

### Gate GC — Editor smoke
- Playwright-Electron script types "Hello World", bolds it, applies Heading 1, inserts table, inserts image, hits Ctrl+S → resulting DOCX opens cleanly in MS Word and shows the same content.

---

## Wave D — Serializer

### D.1 — Document.xml writer
- **Description**: AST → XML via fast-xml-parser builder with attribute-order map matching Word's output (so diffs against original are minimal). Preserve `UnknownNode.xml` verbatim.
- **Agent**: Category `ultrabrain`, Skills `[]`.

### D.2 — Styles, numbering, headers, footers, footnotes, comments writers
- **Agent**: Category `ultrabrain`, Skills `[]`.

### D.3 — Relationships + ContentTypes writers
- **Description**: Allocate new rIds for inserted images/hyperlinks/comments. Update `[Content_Types].xml` overrides for new media types.
- **Agent**: Category `unspecified-high`, Skills `[]`.

### D.4 — Zip packager
- **Description**: Deterministic file order matching MS Word's order (helps diff). Compression level matching Word's default.
- **Agent**: Category `quick`, Skills `[]`.

### Gate GD — Roundtrip suite
- For each of 30 fixture DOCX files: open in Atlas, save without edits, open saved file in MS Word — MUST open without "Word found unreadable content" dialog. Manual user QA.
- For 10 fixtures with "made one bold edit" scenario: result opens in MS Word and the edit shows correctly.

---

## Wave E — Advanced Features

### E.1 — Track Changes
- **Description**: Wrap edits in `<w:ins author=... date=... id=...>` / `<w:del>` when track-changes mode is on. Render with strikethrough/underline + author color. Accept/reject = remove wrapper / drop content.
- **Agent**: Category `ultrabrain`, Skills `[]`.

### E.2 — Comments
- **Description**: `commentRangeStart`/`commentRangeEnd`/`commentReference` insertion. Comments pane (right rail) showing thread, replies, resolve.
- **Agent**: Category `visual-engineering`, Skills `[frontend-ui-ux]`.

### E.3 — Spell check (hunspell + worker)
- **Description**: `spell/worker.ts` runs nspell off main thread. Underline misspellings (red squiggle) via overlay div. Right-click → suggestions → ApplyText command.
- **Agent**: Category `deep`, Skills `[]`.

### E.4 — Image insertion + resize handles
- **Agent**: Category `visual-engineering`, Skills `[frontend-ui-ux]`.

### E.5 — Paste from Word/HTML
- **Description**: Clipboard `text/html` → strip Word's `xmlns:o`, `mso-*` styles, `<o:p>` tags. Convert into AST fragment via `htmlToAst.ts` mapping (h1→Heading1 style, b→bold run, table→Table, etc.). Insert via command.
- **Agent**: Category `ultrabrain`, Skills `[]`.

### Gate GE — Real-world DOCX battery
- Open, edit one paragraph in each of 50 real DOCX files (mix from user-provided corpus + curated GitHub OOXML test suite). All 50 save and reopen in Word without errors. Bold/italic/list/table edits visible and correct.

---

## Wave F — Polish + Ship

### F.1 — Print preview + print
- **Agent**: Category `unspecified-high`, Skills `[]`.

### F.2 — Performance pass (large docs)
- 1000-page document opens in <5s, scrolls at 60fps.
- Virtualize page list (only render ±2 pages around viewport).
- Web Worker for parsing on big files.
- **Agent**: Category `deep`, Skills `[]`.

### F.3 — Accessibility (a11y)
- ARIA roles for editor (`role="textbox" aria-multiline="true"`), Toolbar (`role="toolbar"`), comment pane.
- Screen reader announces formatting changes.
- **Agent**: Category `unspecified-high`, Skills `[]`.

### F.4 — Keyboard shortcuts parity
- Map all Word shortcuts: Ctrl+B/I/U, Ctrl+1/2/5 (line spacing), Ctrl+L/E/R/J (alignment), Ctrl+Shift+L (bullet list), Ctrl+K (hyperlink), F7 (spellcheck), Ctrl+F (find), Ctrl+H (replace), Ctrl+Shift+S (style), etc.
- **Agent**: Category `quick`, Skills `[]`.

### F.5 — Build & ship
- Bump `package.json` version → `3.0.0`.
- `Atlas-Setup-3.0.0.exe`.
- Manual upgrade test: install over 2.0.0, edit a docx, save, reopen in Word.
- **Agent**: Category `quick`, Skills `[]`.

---

# Success Criteria (Phase 2 Done)

1. `release/Atlas-Setup-3.0.0.exe` exists, installs over Atlas 2.0.0.
2. Open ANY `.docx` from Microsoft test suite + 50-doc real-world corpus → renders without crashes.
3. Edit operations supported: type/delete text, format runs (B/I/U/strike/sub/sup/font/size/color/highlight), format paragraphs (align/spacing/indent/borders/shading), insert/delete tables + rows + cols, merge/split cells, insert/resize images, insert hyperlinks, insert page/section breaks, change page setup, insert/edit headers + footers, lists at multiple levels, comments insert+reply+resolve, track changes on/off + accept/reject.
4. Save → resulting DOCX opens in Microsoft Word 2019+ without warnings on ≥45/50 real-world fixtures.
5. Spell check working in en + fr.
6. Find & Replace with regex working.
7. Undo/Redo working across 200+ ops.
8. Print preview accurate (page count matches Word ±0 on 9/10 fixtures).
9. lint, tsc, test, build all green.
10. Main bundle did not regress >40% vs Phase 1 (DOCX module is lazy-loaded).

# Atomic Commit Strategy
One commit per task ID. Conventional Commits. After each Wave Gate, run `skill(name="review-work")` against cumulative diff.

# Risk Register
- **R1 — OOXML edge cases**: real-world DOCX has malformed XML, undocumented attributes, vendor extensions (LibreOffice/Pages). Mitigation: `UnknownNode` fallback preserves content; expand parser via fixture-driven TDD.
- **R2 — Layout fidelity**: matching Word's exact line breaks requires reverse-engineering its algorithm. Mitigation: aim for "visually plausible" not "byte-identical pagination". Document the gap.
- **R3 — Font availability**: Calibri/Cambria are Microsoft proprietary. Mitigation: bundle metric-compatible substitutes (Carlito/Caladea); on Windows, attempt to load system Calibri first.
- **R4 — IME bugs**: text input across locales is hard. Mitigation: dedicated W C.4 task; QA with French + CJK input by user.
- **R5 — Performance regression**: large docs may freeze. Mitigation: Web Worker parsing + page virtualization in W F.2.
- **R6 — Microsoft Word "unreadable content" warning**: serialization round-trip is the highest-risk gate (D + E). Mitigation: 50-fixture battery, AST-equality property test, attribute-order preservation.

# Phase 3+ Preview (NOT in scope here)
- Phase 3: XLSX editing + formula engine (~9–18 months solo work)
- Phase 4: PPTX editing
- Phase 5: PDF editing
- Phase 6: ODT/ODS/ODP editing
- Phase 7: Real-time collaboration (CRDT)
