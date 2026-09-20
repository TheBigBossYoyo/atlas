# Atlas

A universal document viewer & editor for Windows, built on Electron + React.
Opens Word, Excel, PowerPoint, PDF, OpenDocument, plain text/code, RTF, and
Markdown — 12 binary formats plus Markdown, all through one app.

---

## Features by format

| Format | View | Edit + Save | Export |
|---|---|---|---|
| **Markdown** (`.md`) | Live preview, raw editor, split view, math (KaTeX), Mermaid diagrams, GFM tables/task lists, auto TOC, in-document find | Yes — autosave draft + Save/Save As | HTML, PDF (vector, via Chromium's `printToPDF`), DOCX, Markdown |
| **DOCX** (`.docx`) | Full layout fidelity (style cascade, pagination, headers/footers/footnotes/endnotes, tab stops, section vertical alignment, lists, hyperlinks, anchored/floating images with crop/rotation/flip, embedded fonts, fields/TOC, track changes, comments) | Yes — text editing, formatting, lists/tables (insert/delete row+column, horizontal merge/split, resize)/hyperlinks/images/page breaks, rich paste, undo/redo, track-changes recording (`w:ins`/`w:del`) + accept/reject, Update Field(s)/Update TOC; Save/Save As. *Not yet: vertical table cell merge, RTL, text wrap around a floating image, TOC regeneration on a multi-paragraph TOC field, formatting-change tracking — see [`docs/KNOWN_LIMITATIONS.md`](docs/KNOWN_LIMITATIONS.md)* | PDF (real per-format export, not a screenshot) |
| **PDF** (`.pdf`) | Virtualized page rendering, text layer (select/copy), find, print, page rotation, thumbnails, password-protected files, link/form-field annotations | View-only | Save a copy |
| **PPTX** (`.pptx`) / **ODP** (`.odp`) | Slide layout + master inheritance, run formatting, bullets, tables, grouped shapes, speaker notes, keyboard navigation, thumbnail rail | Yes — edit text in place, move/resize shapes (including rotated ones), insert/delete text boxes, add/duplicate/delete/reorder slides, speaker notes, presenter view; undo/redo; Save/Save As through the original package. *Not yet: tables, charts, images (movable, not replaceable), animations/transitions — see [`docs/KNOWN_LIMITATIONS.md`](docs/KNOWN_LIMITATIONS.md)* | PDF (one page per slide at deck aspect ratio, not a screenshot) |
| **XLSX/ODS/legacy spreadsheets** (`.xlsx`, `.ods`, `.xls`, `.xlsb`, `.fods`) | Formatted values, merged cells, column widths, hidden sheets, frozen panes (XLSX/ODS only); parsed off the main thread | Cell edit, formulas (small in-house evaluator), insert/delete rows/columns, add/rename/delete sheets, undo/redo, copy/paste; `.xlsx`/`.xlsm` save through the *original* package (styles, number formats, charts, filters, tables) with formula references and defined names automatically re-anchored across every structural edit above; Save/Save As to `.xlsx`/`.xlsm`/`.xlsb`/`.xls`/`.ods`/`.fods` (`.xls`/`.xlsb` always Save-As on first save, never silently re-encoded) | PDF (real table export) + CSV per visible sheet |
| **CSV / TSV** (`.csv`, `.tsv`) | Grid view via the same spreadsheet engine | Cell edit, insert/delete rows/columns, undo/redo, copy/paste; Save/Save As | PDF (real export) or a faithful CSV/TSV re-export |
| **Plain text** (`.txt` and 40+ others) | Virtualized for large files, in-document find | View-only | PDF or HTML (full raw content, not a screenshot) |
| **Code** | CodeMirror 6 rendering (30+ languages, syntax highlighting derived from Shiki), virtualized, in-document find | Yes — full text editing, find/replace, go to line, folding, multi-cursor, save; an explicit, confirmed, sandboxed **Run** for JavaScript/TypeScript/Python | PDF or HTML (full raw content, not a screenshot) |
| **RTF** (`.rtf`) | Rendered via `rtf.js` (including embedded WMF/EMF images), sanitized with DOMPurify, in-document find | View-only | PDF (real export, not a screenshot) |
| **ODT** (`.odt`) | Rendered via `odf-kit`, sanitized with DOMPurify, in-document find | View-only | PDF (real export, not a screenshot) |
| Legacy **Word 97-2003** (`.doc`) / **PowerPoint 97-2003** (`.ppt`) | Read-only text extraction (OLE2/CFB container parsing, `src/legacy/`) — the document's text / each slide's text, no layout/images/formatting | Not editable | — |

See [`docs/KNOWN_LIMITATIONS.md`](docs/KNOWN_LIMITATIONS.md) for the honest,
detailed gap list per format, and
[`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) for how the file-open →
detect → route → viewer pipeline and the editing/save contract work.

### Shell features

- Several documents open at once in tabs (`Ctrl+Tab`, `Ctrl+W`,
  `Ctrl+Shift+T` to reopen the last closed one, middle-click to close,
  drag to reorder); only the visible tab is mounted, and a background tab
  re-reads its file from disk when shown again
- A **New** menu (toolbar + `Ctrl+N`) creates a blank Markdown/DOCX/XLSX/
  ODS/PPTX/ODP document from a real, Office/LibreOffice-compatible
  template; opening a genuinely 0-byte file of one of those formats opens
  that same blank template instead of a parse error
- Five built-in themes (Light, Dark, Sepia, Nord, Dracula), persisted across
  sessions, cycled with `Ctrl+T`; every theme's foreground/background pairs
  pass WCAG AA contrast (enforced by a unit test that reads the real CSS
  tokens)
- English/French UI language (menu next to the theme picker, or follow the
  OS locale), persisted like the theme — see
  [`docs/KNOWN_LIMITATIONS.md`](docs/KNOWN_LIMITATIONS.md) for what is not
  yet translated
- Recent files (welcome screen), drag-and-drop opening (including several
  files at once)
- Adjustable font size, status bar with word/character/line counts and
  reading time
- Close-confirmation prompt for unsaved changes; draft recovery
- Shortcuts modal (`Ctrl+/`); every dialog/overlay (Shortcuts, unsaved-
  changes prompt, in-document find, PDF password prompt, presenter view)
  traps `Tab` inside itself and restores focus to whatever opened it on
  close

---

## Keyboard shortcuts

The authoritative list is the in-app `ShortcutsModal` (`Ctrl+/`,
`src/components/ShortcutsModal.tsx`) — the table below mirrors it exactly.

| Shortcut | Action |
|---|---|
| `Ctrl+N` | New document menu |
| `Ctrl+O` | Open file |
| `Ctrl+S` | Save |
| `Ctrl+Shift+S` | Save As |
| `Ctrl+W` | Close file |
| `Ctrl+Tab` / `Ctrl+Shift+Tab` | Next / previous tab |
| `Ctrl+P` | Print / export |
| `Ctrl+E` | Export menu |
| `Ctrl+1` / `Ctrl+2` / `Ctrl+3` | Preview / Split / Editor view (Markdown only) |
| `Ctrl+B` | Toggle sidebar |
| `Ctrl+T` | Cycle theme |
| `Ctrl+Shift+T` | Reopen last closed document |
| `Ctrl+=` / `Ctrl+-` / `Ctrl+0` | Increase / decrease / reset font size |
| `Ctrl+F` | Find in document (Markdown, Text, Code, RTF, ODT) |
| `Enter` / `Shift+Enter` | Next / previous find match |
| `Esc` | Close the open dialog/menu |
| `Ctrl+/` | Toggle this shortcuts dialog |

While editing a DOCX, its own editor shortcuts (bold/italic/underline,
alignment, find/replace, line spacing, undo/redo, save, print) take priority
over the shell shortcuts above with the same key — see
[`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md#3-shortcut-dispatcher-precedence)
for exactly how that precedence works.

---

## Development

```bash
npm install
npm run dev            # Vite dev server (browser — no file save/open)
npm run electron:dev   # Electron + Vite dev server (preferred: real file I/O)
```

See [`CONTRIBUTING.md`](CONTRIBUTING.md) for the environment rules that
matter most when working in this repo (test-runner worker limits, the
worktree `node_modules` junction, the bundle-size gate, and more).

## Build

```bash
npm run build            # prebuild (extension manifest) + tsc -b + vite build
                          # + postbuild bundle-regression check
npm run electron:build   # same, then packages a Windows installer
npm run electron:preview # build + launch the packaged-shape app (no dev server)
```

## Test

```bash
npm test              # vitest run — unit + characterization + corpus suites
npm run test:watch    # vitest watch mode
npm run coverage       # vitest run --coverage (v8; see vitest.config.ts for
                        # the enforced threshold floor and how to raise it)
npx playwright test    # Electron end-to-end smoke suite (run `npx vite build`
                        # first; workers: 1 — Electron holds a single-instance
                        # lock, so e2e specs can't run in parallel)
```

## Lint

```bash
npm run lint
```

---

## Architecture

See [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) for the file-open →
detect → route → viewer pipeline, the `ViewerContext` document-session
contract (`setDirty`/`registerSave`/`getExportableContent`), the shortcut
dispatcher's precedence rules, the Electron IPC trust boundary (preload
bridge, path allowlist, CSP, atomic writes), the hand-written DOCX engine
pipeline, and the testing strategy.

## Stack

- **React 19** + **TypeScript 6**, built with **Vite 8**
- **Electron 44** for desktop file I/O (sandboxed renderer, no Node
  integration — see `docs/ARCHITECTURE.md`)
- **DOCX**: a hand-written parser/model/layout/editor/serializer under
  `src/docx/` (no `docx-preview`, no SDK) — `jszip` + `fast-xml-parser` for
  the OOXML plumbing; the `docx` package remains only for the
  markdown→DOCX export path
- **PDF**: `pdfjs-dist`
- **Spreadsheets/CSV**: `xlsx` (SheetJS) + `papaparse`, rendered with
  `@glideapps/glide-data-grid`
- **PPTX/ODP**: hand-written parsers + editors under `src/viewers/slides/`
  (direct OPC/ODF XML surgery — no PptxGenJS/SDK — sharing an edit-queue/
  undo-redo core between the two formats)
- **ODT**: `odf-kit`
- **RTF**: `rtf.js`
- **Code editor**: CodeMirror 6, syntax highlighting derived from `shiki`'s
  language definitions; sandboxed **Run** for JS/TS/Python via a
  child-process helper in `electron/lib/codeRunner.cjs`
- **Markdown**: `react-markdown` + `remark-gfm`/`remark-math`, `rehype-katex`
  for math, `mermaid` for diagrams
- **Export**: `docx` (DOCX), `jspdf` + `html2canvas-pro` (PDF/screenshot
  export), `dompurify` for sanitizing library-rendered HTML (RTF/ODT)
- **i18n**: a small, dependency-free English/French message catalogue
  (`src/i18n/`) — no `i18next`/`react-intl`
- Virtualized rendering (`react-window`) for large text/code/spreadsheet/
  DOCX-page content
- **Office/ODF package validation**: `scripts/validate-office-file.mjs`, an
  independent structural validator (own from-scratch PKZIP reader, not
  JSZip) checking every save path's output against the actual OPC/OOXML/ODF
  package specs rather than Atlas's own parsers — see
  `docs/ARCHITECTURE.md`'s testing section

---

## Docs

- [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) — shared-shell design
- [`docs/KNOWN_LIMITATIONS.md`](docs/KNOWN_LIMITATIONS.md) — per-format gaps
- [`docs/RELEASE.md`](docs/RELEASE.md) — release checklist, installer
  metadata, code-signing notes
- [`CONTRIBUTING.md`](CONTRIBUTING.md) — environment rules for working in
  this repo
- [`CHANGELOG.md`](CHANGELOG.md) — what shipped, by wave
- [`.sisyphus/plans/`](.sisyphus/plans/) — the improvement program's plans,
  findings register, and execution status
