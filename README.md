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
| **PPTX** (`.pptx`) / **ODP** (`.odp`) | Slide layout + master inheritance, run formatting, bullets, tables, grouped shapes, speaker notes, keyboard navigation, thumbnail rail | View-only | PDF (one page per slide at deck aspect ratio, not a screenshot) |
| **XLSX/ODS/legacy spreadsheets** (`.xlsx`, `.ods`, `.xls`, `.xlsb`, `.fods`) | Formatted values, merged cells, column widths, hidden sheets, frozen panes (XLSX/ODS only); parsed off the main thread | Cell edit, formulas (small in-house evaluator), insert/delete rows/columns, add/rename/delete sheets, undo/redo, copy/paste; Save/Save As to `.xlsx`/`.xlsm`/`.xlsb`/`.xls`/`.ods`/`.fods` (`.xls`/`.xlsb` always Save-As on first save, never silently re-encoded) | PDF (real table export) + CSV per visible sheet |
| **CSV / TSV** (`.csv`, `.tsv`) | Grid view via the same spreadsheet engine | Cell edit, insert/delete rows/columns, undo/redo, copy/paste; Save/Save As | PDF (real export) or a faithful CSV/TSV re-export |
| **Plain text** (`.txt` and 40+ others) | Virtualized for large files, in-document find | View-only | PDF or HTML (full raw content, not a screenshot) |
| **Code** | Shiki syntax highlighting (30+ languages), virtualized, in-document find | View-only | PDF or HTML (full raw content, not a screenshot) |
| **RTF** (`.rtf`) | Rendered via `rtf.js` (including embedded WMF/EMF images), sanitized with DOMPurify, in-document find | View-only | PDF (real export, not a screenshot) |
| **ODT** (`.odt`) | Rendered via `odf-kit`, sanitized with DOMPurify, in-document find | View-only | PDF (real export, not a screenshot) |
| Legacy `.doc`/`.xls`/`.ppt` | Detected (CFB magic bytes) and given a friendly "open the modern equivalent" message | Not parsed | — |

See [`docs/KNOWN_LIMITATIONS.md`](docs/KNOWN_LIMITATIONS.md) for the honest,
detailed gap list per format, and
[`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) for how the file-open →
detect → route → viewer pipeline and the editing/save contract work.

### Shell features

- Five built-in themes (Light, Dark, Sepia, Nord, Dracula), persisted across
  sessions, cycled with `Ctrl+T`
- Recent files (welcome screen), drag-and-drop opening
- Adjustable font size, status bar with word/character/line counts and
  reading time
- Close-confirmation prompt for unsaved changes; draft recovery
- Shortcuts modal (`Ctrl+/`)

---

## Keyboard shortcuts

The authoritative list is the in-app `ShortcutsModal` (`Ctrl+/`,
`src/components/ShortcutsModal.tsx`) — the table below mirrors it exactly.

| Shortcut | Action |
|---|---|
| `Ctrl+O` | Open file |
| `Ctrl+S` | Save |
| `Ctrl+Shift+S` | Save As |
| `Ctrl+W` | Close file |
| `Ctrl+P` | Print / export |
| `Ctrl+E` | Export menu |
| `Ctrl+1` / `Ctrl+2` / `Ctrl+3` | Preview / Split / Editor view (Markdown only) |
| `Ctrl+B` | Toggle sidebar |
| `Ctrl+T` | Cycle theme |
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
- **Electron 35** for desktop file I/O (sandboxed renderer, no Node
  integration — see `docs/ARCHITECTURE.md`)
- **DOCX**: a hand-written parser/model/layout/editor/serializer under
  `src/docx/` (no `docx-preview`, no SDK) — `jszip` + `fast-xml-parser` for
  the OOXML plumbing; the `docx` package remains only for the
  markdown→DOCX export path
- **PDF**: `pdfjs-dist`
- **Spreadsheets/CSV**: `xlsx` (SheetJS) + `papaparse`, rendered with
  `@glideapps/glide-data-grid`
- **PPTX/ODP**: hand-written parsers under `src/viewers/slides/`
- **ODT**: `odf-kit`
- **RTF**: `rtf.js`
- **Code highlighting**: `shiki`
- **Markdown**: `react-markdown` + `remark-gfm`/`remark-math`, `rehype-katex`
  for math, `mermaid` for diagrams
- **Export**: `docx` (DOCX), `jspdf` + `html2canvas-pro` (PDF/screenshot
  export), `dompurify` for sanitizing library-rendered HTML (RTF/ODT)
- Virtualized rendering (`react-window`) for large text/code/spreadsheet
  content

---

## Docs

- [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) — shared-shell design
- [`docs/KNOWN_LIMITATIONS.md`](docs/KNOWN_LIMITATIONS.md) — per-format gaps
- [`CHANGELOG.md`](CHANGELOG.md) — what shipped, by wave
- [`.sisyphus/plans/`](.sisyphus/plans/) — the improvement program's plans,
  findings register, and execution status
