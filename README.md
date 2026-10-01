# Atlas

Atlas is a document viewer and editor for Windows that opens Word, Excel,
PowerPoint, PDF, OpenDocument files, plain text, code, RTF, and Markdown
through one Electron + React app, instead of a different program for each.

## Why I built it

This started as a small Markdown viewer (the repo folder is still called
`md-reader`), and it kept growing. I got tired of opening a different, often
slow-to-start program every time I just wanted to look at a file on Windows,
so I set out to build one app that opens all of them, and edits the ones
where editing actually makes sense to support.

## What it does

Every format opens through the same shell (tabs, themes, find, export), but
how much you can do with it depends on the format:

- Markdown: live preview, raw editor, split view, math (KaTeX), Mermaid
  diagrams, GFM tables. Full editing with autosave, plus export to HTML, PDF,
  DOCX, or Markdown.
- DOCX: I wrote my own parser, layout engine, and editor for this instead of
  using an existing library, so it renders the real style cascade,
  headers/footers, footnotes, tracked changes, and comments. You can edit
  text, tables, images, and hyperlinks, record and accept/reject tracked
  changes, and save back to a real `.docx`.
- PPTX and ODP: same idea, a custom slide editor (no editing library). Edit text in place,
  move and resize shapes, add or reorder slides, use speaker notes and
  presenter view, and save through the original package.
- XLSX, ODS, and legacy spreadsheets: cell editing, a small in-house formula
  evaluator, row/column/sheet operations, and saving that keeps the original
  file's styles and charts intact for `.xlsx`/`.xlsm`.
- CSV and TSV: the same grid engine as the spreadsheets, full editing.
- PDF: view-only. Text selection, search, print, thumbnails, and
  password-protected files.
- Plain text and code: code gets a real CodeMirror editor with highlighting
  for 30+ languages, find/replace, and a sandboxed Run button for
  JavaScript, TypeScript, and Python. Plain text is view-only.
- RTF and ODT: view-only, rendered and sanitized rather than converted.
- Legacy Word 97-2003 (`.doc`) and PowerPoint 97-2003 (`.ppt`): text
  extraction only, no layout or formatting.

I try to be honest about the gaps instead of hiding them:
[`docs/KNOWN_LIMITATIONS.md`](docs/KNOWN_LIMITATIONS.md) lists what each
format still can't do (things like vertical table-cell merging in DOCX, or
RTL text).

The shell itself has the things you'd expect from any editor: multiple tabs,
a New Document menu that starts from a real Office/LibreOffice-compatible
template, five themes with WCAG-checked contrast, English/French UI, recent
files, drag-and-drop, and a shortcuts modal (`Ctrl+/`).

**Version history** keeps a copy of a document each time you save it, for every
format, and the toolbar's Version history panel lists them with the option to put
any of them back. A document you're actively editing is also captured
every couple of minutes — a Word document, a spreadsheet and a slide deck as much
as a markdown file — so the history reflects the writing rather than just the
saving. Saving the same content twice doesn't make a second copy and identical
content is only stored once, so pressing Ctrl+S out of habit costs nothing.
Restoring is only offered when your current version is already saved — otherwise
an older copy could replace work that exists nowhere else. Each document keeps its
last 100 versions, and the history lives in the app's own data directory, not
beside your file.

## How it works

Opening a file goes through one pipeline regardless of how it arrived
(dialog, drag-drop, double-click from Explorer): Atlas checks the extension,
then verifies the actual file bytes against it (so a renamed `.pdf` doesn't
get treated as a `.docx`), and only then lazily loads that format's viewer
component. Each viewer is a separate code chunk, so opening a Markdown file
never pulls in the spreadsheet engine.

The part I spent the most time on is the DOCX and PPTX/ODP support. Rather
than reach for an existing conversion library, I wrote my own parsers,
in-memory document models, layout code, and editors for both, operating
directly on the OOXML/ODF XML inside the zip. That's what makes formatting
fidelity and round-trip saving possible, but it also means I own every bug
in that fidelity, which is why there's a dedicated, from-scratch structural
validator (`scripts/validate-office-file.mjs`) that checks every save
against the actual OPC/OOXML/ODF package spec rather than trusting Atlas's
own parser to agree with itself.

On the Electron side, the renderer runs sandboxed with no direct Node
access; all file I/O goes through a preload bridge and a path allowlist, and
writes are atomic so a crash mid-save can't corrupt a file. Details are in
[`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md).

## Running it

```bash
npm install
npm run electron:dev   # Electron + Vite dev server, real file open/save
npm run dev             # Vite dev server only, in a browser, no file I/O
```

Build:

```bash
npm run build           # type-check, build, and check bundle size
npm run electron:build  # same, then packages a Windows installer
```

Test:

```bash
npm test               # vitest, unit and characterization tests
npm run coverage        # vitest with coverage
npx playwright test     # Electron end-to-end smoke tests (run `npx vite build` first)
```

Lint with `npm run lint`.

## Stack

React 19 and TypeScript, built with Vite, packaged with Electron. PDF
rendering is `pdfjs-dist`; spreadsheets use SheetJS (`xlsx`) with
`@glideapps/glide-data-grid` for the grid; the code editor is CodeMirror 6.
DOCX and PPTX/ODP are my own parser/editor code on top of `jszip` and
`fast-xml-parser`. Markdown rendering is `react-markdown` with `remark-gfm`,
KaTeX for math, and Mermaid for diagrams.

## Limitations and what's next

The formula evaluator for spreadsheets is small and in-house, not a full
spreadsheet engine, so some formulas fall back to their last cached value.
Legacy `.doc`/`.ppt` files are read-only text extraction with no layout. A
handful of DOCX/PPTX edge cases (vertical cell merge, RTL text, table/chart
editing in PPTX) aren't supported yet. All of this is tracked, per format,
in [`docs/KNOWN_LIMITATIONS.md`](docs/KNOWN_LIMITATIONS.md), which I try to
keep as accurate as the code rather than aspirational.
