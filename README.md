# Atlas

A universal document viewer & editor for Windows. Opens Word, Excel, PowerPoint, PDF, markdown, code, and more.

Open files, edit them live, switch between five themes, render math and Mermaid diagrams, and export to PDF, DOCX, HTML, or Markdown.

---

## Features

### Reading & Editing
- **Live preview** with GitHub-flavored Markdown
- **Editor** view for raw markdown
- **Split** view — editor and preview side-by-side
- **Autosave** drafts to local storage
- **Save** / **Save As** to disk (Electron) or download (browser)
- **Recent files** quick access from the welcome screen

### Rendering
- Syntax highlighting for all major languages
- Math / LaTeX (KaTeX) — inline and block
- Mermaid diagrams (flowcharts, sequence, gantt, etc.)
- Auto-generated table of contents
- In-document search (`Ctrl+F`)

### Themes
Five built-in themes, persisted across sessions:
- Light
- Dark
- Sepia
- Nord
- Dracula

Cycle with `Ctrl+T`.

### Export
- **PDF** — true PDF via html2canvas + jsPDF (multi-page A4)
- **DOCX** — proper Word document via `docx` (headings, lists, tables, code, quotes, links)
- **HTML** — standalone HTML5 file with embedded styles
- **Markdown** — raw `.md` download

### Quality of Life
- Adjustable font size (`Ctrl++` / `Ctrl+-` / `Ctrl+0`)
- Status bar with word/character/line counts and reading time
- Shortcuts modal (`Ctrl+/`)
- Drag-and-drop file opening

---

## Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| `Ctrl+O` | Open file |
| `Ctrl+S` | Save |
| `Ctrl+Shift+S` | Save As |
| `Ctrl+W` | Close file |
| `Ctrl+P` | Print / export (format-dependent) |
| `Ctrl+E` | Export menu |
| `Ctrl+T` | Cycle theme |
| `Ctrl+B` | Toggle sidebar |
| `Ctrl+F` | Search in document (Markdown, Text, Code, RTF, ODT) |
| `Ctrl+1` / `Ctrl+2` / `Ctrl+3` | Preview / Split / Editor view (Markdown only) |
| `Ctrl++` / `Ctrl+-` / `Ctrl+0` | Increase / decrease / reset font size |
| `Ctrl+/` | Show all shortcuts |

While editing a DOCX, its own editor shortcuts (Ctrl+B/I/U/E/L/R/J/P/S/K/F/H, Ctrl+1/2/5 for line spacing, Ctrl+Z/Y for undo/redo) take priority over the shell shortcuts above with the same key.

---

## Development

```bash
npm install
npm run dev          # Vite dev server (browser)
npm run electron:dev # Electron dev (preferred for save/open)
```

## Build

```bash
npm run build        # Type-check + Vite production build
npm run electron:build
```

## Lint

```bash
npm run lint
```

---

## Stack

- **React 19** + **TypeScript**
- **Vite** for bundling
- **Electron 35** for desktop file I/O
- **docx-preview**, **xlsx**, **papaparse**, **react-pdf**
- **shiki**, **react-window**, **glide-data-grid**
- **marked** for parsing
- **KaTeX** for math
- **mermaid** for diagrams
- **docx**, **jspdf**, **html2canvas** for exports

