# Known Limitations

An honest, per-format list of what Atlas does **not** do today, so a user
(or a future contributor) doesn't have to discover a gap by hitting it. This
is a living document — update it whenever a limitation below is closed or a
new one is found, per Task P4.9/QA-18's "documentation must track reality"
fix. Written against `main` after wave 2 (commit `03b2e2a`); items already
under active (unmerged) development on a `wave3/*` branch are called out
explicitly as **in progress**, not shipped.

For the reasoning behind *why* something below is deferred rather than
fixed, see `.sisyphus/plans/atlas-phase3-improvement.md` Section 8
("Explicitly De-Scoped / Deferred Items") — most rows here map directly to
a `DEFER-N` entry there.

---

## Editing scope, overall

Only **Markdown** and **DOCX** support editing + save today. Every other
format — PDF, PPTX, ODP, XLSX/ODS spreadsheets, CSV/TSV, plain text, code,
RTF, ODT — is **view-only**: you can open, read, search, and (per-format,
see below) export or print, but not edit the source file in place.

**In progress (wave3/sheets, unmerged):** real spreadsheet cell editing
(edit a cell, insert/delete rows and columns, undo/redo, TSV-shaped
clipboard paste) with write-back to `.xlsx`/`.xlsm`/`.xls`/`.ods`/`.fods`
and CSV/TSV. Not yet available in a build off `main`.

## DOCX

Supported: paragraphs/runs/tables/images/hyperlinks/bookmarks/comments/
footnotes-endnotes; run formatting (bold/italic/underline/strike/sub-super/
font/color/highlight/spacing/caps); paragraph formatting (alignment,
spacing, indent, shading, borders, tab stops); ordered/unordered/multi-level
lists; the paragraph/character/table style cascade including inheritance;
insert-table, page breaks, section breaks; headers/footers; find & replace;
track-changes accept/reject (single and all) with a `trackChanges` on/off
setting; comments; undo/redo (command-pattern, bounded history); image
insert; Save and Save As with an atomic, lock-aware write path.

**Not supported:**
- **Table structural editing** — no row/column insert or delete, no cell
  merge/split, no column resize once a table exists (only inserting a whole
  new table is supported). *(DEFER-1)*
- **Paste fidelity** for tables/hyperlinks/images/colors pasted from Word or
  the web is limited — plain-formatting paste works, structural paste does
  not yet reconstruct the source's table/image structure. *(DEFER-1)*
- **RTL / bidi text** (Arabic, Hebrew, Persian, Urdu) is not modeled or
  laid out. *(DEFER-2)*
- **Embedded fonts** (`w:embedTrueTypeFonts`, obfuscated font parts) are not
  read — a document embedding a non-standard font falls back to a
  substitute. **In progress (wave3/docx-fields-fonts, unmerged):**
  de-obfuscation and parsing of embedded fonts. *(DEFER-4)*
- **Live field/TOC regeneration** — `PAGE`/`REF`/table-of-contents fields
  round-trip with their last-cached text preserved but do not recalculate
  on edit. **In progress (wave3/docx-fields-fonts, unmerged):** field
  evaluators and an explicit "Update Field(s)"/"Update TOC" action.
  *(DEFER-5)*
- **Equations (OMML)** render via a MathML conversion; not editable.
- **SmartArt and embedded charts** render as a static placeholder image;
  not editable.
- **Macros** (`vbaProject.bin`) round-trip byte-for-byte but are never
  executed — Atlas has no VBA runtime.
- **Legacy form fields** (`FORMTEXT` etc.) round-trip but have no editing
  UI.
- **Ink annotations** and **embedded 3D models** are not rendered or
  editable.
- **Real-time collaboration** (no CRDT) and **DRM/IRM**-protected documents
  are not supported.
- **In progress (wave3/docx-pagination, unmerged), not yet on `main`:**
  table rows splitting across a page break, footnote/endnote layout,
  per-section header/footer vertical alignment, hyphenation, a pagination
  zoom control.
- **In progress (wave3/docx-drawings and wave3/docx-editing, unmerged):**
  no commits yet at time of writing — branches exist but scope has not
  landed.
- **Spell check** runs through Electron/Chromium's own native spellchecker
  (`window.electronAPI.spellcheck`, backed by `session.setSpellCheckerLanguages`)
  rather than the bundled-Hunspell-dictionary (`nspell` + `dictionary-en`/
  `dictionary-fr`) approach the original Phase 2 plan locked in — see
  `.sisyphus/plans/atlas-phase2-docx.md`'s Actual-State addendum for why.
  Practical effect: available languages are whatever the OS/Chromium
  spellchecker supports, not a bundled dictionary pair.

## Legacy binary Office formats (`.doc`, `.xls`, `.ppt`)

Detected by CFB (OLE Compound File Binary) magic bytes and given a specific,
friendly message pointing at the modern equivalent (`.docx`/`.xlsx`/`.pptx`)
— Atlas does **not** parse the legacy binary format itself. Full OLE-CFB
parsing for pre-2007 Office documents is out of scope for this plan
(declining real-world frequency; a separate, large undertaking).

## PPTX / ODP (slides)

View-only: layout and master/slide inheritance, run formatting, bullets,
tables, grouped shapes, speaker notes, keyboard navigation between slides,
and a thumbnail rail. No editing of any kind (text, shapes, layout, or
reordering slides).

## PDF

View-only: page virtualization (large PDFs render without loading every
page up front), a text layer (select/copy text, in-document find), an
annotation layer (links and form-field widgets are rendered visually),
print, page rotation, a thumbnail rail, and password-protected PDF support.
**Form fields are rendered but not interactively fillable, and there is no
PDF editing or save-back** — Atlas does not write PDF files.

## Spreadsheets (XLSX/ODS) and CSV/TSV

On `main` today: view-only. Formatted cell values (numbers/dates/currency
per the workbook's own number formats), merged cells, column widths, hidden
sheets, and frozen-pane metadata are read and displayed; parsing runs off
the main thread in a worker for large files. No cell editing, no formula
evaluation UI, and no save path exist on `main`.

**In progress (wave3/sheets, unmerged)** — see "Editing scope" above.

Number/date/currency formatting follows the workbook's own stored format,
not an explicit user-chosen locale; no UI localization exists elsewhere in
Atlas either (see the improvement plan's DEFER-6).

## Text / Code

View-only, virtualized for large files (renders visible rows only). Code
gets syntax highlighting for 30+ languages via Shiki. In-document find
(`Ctrl+F`) works in both. No editing.

## RTF / ODT

View-only. Rendered output is sanitized with DOMPurify (forbidding
`<script>`/`<style>`/`<iframe>`/`<object>`/`<embed>` and `on*` handlers)
since both formats hand back a live DOM tree the underlying library builds
itself from untrusted document bytes. In-document find works in both. No
editing.

RTF rendering goes through `rtf.js`, which also renders embedded Windows
Metafile (WMF) and Enhanced Metafile (EMF) images — see
["Build & performance: the `rtf.js` bundle size"](#build--performance-the-rtfjs-bundle-size)
below for why that library is large and why Atlas keeps it anyway.

## Multi-document support

Atlas is single-document: opening a new file replaces the current one
(after an unsaved-changes prompt if needed). There are no tabs and no
side-by-side multi-document view. This is a deliberate, larger
architectural feature left for a future roadmap item, not a bug.

## Export

Export targets and fidelity vary by format — see the README's Export table
for the current, per-format matrix (kept as the single source of truth
there so this document doesn't drift out of sync with it).

## Dev tooling: dev/prod detection depends on `dist/` existing

`electron/main.cjs` decides dev vs. prod (RUN-12, `electron/lib/devDetect.cjs`)
by checking whether `dist/index.html` exists on disk, not by how the process
was launched. This fixed the real bug it targets (`npm run electron:preview`
silently falling back to a dev server instead of previewing the production
build), but it means a `dist/` left over from an earlier `npm run build` or
`npm run electron:build` makes every subsequent plain `electron .` — including
a local `npx playwright test` run, since `tests/e2e/*.spec.ts` launch
Electron directly with no explicit mode — resolve to prod against that
possibly-stale build instead of the live dev server, with no warning that
this happened. Set `ATLAS_DEV=1` explicitly when you want the dev server
regardless of what's on disk (or delete `dist/` first). CI's `e2e-windows`
job never hits this, since it always runs a fresh `npx vite build`
immediately before the e2e step.

## Build & performance: the `rtf.js` bundle size

`rtf.js` is, by a wide margin, the single largest chunk in Atlas's
production build (~2.24 MB minified / ~841 KB gzipped as of this wave's
measured baseline — see `.sisyphus/baselines/atlas-phase3-bundle.json`),
for what is likely the least-used of the 13 supported formats. This was
investigated (Task P4.3 / RUN-15 / QA-17) as a candidate for replacement
with a lighter converter; the conclusion was **not to replace it**, for two
reasons:

1. **It is lazy-loaded.** `formats/registry.ts` only `import()`s `rtf.js`
   when a user actually opens an `.rtf` file — it contributes nothing to
   Atlas's cold-start bundle or memory footprint otherwise, and Atlas ships
   as a bundled Electron installer rather than a byte transferred over a
   slow network on every session, so this size does not affect app launch
   time the way it would for a web app.
2. **The size is inherent to real functionality, not bloat.** `rtf.js`
   bundles three sub-modules — `RTFJS` (RTF parsing), `WMFJS` and `EMFJS`
   (Windows/Enhanced Metafile image rasterization) — because real-world RTF
   files exported from legacy Word/WordPad commonly embed `\pict\wmetafile8`
   clipart and images. A lighter text-only RTF→HTML converter exists as an
   option, but every such alternative surveyed drops WMF/EMF image
   rendering entirely, which is a genuine fidelity regression for that
   class of document, not a size optimization. Prototyping confirmed the
   trivial e2e RTF fixture (plain text, no embedded images) would still
   render fine under a lighter converter, but that fixture doesn't exercise
   the actual reason the library is large, so passing it isn't evidence the
   swap is safe for real documents.

The bundle regression gate (`scripts/check-bundle.mjs`, wired as a
`postbuild` step) now tracks every chunk including this one, so any
*further* unexpected growth is still caught — this is a decision to keep
the current size, not to stop watching it.
