# Changelog

All notable changes to Atlas, adopted per Task P4.9/QA-24 (no changelog had
existed across 9 undocumented patch releases before this). Entries are
grouped by the internal "wave" merges that produced them — see
`.sisyphus/plans/atlas-phase3-improvement.md` for the task IDs and findings
each wave closed — rather than by individual commit, since a wave is this
project's real unit of shipped, reviewable work. Dates are merge dates from
`git log`.

## [Unreleased]

### Added

- **Spreadsheet formulas now do conditionals and lookups.** `=1>0` used to be a
  syntax error, which meant `IF` did not exist — and with it nothing conditional,
  no lookups and no criteria. A formula that used any of them displayed as plain
  text instead of a result.

  Working now: the comparison operators, `IF`, `IFERROR`, `AND`/`OR`/`NOT`,
  `SUMIF`/`COUNTIF`/`AVERAGEIF`, `VLOOKUP`/`HLOOKUP`/`INDEX`/`MATCH`,
  `ROUND`/`ROUNDUP`/`ROUNDDOWN`/`ABS`/`INT`/`SIGN`/`SQRT`/`MOD`/`POWER`,
  `LEFT`/`RIGHT`/`MID`/`LEN`/`TRIM`/`UPPER`/`LOWER`, `COUNTBLANK`, `PRODUCT`,
  `MEDIAN` and the `IS...` type tests.

  Dates (`TODAY`, `DATE`, `YEAR`, ...), `XLOOKUP`, `SUMPRODUCT` and array
  formulas are still not evaluated, and a formula Atlas cannot evaluate still
  shows its own text rather than a wrong answer. `VLOOKUP` and `MATCH` do exact
  matching only: Excel's default is an approximate match that returns a
  confidently wrong row on unsorted data, so asking for it is refused rather than
  answered incorrectly.

### Fixed

- **An error inside a formula is now reported instead of hiding the formula.**
  `=SUM(1/0,1)` shows `#DIV/0!`, the way a spreadsheet should, rather than
  falling back to displaying the formula text.
- **A blank cell is consistent now.** It counted as zero when referenced
  directly but as empty inside a range, so the two disagreed about the same
  cell. It is now blank either way, and still zero in arithmetic.

## [3.11.0] — 2026-10-01 (spreadsheet cell formatting, read and applied)

### Added

- **Spreadsheets now show the formatting they actually have.** Atlas used to draw
  every workbook as plain text: a file whose header row is bold white-on-blue
  with a currency column looked exactly like a CSV. Bold, italic, font size and
  family, text colour, cell fills and alignment are now read from the file and
  drawn in the grid.

  The formatting was never lost — saving always kept it — it was simply never
  read, because the spreadsheet library Atlas uses does not expose cell styles at
  all. Atlas now reads them out of the file itself.

  A cell the file does not style is drawn exactly as before, so an unstyled
  workbook looks unchanged, and your own theme's font is still used wherever the
  file does not ask for a specific one.

- **And you can now apply it.** The spreadsheet toolbar has bold, italic,
  underline, strike-through, alignment, a text colour, a fill colour and a
  number-format picker, applied to whatever you have selected. The toggles show
  the selected cell's current state, so pressing Bold on an already-bold cell
  clears it.

  It writes real spreadsheet formatting, not an Atlas-only annotation: Excel,
  LibreOffice and anything else that opens `.xlsx` sees the same thing. That is
  checked against LibreOffice itself rather than assumed.

  Formatting a cell you have not typed in keeps everything else about it — a date
  stays a date, a formula stays a formula, and a cell's cached value is not
  rebuilt from the text you see on screen.

  Borders, font family and font size are not in this release. Underline and
  strike-through are saved correctly but the grid does not draw them yet.

## [3.10.0] — 2026-10-01 (version history for every editable format)

### Added

- **Version history now covers every kind of document you can edit, not just
  markdown.** A Word document, a spreadsheet, a slide deck or a source file you
  are working on is captured every couple of minutes, the same way a markdown
  file already was — so an hour of editing a `.docx` without pressing Ctrl+S is
  recoverable, and the panel's **Save a version now** works wherever a document
  can be edited at all.

  A capture is not a save: it records what the document looks like right now and
  leaves the file on disk, and your unsaved-changes state, exactly as they were.
  A format with nothing to edit — a PDF, an image — still gets a version on every
  write, and does not offer the button rather than offering one that does nothing.

### Fixed

- **A restored version of a text file keeps the file's own line endings and
  encoding.** A version is put back by writing its bytes over the document, and a
  capture taken between saves recorded plain UTF-8 with Unix line endings — so
  restoring one could quietly turn a Windows-line-ending or UTF-16 file into
  something else. Captures now carry the convention the file was opened with.

## [3.9.0] — 2026-09-30 (version history; the toolbar menus made clickable again)

### Added

- **Version history: see how a document got to where it is, and go back.** Every
  time you save, Atlas keeps a copy of what the document looked like. Open
  **Version history** in the toolbar to see them listed with when they were taken,
  and put any of them back. A markdown document you are actively writing is also
  captured every couple of minutes, so the history shows the writing rather than
  just the saving.

  Saving the same thing twice does not make a second copy, and identical content
  is only ever stored once — so leaning on Ctrl+S costs nothing. Restoring is
  offered only when the current version is already saved, so nothing you have
  written can be replaced by an older copy without being kept itself. Each
  document keeps its most recent 100 versions.

### Fixed

- **Spreadsheets: a cell edit no longer loses everything after the first letter.**
  Opening the cell editor rebuilds the grid behind it, and the editor could be left
  without the keyboard — so typing "two" left just "t", Enter never committed it,
  and the Ctrl+S afterwards saved nothing. Worse on a slow or busy machine. The
  editor now takes the keyboard for itself as soon as it opens.
- **Tabs: closing a tab any way you like now puts the keyboard back where you
  would expect.** Ctrl+W already did; the X button and the middle-click close did
  not, and after using them Tab and the arrow keys went nowhere until you clicked
  something.
- **Word: typing works straight after Save As, without clicking back in.** The
  OS save dialog takes focus away from Atlas, and renaming the tab rebuilt the
  document view from scratch, so the cursor was gone and anything typed next
  simply did not appear. The cursor, the focus and the scroll position now come
  back where they were. Undo history is still cleared by a Save As.
- **Word: a floating image keeps its front-to-back order.** Two images stacked
  on top of one another both came back at the same depth when you saved, so
  whichever was in front could end up behind. The space set around an image, and
  a few display hints, were being dropped in the same way.
- **Word: a numbered or bulleted list keeps two settings Word puts on it.**
  Whether numbering restarts after a section break, and a marker Word writes on
  each list level, were dropped the first time you saved.
- **Word: redo now restores exactly the state you undid.** Redo used to be
  recomputed by re-running the edit against the document as it stood at that
  moment, which could place it differently if anything had shifted in between.
  It now restores the state directly, so redo cannot land in the wrong place.
- **Word: a style no longer loses formatting Atlas does not itself understand.**
  `styles.xml` is rebuilt from scratch on every save, so any run property Atlas
  had no field for — including `w:noProof` and `w:snapToGrid`, which sit on
  Word's own built-in styles — silently disappeared the first time you saved.
  Those properties are now carried through untouched, in the right place, along
  with the file's namespace declarations. The same applies to list levels in
  `numbering.xml`.
- **Accessibility: the tab bar, the slide rail, the slide editing surface and the
  code editor are now described correctly to a screen reader.** An automated
  accessibility audit found five real faults — among them a tab bar that
  described itself in a way assistive software rejects, on every screen in the
  app, and a code-editor colour too faint to meet the contrast standard.
- **Accessibility: an unsaved-changes tab now says so, not just shows a dot.**
  A coloured dot tells a screen reader nothing.

### Changed

- **The X on a tab is now for the mouse; the keyboard closes a tab with Delete
  or Backspace** (Ctrl+W still works too). A button inside a tab is something
  assistive software cannot describe correctly, and Delete is one keystroke
  instead of tabbing to a button. It is listed in the shortcuts dialog.
- **Windows file types are registered under Atlas-specific names.** Installing
  used to claim about thirty generic names — "Word Document", "Source Code" —
  in a registry area shared by every program on the machine, and uninstalling
  removed them again even if another program had put them there. The names
  Explorer shows you are unchanged. A side effect: `.log` files are now labelled
  "Plain Text Log" reliably, where before they and `.txt` could end up sharing
  whichever label was written last.

### Added

- **Word comments: a resolved comment can be un-resolved again.** Resolving used
  to be one-way inside Atlas — the thread disappeared from the pane and only Word
  could bring it back. The pane now has a "Show resolved" tick box; resolved
  threads appear greyed out and marked "Resolved", with an Unresolve button.

## [3.8.0] — 2026-09-27 (audit follow-through: data-loss fix, security hardening, DocxViewer decomposition)

### Fixed

- **Spreadsheets: the first character typed after clicking a cell is no longer
  lost on a slow or busy computer.** 3.7.0's fix for this held only on a fast
  machine: the grid takes focus one frame after the click, and anything typed
  before then went nowhere. Typing "HELLO" could save "ELLO", or leave the cell
  unchanged. Keys typed in that moment are now held and delivered to the grid,
  in order, once it is ready. Affects `.xlsx`, `.ods` and `.csv` alike.
- **Spreadsheets: pressing Ctrl+S straight after typing into a cell now saves
  what you typed.** On a slow machine the save could run before the edit
  reached the sheet, so the file on disk lacked it until you saved again.
- **Tabs: a Save As finishing just as you switch tabs can no longer show the
  previous document's text under the new tab's name.** Very hard to hit by
  hand; found by a test that failed only under load.
- **Spreadsheets: editing a second cell no longer loses what you typed, and the
  save after it no longer does nothing.** After committing one cell's edit, clicking
  another cell and typing straight away could silently leave the cell unchanged —
  and the Ctrl+S that followed wrote nothing at all, with no error and no hint that
  anything was wrong. Worse on a slow or busy machine, but it happened at normal
  speed too.
- **Spreadsheets: a screen reader announces a frozen row's cells in French**
  when Atlas is in French. Those cells were the one place still named in
  English.
- **An action that fails unexpectedly now says so** instead of appearing to do
  nothing. Anything the app did not anticipate used to vanish silently; it now
  raises the same kind of notification every other failure does.

### Fixed — translation

- **Twelve messages appeared in English even with Atlas set to French.** Every
  "failed to render" message for PDF, PPTX, ODP, ODT and RTF, the "loading
  slides" and "preparing document for print" notices, ODT's tracked-changes
  banner, and the placeholder shown in place of a Word drawing Atlas cannot
  display. All now translated. A new check reads the real code structure rather
  than pattern-matching text, so the next one cannot slip through the same way.

### Added

- **Word documents: resolving a comment now sticks.** Resolving a thread used to
  hide it for the current session only — reopen the file and it was back — and a
  thread you had already resolved in Word showed up as unresolved in Atlas. Both
  directions now work: Atlas reads the resolved state the file carries, and a
  thread you resolve here is written into the file when you save. Replying to a
  resolved thread re-opens it, since the reply would otherwise be hidden.
  Note the consequence: because it now persists, resolving is no longer undone by
  reopening the file, and Ctrl+Z does not reach it yet either — replying to the
  thread is currently the way back. See `docs/KNOWN_LIMITATIONS.md`.

### Security

- Four **spellcheck** channels, the **title-bar theme** channel and the
  startup **"which file did you double-click"** channel accepted a request from
  any frame in the window, not just the document itself — unlike every other
  channel, which already checked. Adding a word to your personal dictionary,
  replacing a misspelling in your document, changing the spellcheck languages
  and consuming the file queued at launch now all require the real renderer.
- A **Save As** dialog is no longer handed unchecked file-type filters, so a
  malformed request fails as a refused save instead of an unhandled error.
- The spreadsheet **decompression limit could be bypassed** by one flag. A
  workbook that declared itself Zip64 made the size check skip the file entirely
  rather than risk misreading a marker value, so a crafted file could sidestep the
  limit that exists to stop a small archive from expanding to gigabytes. Zip64
  only exists for archives over 4GB or with more than 65,535 entries — far past
  what Atlas will open at all — so such a file is now refused outright instead of
  waved through.

### Release engineering

- **Atlas will not be code signed** — an explicit decision, not an outstanding task.
  So the installer will always show Windows' "unknown publisher" warning (use
  **More info → Run anyway**), the SHA256 published with each release is the only
  way to verify you got an unmodified installer, and there will be no auto-updater
  while the build is unsigned. `docs/RELEASE.md` and `docs/KNOWN_LIMITATIONS.md`
  spell out what that means.
- **Upgrading in place was verified for the first time**, 3.7.0 → 3.8.0 on a real
  machine: your recent files, window size and settings all come through untouched,
  the Start Menu shortcut survives, and you get one entry in Add/Remove Programs
  rather than two. The cell-edit fix above was then re-checked against the actually
  installed application, not just the source tree.

### Housekeeping

- A test that waited on a lazily-loaded viewer failed intermittently in a full
  run but never on its own. Every wait in the suite shared a 1-second budget
  that only one machine's timing had to miss; it is now 5 seconds, which costs
  a passing run nothing.
- The coverage floor had been left at wave 2's measurement while five more
  waves of tests landed, leaving ~6 points of slack a regression could hide in.
  Re-measured and raised to 82/70/84/84.
- An e2e test failed once in a full run and passed on its own. Investigating it
  instead of retrying turned up the data-loss bug above (F6c), which is now fixed.
  Three earlier attempts at it were measured and reverted before the real cause was
  found; a control experiment confirms the fix rather than assuming it.
- Five items the backlog listed as open were already fixed in the code; they are
  now marked resolved with the evidence, and the list carries a note about why the
  drift happens. Two of them had been re-investigated from scratch by separate
  sessions.
- The three least-tested files in the tree now have tests: the slide-editor core
  both the PowerPoint and OpenDocument editors share (12% of its branches were
  checked, and it had no test file of its own), the spreadsheet editing toolbar,
  and the frozen-rows strip. All three are pinned by per-file coverage floors so
  they cannot quietly slide back.
- The bundle-size gate had only a *relative* limit, which a deliberate baseline
  refresh moves; it now also has absolute ceilings that a refresh cannot.
- `DocxViewer.tsx` was ~3,600 lines, most of it one component. Fourteen modules
  came out of it — font registration, pagination, selection painting,
  find/replace, zoom, the font picker, the comments pane, header/footer editing,
  saving, the prompt dialog, hyperlink insertion, field updates, print/export
  full-render and the deferred vertical-caret retry — with no behaviour change:
  3,611 lines down to 2,380, and 29 state declarations/23 effects/51 callbacks
  down to 10/13/31. What remains is the document, the selection, the undo history
  and the DOM input handlers that drive them: one component's actual job.
  `docs/ARCHITECTURE.md` records why each piece went where it did.

### Fixed — found by driving the app overnight (2026-09-25)

- **Word documents: Up/Down move the cursor one line, and Page Up/Page Down
  about one screen.** They used to send the insertion point to the very start or end of
  the whole document, so the next thing you typed landed there.
- **Word documents: clicking in the margin next to a paragraph places the
  cursor.** With the Comments panel open, a click just left of the text left a
  cursor that looked right but silently ignored everything you typed.
- **Word documents: typing in a paragraph that contains a comment, a footnote
  reference or an equation works.** It used to be refused without a word.
- **Crash recovery: reopening the file you were editing offers your unsaved
  changes.** Relaunching straight into it (double-click, Open with) used to
  skip the offer, and then overwrite the recovered changes within a second.
- **Text files keep their encoding, BOM and line endings on save.** Markdown,
  code and CSV files written with Windows line endings, a UTF-8 BOM, UTF-16 or
  Windows-1252 used to be rewritten as plain UTF-8 with Unix line endings, on
  every line, not just the ones you edited.
- **CSV: a semicolon-separated file stays semicolon-separated**, and keeps its
  final line break.
- **Spreadsheets: Ctrl+End goes to the last cell that has data**, not into the
  blank margin (where typing wrote a stray cell into the file).
- **Spreadsheets: `50%` and dates like `2024-03-14` are stored as numbers**
  with a percent/date format, as in Excel, instead of as text. `'007` (leading
  apostrophe) and cells formatted as text keep what you typed.
- **Slides: shapes without an id (files from some tools) can be selected and
  edited**, and "New slide" works in a deck that has no slide layouts.
- **Ctrl+O and Ctrl+N work while you are typing** in a Word document, the
  Markdown editor or the code editor.
- **Keyboard shortcuts pressed immediately after opening a file or typing act
  on what is on screen**, not on a moment earlier (a shortcut could run
  against the previous state of the document).

## [3.7.0] — 2026-09-21 (phase 4: features that never worked, and the tests that hid them)

Every defect below was found by **driving the real application**, not by the test
suite. The suite was fully green throughout — and in three cases was actively
asserting the broken behaviour as correct.

### Fixed — features that had never worked

- **Insert Hyperlink, Add Comment and Reply now work at all.** All three called
  `window.prompt`, which Electron does not implement, so nothing happened. Their
  tests passed because they mocked `prompt`. They now use a real in-app dialog
  with a proper label, focus trap and Escape-to-cancel.
- **Table Properties no longer discards everything you set.** Ticking "Show
  borders" and pressing Apply silently did nothing — and so did width and
  alignment, because one invalid control blocks a whole HTML form submission.
  The width field's `step` rejected Word's own default table width (6.25in), so
  the form never submitted at all.
- **Typing into a spreadsheet cell no longer eats the first character.** Click a
  cell, type `HELLO`, and the file contained `ELLO`. Measured at every typing
  speed and with pauses up to a second: 24 out of 24 attempts lost a character.
- **Insert Table produces a usable table.** It rendered at zero width, and text
  typed straight afterwards vanished entirely.
- **The caret can be placed inside a table cell.** Clicking a cell put the cursor
  in the paragraph after the table, which made cell merge unreachable.
- **Insert Text Box on a slide keeps what you type**, and creates one shape
  instead of two.

### Fixed — documents Word may have refused to open

- **Every colour Atlas wrote was invalid OOXML** (`w:val="#ff0000"` — the leading
  `#` is not legal). Files already saved by an earlier version repair themselves
  when reopened.
- **New bullet and numbered lists were invalid OOXML** (a non-integer
  `abstractNumId`).
- The package validator now checks attribute datatypes, so it would have caught
  both. It passed them clean before.

### Fixed — silent data loss on save

- **Theme fonts** (how Word has specified fonts since 2007) were dropped on every
  save. A general passthrough now preserves any run property Atlas does not model,
  including a whole class of character effects that were being stripped.
- **Excel table filters and sorts** were wiped on every save, even when the table
  was untouched.
- **Whole-column conditional formats, data validation and hyperlinks** were
  deleted by any row or column insert.
- **Content controls and shape fallbacks in footnotes, endnotes and comments**
  were stripped with no warning.
- **Save As** could rename a different tab and re-read that file over your work.
- An **xlsx save that cannot preserve the original layout now says so** instead of
  silently downgrading.
- **Formulas returning text** (`=A1&"!"`, `CONCATENATE`) were saved as literal
  formula text; formulas depending on later formulas showed stale values.

### Fixed — shell, keyboard and accessibility

- **Ctrl+W did nothing while the cursor was in a text field** — including the
  markdown editor, so the unsaved-changes prompt never appeared on that path.
- **Ctrl+G** now opens Go to line in the code editor.
- Closing a tab by keyboard no longer drops focus to nowhere; the shortcuts
  dialog, unsaved-changes dialog, presenter view and PDF find bar all trap Tab.
- The **markdown source editor** is translated and has a real accessible name.
- A draft you **discarded** no longer comes back as a crash-recovery offer —
  including when discarding at quit.
- A document undone back to its saved state is no longer marked unsaved forever.
- The **New** menu no longer opens partly off the left edge of the window.

### Security

- A compromised renderer could register **any existing file** as writable and
  overwrite it with no dialog. Write access now requires a path the main process
  itself vouched for.
- The spreadsheet read path had no decompression limit, so a crafted file could
  exhaust memory on open.

### Housekeeping

- Test runs leaked an Electron profile directory per launch, never deleted —
  **2,876 directories and 26 GB** on the development machine. They are now pruned.

## [3.6.0] — 2026-09-20 (wave 11: schema conformance, accessibility, memory)

### Fixed
- **Word could have refused documents Atlas saved.** Several parts were
  written with their elements in the wrong order — section properties
  (headers/footers were emitted last instead of first), paragraph and run
  properties, table and cell properties, numbering levels, and the styles
  part's own copies of all of those. Word's schema is order-sensitive, and a
  wrong order is the classic cause of "Word found unreadable content". The
  package validator now checks these sequences on every build, and the
  shipped blank templates were regenerated.
- **Spreadsheets**: a shared formula's range now follows row and column
  inserts; a formula broken by a deleted sheet stores the error as its value
  (so it stays visible when reopened) instead of a stale number that looked
  like a real answer; images belonging only to a deleted sheet are swept,
  while shared ones are kept.
- **A closed spreadsheet tab leaves ~24 MB behind** — measured, root-caused
  to the grid library, and documented in `docs/KNOWN_LIMITATIONS.md` rather
  than guessed at.

### Added
- **Keyboard and screen-reader support**: the tab strip follows the standard
  arrow-key pattern, the spreadsheet grid announces the selected cell, the
  PDF find bar announces its match count, thumbnail rails have proper
  semantics, and smooth scrolling now respects the system "reduce motion"
  setting. A keyboard-only end-to-end journey (no mouse at all) guards it.

## [3.5.0] — 2026-09-20 (wave 10: fidelity warnings, responsive Markdown, memory)

### Added
- **Large Markdown documents render without freezing the app**: parsing now
  runs off the main thread, so the window stays responsive while a big
  document is being prepared, and documents up to 2 MB preview automatically
  (the earlier limit was 750 KB). Rendering output is proven identical to
  before by a parity test over every Markdown characterization fixture.
- **DOCX Save now tells you when it couldn't preserve everything** —
  the lossy-save detector added in 3.4.0 was reporting into a void; it's
  now wired into the Save flow (`DocxViewer.tsx`). A dismissible,
  non-blocking notice appears after a save that dropped something, once
  per document, in plain language ("content controls", "a text box's
  fallback drawing" — never an XML element name).

### Fixed
- **Sheet names with accents** (`Résumé`) were invisible to the formula
  rewriter, so renaming or deleting such a sheet left every formula pointing
  at the old name instead of being updated — a silent wrong-number bug.
- **Closing a tab no longer keeps the document in memory**: up to five closed
  documents' contents stayed reachable for nothing. Reopening a closed tab
  (Ctrl+Shift+T) now re-reads the file, and a file deleted in the meantime
  says so instead of reopening as an empty document.
- **An unedited content control (`w:sdt`) or a shape/text box's legacy
  fallback (`mc:AlternateContent`) now survives a save byte-for-byte** —
  including a content control's id/alias/tag/binding/lock/placeholder/
  appearance, and a shape's rejected `mc:Fallback` branch, none of which
  Atlas models — instead of always being stripped to its bare content.
  Editing content inside one, or one sitting inside a table cell/header/
  footer, still falls back to the previous (reported) behavior. See
  `WrapperPassthrough` in `src/docx/model/document.ts`.

### Verification
- New end-to-end scenarios: password-protected PDFs (with a genuinely
  encrypted fixture built for the purpose), a 36-page image-heavy DOCX under
  fast scrolling and editing, spreadsheet formulas through a rename + delete +
  row insert in one save, closing a tab mid-load, and several dirty tabs at
  once. A full hunt through the French UI, the New-document flow and those
  races found no further defects.

## [3.4.0] — 2026-09-20 (waves 7-9: data-loss fixes, French UI, document fidelity)

### Added
- **The interface can be French** (English by default; a language menu sits
  next to the theme menu, with a "System" option). 509 strings, Office's own
  French terminology and French typography.
- A **spec-level package validator** (`scripts/validate-office-file.mjs`) that
  checks what Atlas writes against the OPC/OOXML and ODF rules — content
  types, relationships, element order, ODF's stored-first `mimetype` — and
  runs in CI. It found two real "Office offers to repair this file" causes,
  both fixed.
- A **lossy-save detector** for DOCX (`src/docx/fidelity/`): it diffs a
  re-save against the original document rather than trusting a hand-written
  list, so anything the engine does not model is reported instead of being
  dropped in silence.

### Fixed
- **Save As no longer loses work.** Ctrl+Shift+S did nothing at all outside
  markdown, and Save As left the document's tab pointing at the file it was
  opened from: coming back to that tab re-read the OLD file over the user's
  work, and the next Ctrl+S wrote to neither file.
- **Every saved DOCX image lost its picture properties** (position, geometry,
  fill), on every save. Right-to-left paragraphs and sections, page borders,
  multi-level list definitions and complex-script bold/italic were dropped
  too. All preserved now.
- **Legacy .doc/.ppt files failed to open** as soon as their text passed
  ~4 KB — that is, every real document. The test fixtures were all smaller.
- **Spreadsheet formulas** that reference a renamed, moved or deleted sheet
  are rewritten (`#REF!` on delete, as Excel does), and multi-area, function-
  wrapped and whole-row/column defined names are re-anchored through
  structural edits.
- A running **Run** program kept going after its tab was switched away, with
  no way to stop it; a **read-only file** was reported as "open in another
  program"; a dropped **folder** produced a raw technical error; a multi-file
  **drag and drop** opened only the first file.
- The **PDF find bar** ignored Escape unless its input had focus, then
  swallowed the next keyboard shortcut.
- **A very large Markdown document no longer freezes the app**: past 750 KB
  the preview waits to be asked for (the editor opens normally), because
  Markdown parsing costs grow superlinearly — a 5 MB file took over two
  minutes.
- **Accessibility**: WCAG AA contrast failures fixed in all five themes,
  keyboard focus trapped and restored in every dialog, overlay and menu, and
  the unsaved-changes dialog could previously be tabbed straight past.

### Changed
- **Start-up is ~33% faster** (766 ms → 514 ms to the welcome screen): the
  export tools and the Markdown renderer no longer load before the first
  paint, cutting the entry bundle by 89%.
- `npx tsc -b` now type-checks everything — `src`, the Electron main process,
  the end-to-end tests and the build scripts (the strict-mode ratchet, P4.1,
  is closed).
- Documentation (README, architecture, known limitations, release guide)
  re-checked against the code it describes.

### Added (detail)
- DOCX header/footer editing (D29 follow-up 2): a paragraph that MIXES plain
  text with a drawing, a field (PAGE/NUMPAGES/DATE/etc), a hyperlink, a
  footnote/comment reference, or an existing tracked change — the extremely
  common real-world shape "Chapter title .......... Page X of Y", or a logo
  image followed by a title — can now have its TEXT edited. The header/
  footer panel renders such a paragraph as one editable input per plain-text
  span, interleaved with a read-only, labelled chip per atom in between
  ("[Image]", "[Page number]", "[Total pages]", "[Date]", "[Link: …]", …);
  the atom itself is never read or rewritten by this editor, so a field's
  own runs (`w:fldSimple`, or the `w:fldChar`/`w:instrText` triple) and a
  drawing's own XML can never come apart from editing the text around them.
  Editing one text segment reuses the same prefix/suffix run diff the
  whole-paragraph case already used, scoped to that segment's own slice of
  the paragraph, so every OTHER run — including every other text segment
  and every atom — keeps its exact original object identity and formatting.
  A bookmark/comment-range boundary sitting mid-paragraph splits the text on
  either side into separate inputs (so an edit can never merge across it)
  without ever being shown as its own chip. Previously such a paragraph was
  ONE opaque, fully read-only placeholder with no way to touch even its
  plain-text parts; a paragraph with no editable text anywhere, and a table
  block, are still a read-only placeholder exactly as before. First-page/
  even/default header-footer variants and section sharing, one-undo-step-
  per-edit, and "Ctrl+S with a field still focused commits pending text
  first" all continue to work unchanged, now per SEGMENT instead of per
  paragraph. See `docs/KNOWN_LIMITATIONS.md`'s DOCX section for exactly
  what still stays fully read-only and why.

## [3.3.0] — 2026-09-20 (waves 5-6: new documents, long-document speed, release hardening)

### Added
- **New document** (NEW-01): a "New" button/menu in the toolbar next to Open,
  plus Ctrl+N, creates a brand-new Word/Excel/PowerPoint/OpenDocument
  Spreadsheet/OpenDocument Presentation/Markdown document — a native Save
  dialog picks the destination, main writes a real, Office/LibreOffice-
  compatible blank template there (generated by `scripts/generate-
  templates.mjs`, committed under `electron/templates/`), and it opens in a
  new tab.
- Opening a genuinely 0-byte file of one of those same formats (e.g. one made
  outside Atlas, like Windows Explorer's "New > Word Document" on a PC
  without Office) now opens that format's blank template instead of failing
  to parse ("Failed to unzip DOCX: End of data reached…"-style errors),
  still bound to the file's own path — a normal Save fills it in for real.
  A 0-byte file of any other format (PDF/RTF/ODT/legacy `.doc`/`.ppt`/
  unrecognized) now shows a plain "this file is empty" message instead of a
  parse error.

### Fixed
- Spreadsheets: saving an .xlsx/.xlsm through the original package (USR-17)
  now also covers adding, deleting, renaming and reordering sheets (workbook
  relationships, content types and, where the formula is a single-area
  reference, defined names are kept in step; a deleted sheet's own
  `localSheetId`-scoped defined names are dropped and later ones renumbered)
  instead of losing all styling to the fresh-workbook fallback; conditional
  formatting, data validation, hyperlinks and the sheet-level autoFilter are
  now re-anchored through row/column inserts and deletes the way Excel itself
  keeps them aligned, instead of silently pointing at the wrong cells.
- Security: the spreadsheet passthrough's zip reads (`xlsxPassthrough.ts`,
  `spreadsheetPanes.ts`, `spreadsheetTables.ts`) now reject a workbook whose
  declared uncompressed size is implausibly large before decompressing
  anything, closing the same zip-bomb gap already fixed for DOCX/PPTX.
- DOCX header/footer editing (D29 follow-up): an edit no longer flattens a
  header/footer paragraph into a single plain-text run — a paragraph holding
  an image, a PAGE/NUMPAGES field, a hyperlink, or other non-text content is
  now shown as a read-only placeholder and left completely untouched;
  editing a plain-text paragraph splices only the changed span into its own
  runs, so unedited text keeps its exact formatting. Also fixed: pressing
  Ctrl+S while a header/footer field still had focus saved the stale
  pre-edit text (the field only committed on blur) — the pending edit is now
  committed before every save.
- A 0-byte `.md` file opened the Welcome screen instead of an (empty)
  markdown editor bound to that path — `hasContent` no longer keys off
  markdown content *length* when a real file is loaded.
- DOCX: typing in a long document is another ~1.5-2x faster on top of wave
  4's fix (D23-PERF). The remaining cost was almost entirely wasted React
  rendering, not layout: `PageStack`'s `document` prop tracked the live,
  every-keystroke-changing document model instead of the document a
  completed page layout was actually produced from, so every page re-ran its
  whole-document run/hyperlink/bookmark metadata walk TWICE per keystroke —
  once before repagination caught up, once after. `PageStack`/`PageView` are
  now memoized, that metadata is computed once per document (in `PageStack`)
  instead of once per page, and `PageStack` only ever sees a `document`/
  `pages` pair from the same pagination pass — measured (instrumented build)
  at 96 `PageView` renders per keystroke before, 24 after, on a 24-page
  document. Also fixed a paginate.ts cache bug: a paragraph after ANY
  footnote reference in its section was wrongly treated as depending on
  document-wide state and permanently skipped the per-paragraph line cache;
  now only a paragraph that itself carries a footnote/endnote reference does.
- A rapid second launch (double-clicking a file, or "Open with → Atlas"
  again) while the first window still hadn't finished starting up could
  silently drop the requested file — not just in the already-fixed case
  where the window object didn't exist yet, but also in the narrower gap
  between the window being created and its renderer finishing mount
  (ELEC-07). The second-instance and macOS `open-file` handlers now queue
  the request until the renderer has proven it's actually listening, instead
  of guessing with a fixed delay.
- Packaged `.exe`: fixed a release-packaging config bug where
  `signAndEditExecutable: false` (meant only to skip code signing, since
  there's no certificate yet) was also silently disabling icon and
  version-metadata embedding — the exe would have shipped with no custom
  icon, FileDescription, CompanyName, or version info. The exe now carries
  correct ProductName/CompanyName/LegalCopyright/FileVersion metadata and
  the Atlas icon even unsigned; see `docs/RELEASE.md` for what real code
  signing would still require (ELEC-09/ELEC-10).

- DOCX: typing in a long document is roughly 2x faster again on top of the
  D23-PERF fix above (D23-PERF-2). Measuring the same keystroke split into
  command-apply/pagination/React-commit/browser-layout-paint found the
  React commit and browser paint steps — not pagination, already cheap
  thanks to the line cache — dominating: `paginate()` hands back a
  brand-new `Page` object for every page on every call, so even with
  `PageStack`/`PageView` memoized, every page on screen still failed its
  memo check and remounted a full DOM subtree on every keystroke regardless
  of whether it was actually visible. `PageStack` now virtualizes: only
  pages near the scroll container's viewport (plus a small buffer, plus
  whichever page holds the caret/selection) mount a real `PageView`; the
  rest render as a lightweight placeholder that reserves the same box, so
  scroll height/the scrollbar/the page-count status bar are unaffected.
  Printing and PDF export force every page to mount first (both read the
  live DOM). Measured back-to-back on the same machine: React commit
  dropped from ~40-50ms to ~4-10ms and browser paint from ~50-60ms to
  ~10-20ms per keystroke; wall-clock median across 5 keystrokes on a
  35-page document went from ~280-300ms to ~141-155ms.

### Changed
- Strict-mode ratchet (P4.1): `electron/` (`main.cjs`, `preload.cjs`,
  `lib/*.cjs`) is now type-checked — a new `tsconfig.electron.json`
  (CommonJS, `allowJs` + `checkJs`, `strict`, Node types) is referenced from
  the root `tsconfig.json`, so plain `npx tsc -b` (what CI already runs) now
  covers it alongside `src/`, `vite.config.ts` and `tests/e2e`. Fixed all ~95
  errors this uncovered across `electron/preload.cjs`, `electron/main.cjs`
  and `electron/lib/{codeRunner,systemFonts,printToPdf,csp,fileSizeGuard,
  recentFilesStore,atomicWrite}.cjs` with real JSDoc parameter/return types
  and genuine null/undefined narrowing — no `@ts-ignore`/`@ts-nocheck`/`any`
  loosening. Two real gaps this turned up and fixed: `main.cjs`'s
  `set-theme` IPC handler indexed the overlay-color table with the
  renderer-supplied theme value without validating it was one of the known
  keys (now checked with a type-guard before indexing, matching how every
  other untrusted-IPC-input handler in this file already validates); and the
  window's `ready-to-show` handler called `mainWindow.show()` with no
  null/destroyed guard, unlike its sibling handlers (`did-fail-load`, the
  5s fallback timer) which already defend against the window having closed
  in the interim. `scripts/*.mjs` and `tests/e2e/fixtures/*.mjs` remain
  unchecked — see `STRICT_MODE_TODO.md`.

- Markdown HTML/PDF export (UX-20): the exported document's `h4`/`h5`/`h6`
  heading sizes, task-list checkbox styling, and `<details>`/`<summary>`
  styling had drifted from the live preview (e.g. `h4` exported at `1em`
  instead of the live preview's `1.05em`); the export CSS now matches, and a
  new crossref test (`markdownExportCss.crossref.test.ts`, mirroring the
  existing theme-color crossref test) fails if the two ever drift apart
  again.

## [3.2.0] — 2026-09-19 (wave 4: owner-reported editor defects)

### Fixed
- DOCX editing: typing, selection, caret placement, bold/italic/underline and
  alignment toggles, Find (Ctrl+F), Ctrl+Y/Ctrl+A, highlight, and a searchable
  font picker fed by the installed Windows fonts (USR-01…USR-14).
- Spreadsheets: cell editing was completely dead in 3.1.0 (missing `#portal`
  element); Enter now commits reliably; Excel tables are read and kept; an
  .xlsx/.xlsm is saved through the original file so styles, number formats,
  charts and filters survive (USR-17).
- Slides: text no longer clipped (body insets/anchor); rotated shapes select
  and resize correctly (USR-15, USR-16).
- Typing in a long DOCX is about 6× faster; right-to-left paragraphs (D23).
- Review fixes: switching tabs after a save no longer shows (and then
  re-saves) the pre-save text; markdown Save As no longer sends later saves to
  the old file; undo followed at once by Save in the slide editors saves the
  undone deck; Run approval is tied to the file's content.
- Hardening: .pptx/.odp zip-bomb budget, escaped layout attributes,
  XML-illegal control characters dropped on save, .ppt slide-count cap.

### Added
- A PowerPoint-like editor for .pptx **and .odp**: edit text in place,
  move/resize/insert/delete shapes, add/duplicate/delete/reorder slides,
  speaker notes, presenter view, Save/Save As (USR-16).
- A CodeMirror code editor and an explicit, confirmed, sandboxed Run for
  JavaScript/TypeScript/Python files (USR-18, USR-19).
- Several documents open at once, in tabs (Ctrl+Tab, Ctrl+W, Ctrl+Shift+T,
  middle-click, drag to reorder) (SHELL-17).
- Header and footer editing in DOCX (plain text) (D29).
- Read-only viewers for legacy Word 97-2003 (.doc) and PowerPoint 97-2003
  (.ppt) files (text only).
- End-to-end scenarios driving the real app for every editor, tabs and the
  open→edit→save→reopen journey (P4.6).

### Changed
- Electron 35 → 44.4.1, electron-builder 26.15.3, sharp 0.35.4; `npm audit`
  reports 0 vulnerabilities (P4.5).
- `strict` is explicit in every tsconfig and the e2e specs are type-checked
  (P4.1).

## [3.1.0] — 2026-09-15 (wave 3, docs-quality slice)

### Added
- `docs/ARCHITECTURE.md` and `docs/KNOWN_LIMITATIONS.md`.
- This changelog.
- A per-chunk bundle regression gate (`scripts/check-bundle.mjs`, wired as a
  `postbuild` step) covering every `dist/assets/*.{js,css}` chunk, not just
  the main entry bundle, with a freshly captured baseline.
- Coverage thresholds (`vitest.config.ts`) and an `npm run coverage` script.
- A shared, typed `createMockElectronAPI()` test fixture
  (`src/__tests__/mocks/electronAPI.ts`) and a dedicated `useRecentFiles`
  test suite.
- CI: Node-24-compatible action majors, a coverage-artifact upload step, and
  a Windows Electron end-to-end job on pushes to `main`.

### Fixed
- `npm run electron:preview` now actually previews the production build
  instead of silently falling back to a dev server that might not be
  running (`electron/main.cjs`'s dev/prod detection now keys off whether
  `dist/index.html` exists, with an `ATLAS_DEV` override).
- README's feature/shortcut/stack documentation brought back in sync with
  what the app actually does (previously described only the Markdown
  feature set from Phase 1, with zero mention of the other 12 formats).

## [3.0.x] — 2026-09-14 (wave 2)

Eight parallel worktrees, merged into `main` over one day. Full task-ID
mapping in the improvement plan; highlights below.

- **DOCX fidelity** (`wave2/docx-model`, `wave2/docx-layout`): style-cascade
  resolution wired through layout, alignment/indent/spacing, justification,
  list markers, hyperlink rendering, manual and section page breaks, table
  style cascade, `sdt`/`AlternateContent`/symbol handling, resolved
  `commentsExtended` state, serializer round-trip fidelity, Aptos font
  mapping with lazy-loaded font assets.
- **DOCX editing** (`wave2/docx-editor`): cross-paragraph editing,
  transactional undo/redo, image insert via the same history, a spellcheck
  fix, accept/reject track changes plus a `trackChanges` setting, list/
  indent/hyperlink/page-break/table insert commands, Save As.
- **PDF viewer** (`wave2/pdf`): page virtualization, text and annotation
  layers, in-document find, print, interactive-form-field rendering, page
  rotation, a thumbnail rail, password-protected PDF support.
- **Slides** (`wave2/slides`): PPTX/ODP layout and master/slide-layer
  inheritance, run formatting, bullets, tables, grouped shapes, speaker
  notes, keyboard navigation, thumbnails.
- **Spreadsheets, CSV, code, RTF, ODT** (`wave2/data-text`): formatted cell
  values, merged cells, column widths, hidden sheets, off-main-thread
  worker parsing (with a 200ms task-perf budget spec), text/code
  virtualization, RTF/ODT size/security guards, expanded fixtures.
- **Format detection & routing** (`wave2/formats`): the single canonical
  extension manifest (`src/formats/extensionManifest.ts`, generated at
  `prebuild`), extensionless/legacy-CFB detection, magic-byte verification
  on the recognized-extension fast path, the `UnknownViewer` empty state.
- **Export & UX polish** (`wave2/export-ux`): markdown export reads the
  live rendered DOM instead of re-parsing (fixing export/live-render
  divergence), DOCX export gained math/Mermaid/hyperlink/task-list
  support, Electron-native save dialogs for exports, a toast system
  replacing blocking `alert()`s, real CSV export, accessibility and
  contrast fixes.
- **Shell** (`wave2/shell`): the centralized shortcut dispatcher
  (`ShortcutManagerProvider`, replacing six independent `keydown`
  listeners), main-process close-confirmation guard, draft recovery,
  close-file + dynamic window title, in-document search extended beyond
  markdown to text/code/RTF/ODT.

## [3.0.x] — 2026-09-13 (wave 1)

- **CI & lint** (`wave1/ci-lint`): CI workflow stood up, lint brought to
  zero warnings/errors.
- **Markdown characterization** (`wave1/markdown-characterization`): the
  snapshot suite that is the executable definition of "markdown behavior
  must not change" for every later wave.
- **DOCX round-trip corpus** (`wave1/docx-corpus`): a fixed fixture set
  parsed → serialized → re-parsed and diffed, gating every later
  serializer-touching change.
- **DOCX comments fix** (`wave1/docx-comments`): corrected
  `extractCommentText` operator-precedence bug.
- **Viewer quickfixes** (`wave1/viewer-quickfixes`): assorted viewer polish
  and error-recovery fixes.
- **DOCX save integrity** (`wave1/docx-save-integrity`): save-path
  robustness hardening.
- **Electron hardening** (`wave1/electron-hardening`): main-process
  security hardening — sandboxing, IPC allowlisting, and related fixes.
- **Shell session** (`wave1/shell-session`): the unified dirty-state/save
  capability contract (`ViewerContext`'s `isDirty`/`setDirty`/
  `registerSave`) and session UX (recent-files dedup, discard-changes
  confirmation, browser-mode guard).
- **Dependency & build-chain security** (`wave1/deps-security`): dependency
  hardening.

## Before 3.0 — Phases 1 & 2

Phase 1 (view-only, 13 formats, Atlas rebrand from "MD Reader") and Phase 2
(hand-coded DOCX editing engine) predate this changelog's adoption. See
`.sisyphus/plans/atlas-phase1.md` and `.sisyphus/plans/atlas-phase2-docx.md`
— both now carry a dated "Actual State" addendum reconciling their
historical claims against what shipped.
