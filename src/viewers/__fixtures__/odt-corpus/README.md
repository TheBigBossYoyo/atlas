# ODT fixture corpus

Plan task: **T8** (finding **DAT-19**).
See `.sisyphus/plans/atlas-phase3-improvement.md` (T8) and
`.sisyphus/plans/atlas-phase3-findings-register.md` for the full task text
and evidence this corpus exists to guard against.

## What this is

Three small, synthetic `.odt` fixtures, each built to exercise one
real-world ODF feature: a table + bulleted/numbered lists
(`tables-and-lists.odt`), an inline base64-embedded image
(`embedded-image.odt`), and a block-level tracked insertion + deletion
(`tracked-changes.odt`). They're consumed by
`src/viewers/__tests__/OdtViewer.corpus.test.ts`, which parses each one
with `odf-kit`'s `readOdt`/`toHtml` and snapshots the resulting HTML.

`odf-kit` is a pre-1.0 dependency (DAT-19), now pinned to an exact version
in `package.json` rather than a `^` range, specifically so a minor bump is
a deliberate, reviewed action rather than something `npm install` can
change silently. This corpus is the second half of that guard: even a pinned
version bump reviewed and accepted here should show its actual HTML-output
diff against real ODF constructs, not just "the version number changed".

## Regenerating the fixtures

```
node scripts/generate-odt-corpus.mjs
```

This overwrites every `.odt` file in this directory. The generator is
deterministic — re-running it with the same `jszip`/Node versions
reproduces byte-identical files (fixed timestamps aren't an issue here:
unlike the DOCX corpus, nothing in this generator calls `new Date()` — the
XML is written by hand with a fixed `dc:date` value).

Each fixture's `content.xml`/`styles.xml` is hand-authored XML (not the
`docx` npm package's ODT equivalent — there wasn't one available), built
directly against the ODF constructs `odf-kit`'s reader
(`node_modules/odf-kit/dist/reader/parser.js`) actually parses:

- `tables-and-lists.odt` — a `<table:table>` with a header row and two data
  rows, plus a `<text:list>` linked to a `text:list-style` with a
  `text:list-level-style-bullet` (unordered) and one with a
  `text:list-level-style-number` (ordered).
- `embedded-image.odt` — a `<draw:frame><draw:image>` with inline
  `<office:binary-data>` (a tiny solid-color PNG from
  `scripts/lib/corpusPng.mjs`, the same helper the DOCX corpus uses) and a
  `loext:mime-type` attribute — required for odf-kit to resolve the image's
  MIME type and therefore emit an `<img src="data:...">`, since there's no
  `xlink:href` into a manifest-registered `Pictures/` entry for an
  inline-base64 image.
- `tracked-changes.odt` — a `<text:tracked-changes>` registry with one
  `text:changed-region` insertion and one deletion, referenced from the
  body via `text:change-start`/`text:change-end` (wrapping an entire
  inserted paragraph) and a `text:change` point marker (restoring an
  entire deleted paragraph). odf-kit only wraps **block-level** tracked
  changes (whole paragraphs between change markers) in a `TrackedChangeNode`
  — inline, within-a-paragraph insertions/deletions render as plain text
  with no `<ins>`/`<del>` wrapper in any mode — so this is the shape that
  actually exercises the "changes" rendering mode end to end.

  Author/date (`dc:creator`/`dc:date`) are direct children of
  `text:insertion`/`text:deletion`, not wrapped in `office:change-info` as
  real ODF 1.2 / LibreOffice output nests them — that's what odf-kit
  0.13.4's `parseChangedRegions` actually reads today (see the comment in
  `scripts/generate-odt-corpus.mjs`). If a future odf-kit version adds
  `office:change-info` support, this fixture's snapshot will need
  regenerating alongside it — exactly the kind of thing this corpus exists
  to surface.

If you add a fixture, also add its id to `FIXTURE_IDS` in
`OdtViewer.corpus.test.ts`.

## Known limitations

No LibreOffice install is available in this environment (same constraint
noted in the DOCX corpus README), so there is no reference rendering to
compare `odf-kit`'s output against — only the structural assertions and
snapshot in `OdtViewer.corpus.test.ts`. If a LibreOffice or real ODF-authoring
install becomes available, prefer replacing these hand-authored fixtures
with real LibreOffice-exported `.odt` files (which would also naturally
exercise the `office:change-info` nesting these hand-authored fixtures
intentionally do not, per the note above).
