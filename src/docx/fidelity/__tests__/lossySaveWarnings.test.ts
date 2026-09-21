// Round-trip fidelity audit (DXS round 2) — tests for the lossy-save
// detector: a real, empirical "would this save silently drop content"
// check (see lossySaveWarnings.ts's module doc for why it's a diff, not a
// hand-written blocklist).
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'

import { loadDocx, saveDocx, type DocxBundle } from '../../index'
import type { Document } from '../../model/document'
import { listCorpusFixtureIds, readCorpusFixture } from '../../__tests__/corpusRoundtripHelpers'
import { categorizeLossySaveWarnings, describeLossySaveWarnings, detectLossySaveWarnings } from '../lossySaveWarnings'

import { validateOfficeFile } from '../../../../scripts/lib/officeValidator.mjs'

// `content-control-alternate-content` used to deliberately trigger two
// then-documented, then-accepted Atlas trade-offs: `w:sdt` content controls
// were unwrapped to their inner content on save (D8/DXP-08 — see
// parser/document.ts's module doc above `expandWrapperNodes`), and
// `mc:AlternateContent` kept only its `mc:Choice` branch, dropping
// `mc:Fallback` (same module, `resolveAlternateContentChildren`). Both lost
// real OOXML structure/metadata (a content control's id/alias/tag/binding;
// the legacy VML fallback shape) even though the visible content survived.
//
// Round-trip fidelity audit, DXS round 2 follow-up: `documentWriter.ts` now
// re-emits an unedited wrapper's exact source bytes instead of discarding
// it (`WrapperPassthrough` — see `../../model/document.ts`'s doc comment),
// so for THIS fixture — saved with no edits, exactly what this test does —
// nothing is lost anymore and the accepted-warnings list below is empty.
// The wrapper is still dropped (and still reported here) for a control/
// shape whose content an edit actually touches, or one sitting somewhere
// the parser doesn't capture a region for (a table cell, header, or
// footer) — this fixture doesn't exercise either case. Kept as an explicit
// list (matching `roundtrip.corpus.test.ts`'s own `KNOWN_FAILURES`
// convention) so a regression in either direction flips this assertion
// instead of silently doing nothing.
//
// B4 (attribute-value / count-delta detection) found four genuinely NEW,
// previously-invisible gaps this way — confirmed real by reading the exact
// parser/writer code responsible for each (see each entry's own comment
// below), not assumed from the warning alone. None are in files this task
// owns (`lossySaveWarnings.ts`/this test file only), so they're recorded
// here — exactly like `content-control-alternate-content`'s own pre-fix
// shape once was — rather than silently swallowed:
//
//   1. `word/numbering.xml`'s `w:abstractNum/@w15:restartNumberingAfterBreak`
//      is not modeled anywhere in `parser/numbering.ts` — every
//      `w:abstractNum` in every fixture that has one loses it on every
//      save. Low real-world severity (governs whether numbering restarts
//      after a section break), but a genuine, silent, permanent loss.
//   2. `word/numbering.xml`'s `w:lvl/@w15:tentative` — `parser/numbering.ts`
//      reads `getAttr(node, 'w:tentative')`, the WRONG namespace prefix;
//      real `.docx` files (this whole corpus included) write it as
//      `w15:tentative`, so the parser's own lookup never matches and the
//      value is silently dropped every time, under either prefix.
//   3. `word/document.xml`'s `wp:inline`'s own `distT`/`distB`/`distL`/
//      `distR` (space-around-image) attributes aren't modeled at all —
//      `buildDrawingNode` (documentWriter.ts) only builds attributes for
//      `wp:anchor` (`buildAnchorAttributes`), passing `undefined` for
//      `wp:inline`. Harmless in this corpus (every source value happens to
//      be the OOXML default, `"0"`), but a real inline image with an
//      explicit nonzero distance would silently lose it — latent, not yet
//      exercised by a fixture with a nonzero value.
//   4. `word/document.xml`'s `wp:anchor`'s own `distT`/`distB`/`distL`/
//      `distR`/`simplePos`/`relativeHeight`/`locked` are HARDCODED constants
//      in `buildAnchorAttributes` (documentWriter.ts — `attributes['@_...'] =
//      '0'`/`'1'` literals), never read from the source at all.
//      `relativeHeight` is the most consequential: it's the floating
//      image's Z-ORDER. Two anchored images with different stacking order
//      in the source both collapse to `relativeHeight="0"` on save,
//      silently flattening their front-to-back order. `distT`/`distR`/
//      `simplePos`/`locked` aren't listed as separate warnings below only
//      because this fixture's own source values happen to already be `"0"`
//      — same latent-vs-exercised distinction as (3).
//   5. `word/document.xml`'s `a:blip/@cstate` (image compression hint) and
//      `pic:spPr/@bwMode` (black-and-white print-preview hint) are never
//      emitted by `buildGraphicNode`/the shape-properties builder at all —
//      purely cosmetic rendering hints, lowest real-world severity of the
//      five, but still a genuine, silent, unconditional loss.
//   6. `word/document.xml`'s `w:tblW/@w:w` when `w:tblW/@w:type="auto"` (or
//      `"nil"`) — this one is DIFFERENT from the five above: it is
//      confirmed SAFE, not a real gap. `parseWidth` (parser/document.ts)
//      deliberately sets `value` to `undefined` whenever `type` is `auto`/
//      `nil`, because OOXML itself defines `w:w` as meaningless/ignored by
//      every consumer in that case (`type="dxa"`/`"pct"`, where the value
//      DOES matter, round-trip the literal number exactly as before — this
//      module would still catch a real loss there). Listed here anyway,
//      rather than added as a blanket `w:tblW`/`@w:w` identity exclusion in
//      the detector itself, because this module has no way to verify
//      "`type` was `auto`/`nil` for the SAME occurrence" without tracking
//      per-occurrence attribute correlation it deliberately doesn't do (see
//      `lossySaveWarnings.ts`'s module doc) — excluding the attribute name
//      outright would also hide a real `type="dxa"` loss, which is not an
//      acceptable trade. A per-fixture accepted entry, verified by reading
//      `parseWidth`, is the honest way to record "this specific instance is
//      fine" without weakening the general rule.
const NUMBERING_TENTATIVE_AND_RESTART_TAGS: ReadonlyArray<{ readonly tag: string }> = [
  { tag: 'w:abstractNum' },
  { tag: 'w:lvl' },
]
const KNOWN_ACCEPTED_WARNINGS: Partial<Record<string, ReadonlyArray<{ readonly tag: string }>>> = {
  'image-anchored-floating': [
    { tag: 'wp:anchor' },
    { tag: 'wp:anchor' },
    { tag: 'wp:anchor' },
    { tag: 'a:blip' },
    { tag: 'pic:spPr' },
  ],
  'image-crop-rotation-flip': [
    { tag: 'wp:inline' },
    { tag: 'wp:inline' },
    { tag: 'wp:inline' },
    { tag: 'wp:inline' },
    { tag: 'a:blip' },
    { tag: 'pic:spPr' },
  ],
  'image-emf-wmf': [
    { tag: 'wp:inline' },
    { tag: 'wp:inline' },
    { tag: 'wp:inline' },
    { tag: 'wp:inline' },
    { tag: 'a:blip' },
    { tag: 'pic:spPr' },
  ],
  'image-inline': [
    { tag: 'wp:inline' },
    { tag: 'wp:inline' },
    { tag: 'wp:inline' },
    { tag: 'wp:inline' },
    { tag: 'a:blip' },
    { tag: 'pic:spPr' },
  ],
  'table-fixed-grid-merged-cells': [{ tag: 'w:tblW' }],
  'table-styled-banded': [{ tag: 'w:tblW' }],
}

describe('detectLossySaveWarnings', () => {
  it('reports no warnings for any corpus fixture, saved with no edits, beyond the known-accepted ones', async () => {
    const fixtureIds = await listCorpusFixtureIds()
    expect(fixtureIds.length).toBeGreaterThan(0)

    for (const fixtureId of fixtureIds) {
      const buffer = await readCorpusFixture(fixtureId)
      const bundle = await loadDocx(buffer)
      const saved = await saveDocx(bundle)

      const warnings = await detectLossySaveWarnings(bundle, saved)
      // Every fixture in this corpus has a `word/numbering.xml` part
      // exhibiting the same two real, unmodeled-attribute gaps (see the
      // long comment above), plus whatever fixture-specific ones apply.
      const expectedWarnings = [...NUMBERING_TENTATIVE_AND_RESTART_TAGS, ...(KNOWN_ACCEPTED_WARNINGS[fixtureId] ?? [])]
      const expectedTags = expectedWarnings.map((w) => w.tag)
      expect(
        warnings.map((w) => w.tag).sort(),
        `${fixtureId} should round-trip with no detected fidelity loss beyond the known-accepted set`,
      ).toEqual([...expectedTags].sort())
    }
  })

  it('detects a genuinely unmodeled element disappearing from word/document.xml on save', async () => {
    const buffer = await readCorpusFixture('plain-paragraphs-styles')
    const bundle = await loadDocx(buffer)
    const decoder = new TextDecoder()
    const originalXml = decoder.decode(bundle.rawArchive?.get('word/document.xml'))

    // `w:printerSettings` (a reference to a printer-settings binary part) is
    // not modeled anywhere in Atlas's parser/serializer — splice one into
    // the loaded archive's own copy of word/document.xml to prove the
    // detector actually catches an unmodeled element rather than always
    // reporting clean.
    const withPrinterSettings = originalXml.replace(
      '<w:sectPr>',
      '<w:sectPr><w:printerSettings r:id="rIdPrinterSettings"/>',
    )
    expect(withPrinterSettings).not.toEqual(originalXml)

    const mutatedRawArchive = new Map(bundle.rawArchive)
    mutatedRawArchive.set('word/document.xml', new TextEncoder().encode(withPrinterSettings))
    const mutatedBundle = { ...bundle, rawArchive: mutatedRawArchive }

    // A real no-op save: `saveDocx` always rewrites word/document.xml from
    // `bundle.document` (parsed from the ORIGINAL, unmutated xml), so the
    // saved output has no way to carry the spliced-in element through.
    const saved = await saveDocx(mutatedBundle)

    const warnings = await detectLossySaveWarnings(mutatedBundle, saved)
    // Scoped to word/document.xml: this fixture also carries B4's two
    // genuine, unrelated `word/numbering.xml` gaps (see
    // `KNOWN_ACCEPTED_WARNINGS`'s own doc comment in the corpus test
    // above), which would otherwise show up here too.
    const documentWarnings = warnings.filter((w) => w.part === 'word/document.xml')
    expect(documentWarnings).toEqual([{ part: 'word/document.xml', tag: 'w:printerSettings', kind: 'element', occurrences: 1 }])
    expect(describeLossySaveWarnings(documentWarnings)).toEqual([
      'word/document.xml: w:printerSettings (1 occurrence) is not supported and will be removed',
    ])
  })

  // Round-trip fidelity audit, DXS round 2 follow-up — the flip side of the
  // "no warnings for an unedited save" test above: once an edit invalidates
  // a wrapper-passthrough region's identity check (`WrapperPassthrough` —
  // see `../../model/document.ts`'s doc comment), the wrapper falls back to
  // being stripped exactly as it always was, and the detector must still
  // catch that and categorize it in non-technical, user-facing terms.
  it('reports content-control and shape-fallback categories once an edit invalidates their wrapper-passthrough regions', async () => {
    const buffer = await readCorpusFixture('content-control-alternate-content')
    const bundle = await loadDocx(buffer)

    // `structuredClone` rebuilds every Section/Block/ParagraphChild/RunChild
    // object with a fresh identity while leaving their content untouched —
    // standing in for "the user edited this" without needing to know which
    // paragraph/run index the fixture's content control and shape live at.
    // `bundle.document.wrappers` (left untouched here) still points at the
    // ORIGINAL objects, so this is exactly what a real edit invalidating the
    // region's identity check looks like from `documentWriter.ts`'s side.
    const editedDocument = { ...bundle.document, sections: structuredClone(bundle.document.sections) }
    const saved = await saveDocx({ ...bundle, document: editedDocument })

    const warnings = await detectLossySaveWarnings(bundle, saved)
    // Scoped to word/document.xml: this corpus fixture — like every other
    // one (see `KNOWN_ACCEPTED_WARNINGS`'s own doc comment above) — also
    // carries B4's two genuine, unrelated `word/numbering.xml` gaps
    // (`w:abstractNum`/`w15:restartNumberingAfterBreak`,
    // `w:lvl`/`w15:tentative`), which would otherwise add an unrelated
    // `'other'` to this assertion and obscure what this test actually
    // checks: the wrapper-passthrough fallback's own categorization.
    const documentWarnings = warnings.filter((w) => w.part === 'word/document.xml')
    expect(categorizeLossySaveWarnings(documentWarnings)).toEqual(new Set(['content-control', 'shape-fallback']))

    // The "edit" was purely structural — proving the fallback (stripping
    // the wrapper, same as Atlas has always done) still produces a package
    // Office accepts, not a corrupted file.
    const result = validateOfficeFile(Buffer.from(saved))
    expect(result.format).toEqual({ family: 'opc', kind: 'docx' })
    expect(result.issues.filter((issue) => issue.severity === 'error')).toEqual([])
  })

  // DOCX-2: footnotes/endnotes/comments previously weren't checked at all
  // (see `NOTES_AND_COMMENTS_PARTS`'s doc comment in lossySaveWarnings.ts),
  // so a `w:sdt`/`mc:AlternateContent` wrapper inside one of them was
  // stripped on every save with no warning whatsoever — unlike everywhere
  // else in the document. `partBody.ts`'s shared parsing helper for these
  // three parts doesn't thread `document.wrappers` through to the writers
  // the way `documentWriter.ts` does for the main body, so the wrapper is
  // always dropped here regardless of whether anything was edited — no
  // `structuredClone`-style identity-busting needed, unlike the body test
  // above.
  describe('DOCX-2 — footnotes/endnotes/comments wrapper warnings', () => {
    it('detects a w:sdt content control dropped from word/footnotes.xml', async () => {
      const buffer = await readCorpusFixture('footnotes-endnotes')
      const bundle = await loadDocx(buffer)
      const decoder = new TextDecoder()
      const originalFootnotesXml = decoder.decode(bundle.rawArchive?.get('word/footnotes.xml'))

      // Wrap the one real (non-separator) footnote's paragraph in a
      // content control — `saveDocx` regenerates word/footnotes.xml purely
      // from `bundle.document.footnotes` (parsed from the UNMUTATED xml
      // below), so the splice below only affects what the detector diffs
      // against, exactly mirroring the `w:printerSettings` splice test
      // above for word/document.xml.
      const withSdt = originalFootnotesXml.replace(
        '<w:footnote w:id="1"><w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
          + '<w:footnoteRef/></w:r><w:r><w:t xml:space="preserve">This is the footnote text.</w:t></w:r></w:p></w:footnote>',
        '<w:footnote w:id="1"><w:sdt><w:sdtPr><w:id w:val="123456789"/></w:sdtPr><w:sdtContent>'
          + '<w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
          + '<w:r><w:t xml:space="preserve">This is the footnote text.</w:t></w:r></w:p>'
          + '</w:sdtContent></w:sdt></w:footnote>',
      )
      expect(withSdt).not.toEqual(originalFootnotesXml)

      const mutatedRawArchive = new Map(bundle.rawArchive)
      mutatedRawArchive.set('word/footnotes.xml', new TextEncoder().encode(withSdt))
      const mutatedBundle = { ...bundle, rawArchive: mutatedRawArchive }

      const saved = await saveDocx(mutatedBundle)
      const warnings = await detectLossySaveWarnings(mutatedBundle, saved)

      // Scoped to word/footnotes.xml for both assertions below: this
      // fixture also carries B4's two genuine, unrelated
      // `word/numbering.xml` gaps (see `KNOWN_ACCEPTED_WARNINGS`'s own doc
      // comment in the corpus test above), which would otherwise add an
      // unrelated `'other'` category and unrelated tags here.
      const footnotesWarnings = warnings.filter((w) => w.part === 'word/footnotes.xml')
      expect(footnotesWarnings.map((w) => w.tag).sort()).toEqual(['w:id', 'w:sdt', 'w:sdtContent', 'w:sdtPr'].sort())
      expect(categorizeLossySaveWarnings(footnotesWarnings)).toEqual(new Set(['content-control']))
    })

    it('detects an mc:AlternateContent shape fallback dropped from word/endnotes.xml', async () => {
      const buffer = await readCorpusFixture('footnotes-endnotes')
      const bundle = await loadDocx(buffer)
      const decoder = new TextDecoder()
      const originalEndnotesXml = decoder.decode(bundle.rawArchive?.get('word/endnotes.xml'))

      const withAlternateContent = originalEndnotesXml.replace(
        '<w:endnote w:id="1"><w:p><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr>'
          + '<w:endnoteRef/></w:r><w:r><w:t xml:space="preserve">This is the endnote text.</w:t></w:r></w:p></w:endnote>',
        '<w:endnote w:id="1"><mc:AlternateContent><mc:Choice Requires="wps">'
          + '<w:p><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>'
          + '<w:r><w:t xml:space="preserve">This is the endnote text.</w:t></w:r></w:p></mc:Choice>'
          + '<mc:Fallback><w:p><w:r><w:t xml:space="preserve">Fallback text.</w:t></w:r></w:p></mc:Fallback>'
          + '</mc:AlternateContent></w:endnote>',
      )
      expect(withAlternateContent).not.toEqual(originalEndnotesXml)

      const mutatedRawArchive = new Map(bundle.rawArchive)
      mutatedRawArchive.set('word/endnotes.xml', new TextEncoder().encode(withAlternateContent))
      const mutatedBundle = { ...bundle, rawArchive: mutatedRawArchive }

      const saved = await saveDocx(mutatedBundle)
      const warnings = await detectLossySaveWarnings(mutatedBundle, saved)

      // B4 catches more here than the wrapper tags alone: `partBody.ts`'s
      // shared parser keeps only the `mc:Choice` branch's content (see this
      // `describe` block's own doc comment / `resolveAlternateContentChildren`),
      // so the `mc:Fallback` branch's OWN paragraph/run/text ("Fallback
      // text.") is genuinely, silently discarded too — not just the
      // `mc:Fallback` wrapper tag itself. A tag-presence-only check (this
      // file's pre-B4 shape) couldn't see this: `w:p`/`w:r`/`w:t` all still
      // appear elsewhere in this same endnote (the surviving `mc:Choice`
      // branch has its own paragraph/run/text), so their COUNT merely drops
      // (2 -> 1 for each), which is exactly the count-delta capability this
      // task added.
      const endnotesWarnings = warnings.filter((w) => w.part === 'word/endnotes.xml')
      expect(endnotesWarnings.map((w) => w.tag).sort()).toEqual(
        ['mc:AlternateContent', 'mc:Choice', 'mc:Fallback', 'w:p', 'w:r', 'w:t'].sort(),
      )
      expect(categorizeLossySaveWarnings(endnotesWarnings)).toEqual(new Set(['shape-fallback', 'other']))
    })

    it('detects a w:sdt content control dropped from word/comments.xml', async () => {
      const buffer = await readCorpusFixture('comments-with-reply')
      const bundle = await loadDocx(buffer)
      const decoder = new TextDecoder()
      const originalCommentsXml = decoder.decode(bundle.rawArchive?.get('word/comments.xml'))

      const withSdt = originalCommentsXml.replace(
        '<w:comment w:id="1" w:initials="R1" w:author="Reviewer One" w:date="2020-01-01T00:00:00.000Z">'
          + '<w:p><w:r><w:t xml:space="preserve">This phrase needs clarification.</w:t></w:r></w:p></w:comment>',
        '<w:comment w:id="1" w:initials="R1" w:author="Reviewer One" w:date="2020-01-01T00:00:00.000Z">'
          + '<w:sdt><w:sdtPr><w:id w:val="987654321"/></w:sdtPr><w:sdtContent>'
          + '<w:p><w:r><w:t xml:space="preserve">This phrase needs clarification.</w:t></w:r></w:p>'
          + '</w:sdtContent></w:sdt></w:comment>',
      )
      expect(withSdt).not.toEqual(originalCommentsXml)

      const mutatedRawArchive = new Map(bundle.rawArchive)
      mutatedRawArchive.set('word/comments.xml', new TextEncoder().encode(withSdt))
      const mutatedBundle = { ...bundle, rawArchive: mutatedRawArchive }

      const saved = await saveDocx(mutatedBundle)
      const warnings = await detectLossySaveWarnings(mutatedBundle, saved)

      // Scoped to word/comments.xml — see the footnotes test above for why
      // (same unrelated B4 `word/numbering.xml` gaps this fixture also has).
      const commentsWarnings = warnings.filter((w) => w.part === 'word/comments.xml')
      expect(commentsWarnings.map((w) => w.tag).sort()).toEqual(['w:id', 'w:sdt', 'w:sdtContent', 'w:sdtPr'].sort())
      expect(categorizeLossySaveWarnings(commentsWarnings)).toEqual(new Set(['content-control']))
    })

    // Full end-to-end proof, distinct from the three detection tests above:
    // those splice the wrapper into `rawArchive` WITHOUT reparsing, so the
    // in-memory model never sees it and the detector's warning is really
    // just proving detection works independent of preservation (e.g. for an
    // edited note, or a wrapper shape the parser can't capture a region
    // for). This one loads a real package whose footnotes.xml genuinely
    // contains the `w:sdt` — so `bundle.document.footnotes` captures the
    // wrapper region via `partBody.ts`'s parse — and saves with no edits,
    // proving `footnotesWriter.ts` now round-trips it and the detector
    // correctly reports nothing lost.
    it('round-trips an unedited w:sdt footnote through the full load/save pipeline with no warning (DOCX-2)', async () => {
      const buffer = await readCorpusFixture('footnotes-endnotes')
      const zip = await JSZip.loadAsync(buffer)
      const originalFootnotesXml = await zip.file('word/footnotes.xml')?.async('string')
      expect(originalFootnotesXml).toBeDefined()

      const withSdt = (originalFootnotesXml as string).replace(
        '<w:footnote w:id="1"><w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
          + '<w:footnoteRef/></w:r><w:r><w:t xml:space="preserve">This is the footnote text.</w:t></w:r></w:p></w:footnote>',
        '<w:footnote w:id="1"><w:sdt><w:sdtPr><w:id w:val="123456789"/></w:sdtPr><w:sdtContent>'
          + '<w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
          + '<w:r><w:t xml:space="preserve">This is the footnote text.</w:t></w:r></w:p>'
          + '</w:sdtContent></w:sdt></w:footnote>',
      )
      expect(withSdt).not.toEqual(originalFootnotesXml)
      zip.file('word/footnotes.xml', withSdt)
      const mutatedBuffer = await zip.generateAsync({ type: 'arraybuffer' })

      const bundle = await loadDocx(mutatedBuffer)
      const saved = await saveDocx(bundle)
      const warnings = await detectLossySaveWarnings(bundle, saved)

      expect(warnings.filter((w) => w.part === 'word/footnotes.xml')).toEqual([])

      const savedZip = await JSZip.loadAsync(saved)
      const savedFootnotesXml = await savedZip.file('word/footnotes.xml')?.async('string')
      expect(savedFootnotesXml).toContain(
        '<w:sdt><w:sdtPr><w:id w:val="123456789"/></w:sdtPr><w:sdtContent>'
          + '<w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
          + '<w:r><w:t xml:space="preserve">This is the footnote text.</w:t></w:r></w:p>'
          + '</w:sdtContent></w:sdt>',
      )

      const result = validateOfficeFile(Buffer.from(saved))
      expect(result.format).toEqual({ family: 'opc', kind: 'docx' })
      expect(result.issues.filter((issue) => issue.severity === 'error')).toEqual([])
    })

    it('reports no warning for a document with none of comments/footnotes/endnotes content', async () => {
      // This is the real-world shape of the trap, not a hypothetical: this
      // fixture's `word/comments.xml` is present in the raw archive — a
      // self-closing `<w:comments/>` most real-world DOCX authoring tools
      // emit even for a document with zero comments — while the model's
      // `comments` map is empty, and it has no footnotes.xml/endnotes.xml
      // at all (most documents never do).
      const buffer = await readCorpusFixture('plain-paragraphs-styles')
      const bundle = await loadDocx(buffer)
      expect(bundle.rawArchive?.has('word/comments.xml')).toBe(true)
      expect(bundle.document.comments.size).toBe(0)
      expect(bundle.rawArchive?.has('word/footnotes.xml')).toBe(false)
      expect(bundle.rawArchive?.has('word/endnotes.xml')).toBe(false)

      const saved = await saveDocx(bundle)
      const warnings = await detectLossySaveWarnings(bundle, saved)

      expect(warnings.filter((w) => ['word/comments.xml', 'word/footnotes.xml', 'word/endnotes.xml'].includes(w.part))).toEqual(
        [],
      )
    })

    // The trap this test guards against: naively diffing "original part"
    // against "saved part" for these three reads as a false positive the
    // instant the model's collection legitimately empties out (the user
    // deleted every footnote/comment) while the raw-archive copy of the
    // part is still sitting there from the original load — see
    // `NOTES_AND_COMMENTS_PARTS`'s doc comment for why the module's two
    // existing "part legitimately absent" guards already cover this
    // without any part-type-specific code.
    it('reports no warning when the model empties a footnotes collection that still has a raw-archive part', async () => {
      const buffer = await readCorpusFixture('footnotes-endnotes')
      const bundle = await loadDocx(buffer)
      expect(bundle.rawArchive?.get('word/footnotes.xml')).toBeDefined()
      expect(bundle.document.footnotes.size).toBeGreaterThan(0)

      // Simulate "the user deleted every footnote": the model's collection
      // goes to zero, but `bundle.rawArchive`'s word/footnotes.xml (from
      // the original load) is untouched — `saveDocx` leaves the part as a
      // raw-archive passthrough rather than writing anything for it (see
      // `docx/index.ts`'s "6. Footnotes / endnotes" step), so it stays
      // byte-identical to the original.
      const editedDocument = { ...bundle.document, footnotes: new Map() }
      const editedBundle = { ...bundle, document: editedDocument }

      const saved = await saveDocx(editedBundle)
      const warnings = await detectLossySaveWarnings(editedBundle, saved)

      expect(warnings.filter((w) => w.part === 'word/footnotes.xml')).toEqual([])
    })

    it('reports no warning when the model empties a comments collection that still has a raw-archive part', async () => {
      const buffer = await readCorpusFixture('comments-with-reply')
      const bundle = await loadDocx(buffer)
      expect(bundle.rawArchive?.get('word/comments.xml')).toBeDefined()
      expect(bundle.document.comments.size).toBeGreaterThan(0)

      // Simulate "the user deleted every comment": unlike footnotes/
      // endnotes, `saveDocx` actively removes word/comments.xml from the
      // saved package in this case (`removePartRegistration`) rather than
      // leaving it as a passthrough — the part vanishes from the saved zip
      // entirely.
      const editedDocument = { ...bundle.document, comments: new Map() }
      const editedBundle = { ...bundle, document: editedDocument }

      const saved = await saveDocx(editedBundle)
      const warnings = await detectLossySaveWarnings(editedBundle, saved)

      expect(warnings.filter((w) => w.part === 'word/comments.xml')).toEqual([])
    })
  })

  // B4 — attribute-value and occurrence-count-delta detection. Unlike the
  // tests above, these construct the "original" and "saved" `word/document.xml`
  // XML directly (via a hand-built `DocxBundle`/zip, never `loadDocx`/
  // `saveDocx`) rather than going through the real pipeline: the two bugs
  // this section regression-checks (DOCX-1, DOCX-12) are already fixed on
  // `main`, so the real pipeline can no longer reproduce their pre-fix
  // shapes — exactly the instruction this task was given. Every other test
  // in this `describe` block still exercises the real pipeline as before.
  describe('B4 — attribute-value and count-delta detection', () => {
    /** A `DocxBundle` whose `word/document.xml` raw-archive entry is exactly `originalXml` — nothing else in `DocxBundle` matters to `detectLossySaveWarnings` (it only reads `rawArchive` and `relationships`), so `document` is a harmless stand-in, never actually read. */
    function bundleWithOriginalDocumentXml(originalXml: string): DocxBundle {
      return {
        document: {} as unknown as Document,
        rawArchive: new Map([['word/document.xml', new TextEncoder().encode(originalXml)]]),
      }
    }

    /** The `Uint8Array` a hand-built "saved" package would be — a minimal zip whose only relevant entry is `word/document.xml` (all `detectLossySaveWarnings` ever reads from `savedBytes`). */
    async function savedBytesWithDocumentXml(savedXml: string): Promise<Uint8Array> {
      const zip = new JSZip()
      zip.file('word/document.xml', savedXml)
      return zip.generateAsync({ type: 'uint8array' })
    }

    it('DOCX-1 regression: a w:rFonts that had w:asciiTheme on load and emits only literal attributes on save is flagged', async () => {
      // Pre-fix shape (see 65b068c's commit message): `buildFontSetElement`
      // emitted ONLY `w:ascii`/`w:hAnsi`/`w:cs`/`w:eastAsia`/`w:hint`, never
      // the four theme attributes — so a `w:rFonts` carrying BOTH a literal
      // font and a theme reference kept its literal attribute (the element
      // itself never disappears, so the pre-B4 tag-presence-only check saw
      // nothing wrong) while silently losing the theme one.
      const original = bundleWithOriginalDocumentXml(
        '<w:document><w:body><w:p><w:r><w:rPr>'
          + '<w:rFonts w:ascii="Calibri" w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>'
          + '</w:rPr><w:t>Hi</w:t></w:r></w:p></w:body></w:document>',
      )
      const saved = await savedBytesWithDocumentXml(
        '<w:document><w:body><w:p><w:r><w:rPr>'
          + '<w:rFonts w:ascii="Calibri"/>'
          + '</w:rPr><w:t>Hi</w:t></w:r></w:p></w:body></w:document>',
      )

      const warnings = await detectLossySaveWarnings(original, saved)

      expect(warnings).toEqual(
        expect.arrayContaining([
          { part: 'word/document.xml', tag: 'w:rFonts', kind: 'attribute', attribute: 'w:asciiTheme', occurrences: 1 },
          { part: 'word/document.xml', tag: 'w:rFonts', kind: 'attribute', attribute: 'w:hAnsiTheme', occurrences: 1 },
        ]),
      )
      // And NOT as a `w:rFonts` element-level warning — the element (and
      // its surviving `w:ascii` attribute) is still there; only DOCX-1's
      // specific two theme attributes are gone. Confirms this really is the
      // attribute-value check catching it, not a count/presence fluke.
      expect(warnings.filter((w) => w.tag === 'w:rFonts' && w.kind === 'element')).toEqual([])
    })

    it('DOCX-12 regression: a run whose w:rPr had w:outline/w:em/w:effect and lost them is flagged', async () => {
      // Pre-fix shape (65b068c): none of `w:outline` (a named toggle),
      // `w:em` (a named value element), or `w:effect` (an unmodeled CT_RPr
      // member with no field OR general passthrough yet) had any writer
      // path at all — `buildRunPropertiesNode` simply never emitted them,
      // regardless of whether they were edited.
      const original = bundleWithOriginalDocumentXml(
        '<w:document><w:body><w:p><w:r><w:rPr>'
          + '<w:b/><w:outline/><w:em w:val="dot"/><w:effect w:val="sparkle"/>'
          + '</w:rPr><w:t>Hi</w:t></w:r></w:p></w:body></w:document>',
      )
      const saved = await savedBytesWithDocumentXml(
        '<w:document><w:body><w:p><w:r><w:rPr>'
          + '<w:b/>'
          + '</w:rPr><w:t>Hi</w:t></w:r></w:p></w:body></w:document>',
      )

      const warnings = await detectLossySaveWarnings(original, saved)

      expect(warnings).toEqual(
        expect.arrayContaining([
          { part: 'word/document.xml', tag: 'w:outline', kind: 'element', occurrences: 1 },
          { part: 'word/document.xml', tag: 'w:em', kind: 'element', occurrences: 1 },
          { part: 'word/document.xml', tag: 'w:effect', kind: 'element', occurrences: 1 },
        ]),
      )
      // This particular pre-fix shape happens to be catchable by the
      // ORIGINAL (pre-B4) tag-presence-only check too — each of these three
      // element names has exactly one occurrence in the whole part, so
      // losing it drops the tag's count to zero either way. B4 doesn't
      // change the VERDICT for this exact shape; it changes what happens
      // when a document has ANOTHER occurrence of the same tag elsewhere
      // (DOCX-1's own `w:rFonts` case above is the shape where that
      // distinction actually matters) — recorded here as the honest answer
      // to "does the detector catch this", not to claim it as new coverage.
      expect(warnings.length).toBeGreaterThanOrEqual(3)
    })

    it('flags an element whose occurrence count drops without disappearing entirely (12 w:tab -> 1)', async () => {
      const manyTabs = Array.from({ length: 12 }, () => '<w:tab/>').join('')
      const original = bundleWithOriginalDocumentXml(
        `<w:document><w:body><w:p><w:r>${manyTabs}<w:t>Hi</w:t></w:r></w:p></w:body></w:document>`,
      )
      const saved = await savedBytesWithDocumentXml(
        '<w:document><w:body><w:p><w:r><w:tab/><w:t>Hi</w:t></w:r></w:p></w:body></w:document>',
      )

      const warnings = await detectLossySaveWarnings(original, saved)

      expect(warnings).toEqual([
        { part: 'word/document.xml', tag: 'w:tab', kind: 'element', occurrences: 12, savedOccurrences: 1 },
      ])
      expect(describeLossySaveWarnings(warnings)).toEqual([
        'word/document.xml: w:tab (12 occurrences dropped to 1) is not fully supported and some will be removed',
      ])
    })

    it('does not flag an element whose occurrence count increases (a real edit adding content)', async () => {
      const original = bundleWithOriginalDocumentXml(
        '<w:document><w:body><w:p><w:r><w:tab/><w:t>Hi</w:t></w:r></w:p></w:body></w:document>',
      )
      const manyTabs = Array.from({ length: 12 }, () => '<w:tab/>').join('')
      const saved = await savedBytesWithDocumentXml(
        `<w:document><w:body><w:p><w:r>${manyTabs}<w:t>Hi</w:t></w:r></w:p></w:body></w:document>`,
      )

      const warnings = await detectLossySaveWarnings(original, saved)

      expect(warnings).toEqual([])
    })

    it('flags a changed attribute value even when the attribute name and element count are unchanged', async () => {
      const original = bundleWithOriginalDocumentXml(
        '<w:document><w:body><w:p><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>Hi</w:t></w:r></w:p></w:body></w:document>',
      )
      const saved = await savedBytesWithDocumentXml(
        '<w:document><w:body><w:p><w:r><w:rPr><w:color w:val="00FF00"/></w:rPr><w:t>Hi</w:t></w:r></w:p></w:body></w:document>',
      )

      const warnings = await detectLossySaveWarnings(original, saved)

      expect(warnings).toEqual([
        { part: 'word/document.xml', tag: 'w:color', kind: 'attribute', attribute: 'w:val', occurrences: 1 },
      ])
    })

    it('does not flag an attribute value that merely moved to a different occurrence of the same tag', async () => {
      // Two `w:pStyle`s swap values — legitimate churn (re-serialization
      // visiting paragraphs in a different order, or one edit + one
      // untouched paragraph elsewhere already using the "new" value) this
      // module's own doc comment names as the deliberate precision-over-
      // recall trade-off: every value that existed in the original still
      // exists SOMEWHERE on the tag, so nothing warns.
      const original = bundleWithOriginalDocumentXml(
        '<w:document><w:body>'
          + '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:p>'
          + '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr></w:p>'
          + '</w:body></w:document>',
      )
      const saved = await savedBytesWithDocumentXml(
        '<w:document><w:body>'
          + '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr></w:p>'
          + '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:p>'
          + '</w:body></w:document>',
      )

      const warnings = await detectLossySaveWarnings(original, saved)

      expect(warnings).toEqual([])
    })

    it('does not flag identity-ish attribute churn: renumbered rsids/ids, a namespace-prefix correction, or an ST_OnOff/ST_Jc synonym', async () => {
      const original = bundleWithOriginalDocumentXml(
        '<w:document mc:Ignorable="w14 wp14"><w:body>'
          + '<w:p w:rsidR="00AA0000" w:rsidRDefault="00AA0000"><w:pPr><w:jc w:val="left"/></w:pPr>'
          + '<w:bookmarkStart w:id="3" w:name="X"/><w:bookmarkEnd w:id="3"/>'
          + '<w:r><w:rPr><w:rtl w:val="true"/></w:rPr><w:t>Hi</w:t></w:r></w:p>'
          + '</w:body></w:document>',
      )
      const saved = await savedBytesWithDocumentXml(
        '<w:document mc:Ignorable="wp14 w14 w15"><w:body>'
          + '<w:p w:rsidR="00BB1111" w:rsidRDefault="00BB1111"><w:pPr><w:jc w:val="start"/></w:pPr>'
          + '<w:bookmarkStart w:id="7" w:name="X"/><w:bookmarkEnd w:id="7"/>'
          + '<w:r><w:rPr><w:rtl w:val="1"/></w:rPr><w:t>Hi</w:t></w:r></w:p>'
          + '</w:body></w:document>',
      )

      const warnings = await detectLossySaveWarnings(original, saved)

      expect(warnings).toEqual([])
    })

    it('does not flag w:vMerge/@w:val dropping when it means the schema-default "continue"', async () => {
      const original = bundleWithOriginalDocumentXml(
        '<w:document><w:body><w:tbl>'
          + '<w:tr><w:tc><w:tcPr><w:vMerge w:val="continue"/></w:tcPr></w:tc></w:tr>'
          + '</w:tbl></w:body></w:document>',
      )
      const saved = await savedBytesWithDocumentXml(
        '<w:document><w:body><w:tbl>'
          + '<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr></w:tc></w:tr>'
          + '</w:tbl></w:body></w:document>',
      )

      const warnings = await detectLossySaveWarnings(original, saved)

      expect(warnings).toEqual([])
    })
  })
})
