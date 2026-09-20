// Round-trip fidelity audit (DXS round 2) — tests for the lossy-save
// detector: a real, empirical "would this save silently drop content"
// check (see lossySaveWarnings.ts's module doc for why it's a diff, not a
// hand-written blocklist).
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'

import { loadDocx, saveDocx } from '../../index'
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
const KNOWN_ACCEPTED_WARNINGS: Partial<Record<string, ReadonlyArray<{ readonly tag: string }>>> = {}

describe('detectLossySaveWarnings', () => {
  it('reports no warnings for any corpus fixture, saved with no edits, beyond the known-accepted ones', async () => {
    const fixtureIds = await listCorpusFixtureIds()
    expect(fixtureIds.length).toBeGreaterThan(0)

    for (const fixtureId of fixtureIds) {
      const buffer = await readCorpusFixture(fixtureId)
      const bundle = await loadDocx(buffer)
      const saved = await saveDocx(bundle)

      const warnings = await detectLossySaveWarnings(bundle, saved)
      const expectedTags = (KNOWN_ACCEPTED_WARNINGS[fixtureId] ?? []).map((w) => w.tag)
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
    expect(warnings).toEqual([{ part: 'word/document.xml', tag: 'w:printerSettings', occurrences: 1 }])
    expect(describeLossySaveWarnings(warnings)).toEqual([
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
    expect(categorizeLossySaveWarnings(warnings)).toEqual(new Set(['content-control', 'shape-fallback']))

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

      expect(warnings.filter((w) => w.part === 'word/footnotes.xml').map((w) => w.tag).sort()).toEqual(
        ['w:id', 'w:sdt', 'w:sdtContent', 'w:sdtPr'].sort(),
      )
      expect(categorizeLossySaveWarnings(warnings)).toEqual(new Set(['content-control']))
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

      expect(warnings.filter((w) => w.part === 'word/endnotes.xml').map((w) => w.tag).sort()).toEqual(
        ['mc:AlternateContent', 'mc:Choice', 'mc:Fallback'].sort(),
      )
      expect(categorizeLossySaveWarnings(warnings)).toEqual(new Set(['shape-fallback']))
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

      expect(warnings.filter((w) => w.part === 'word/comments.xml').map((w) => w.tag).sort()).toEqual(
        ['w:id', 'w:sdt', 'w:sdtContent', 'w:sdtPr'].sort(),
      )
      expect(categorizeLossySaveWarnings(warnings)).toEqual(new Set(['content-control']))
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
})
