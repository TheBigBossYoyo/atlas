// Round-trip fidelity audit (DXS round 2) — tests for the lossy-save
// detector: a real, empirical "would this save silently drop content"
// check (see lossySaveWarnings.ts's module doc for why it's a diff, not a
// hand-written blocklist).
import { describe, expect, it } from 'vitest'

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
})
