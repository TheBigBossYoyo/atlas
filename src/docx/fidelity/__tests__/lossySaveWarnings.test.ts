// Round-trip fidelity audit (DXS round 2) — tests for the lossy-save
// detector: a real, empirical "would this save silently drop content"
// check (see lossySaveWarnings.ts's module doc for why it's a diff, not a
// hand-written blocklist).
import { describe, expect, it } from 'vitest'

import { loadDocx, saveDocx } from '../../index'
import { listCorpusFixtureIds, readCorpusFixture } from '../../__tests__/corpusRoundtripHelpers'
import { describeLossySaveWarnings, detectLossySaveWarnings } from '../lossySaveWarnings'

// `content-control-alternate-content` deliberately triggers two
// already-documented, already-decided Atlas trade-offs (not bugs this task
// introduced or is fixing): `w:sdt` content controls are unwrapped to their
// inner content on save (D8/DXP-08 — see parser/document.ts's module doc
// above `expandWrapperNodes`), and `mc:AlternateContent` keeps only its
// `mc:Choice` branch, dropping `mc:Fallback` (same module, `resolveAlternate
// ContentChildren`). Both lose real OOXML structure/metadata (a content
// control's id/alias/tag/binding; the legacy VML fallback shape) even
// though the visible content survives — exactly the class of "known,
// visible-if-you-look-for-it limitation" this detector exists to surface.
// Tracked here explicitly (matching `roundtrip.corpus.test.ts`'s own
// `KNOWN_FAILURES` convention) so a change to either behavior — a fix, or a
// regression — flips this assertion instead of silently doing nothing.
const KNOWN_ACCEPTED_WARNINGS: Partial<Record<string, ReadonlyArray<{ readonly tag: string }>>> = {
  'content-control-alternate-content': [
    { tag: 'w:sdt' },
    { tag: 'w:sdtPr' },
    { tag: 'w:id' },
    { tag: 'w:alias' },
    { tag: 'w:tag' },
    { tag: 'w:text' },
    { tag: 'w:sdtContent' },
    { tag: 'mc:AlternateContent' },
    { tag: 'mc:Choice' },
    { tag: 'mc:Fallback' },
    { tag: 'w:pict' },
    { tag: 'v:rect' },
  ],
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
})
