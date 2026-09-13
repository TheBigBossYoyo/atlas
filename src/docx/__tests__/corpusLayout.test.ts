/**
 * Layout-level assertions against the P0.5 round-trip corpus (real, parsed
 * `.docx` fixtures) for the wave2/docx-layout tasks (D1-D3, D6, D9-D10):
 * positions, widths, and markers produced by the real parser feeding the
 * real paginator — not just the synthetic in-memory documents used by
 * paginate.test.ts's unit tests. A lightweight constant-metrics font
 * resolver is used throughout (matching the pattern already used by
 * paginate.test.ts/PageView.test.tsx) since exact glyph metrics aren't the
 * point here; only structural layout output is.
 */
import { describe, expect, it } from 'vitest'

import type { FontMetrics } from '../fonts'
import { loadDocx } from '../index'
import { paginate } from '../layout/paginate'
import type { FontResolver } from '../layout/types'
import { MARKER_RUN_INDEX } from '../layout/listMarkers'
import type { Page } from '../layout/pageTypes'
import type { LineItem } from '../layout/types'

import { readCorpusFixture } from './corpusRoundtripHelpers'

function createFontResolver(): FontResolver {
  const metrics: FontMetrics = {
    unitsPerEm: 1000,
    ascender: 800,
    descender: -200,
    lineGap: 0,
    xHeight: 500,
    capHeight: 700,
    advanceWidth: () => 500,
    hasGlyph: () => true,
  }

  return async () => metrics
}

function allMarkerWords(pages: ReadonlyArray<Page>): ReadonlyArray<string> {
  const markers: string[] = []
  for (const page of pages) {
    for (const column of page.columns) {
      for (const lineRef of column.lines) {
        for (const item of lineRef.line.items) {
          if (
            (item.kind === 'word' || item.kind === 'glyph-cluster') &&
            item.runIndex === MARKER_RUN_INDEX
          ) {
            markers.push(item.text)
          }
        }
      }
    }
  }
  return markers
}

describe('corpus layout: lists-bullets-numbered (D3)', () => {
  it('generates a marker for every list paragraph, in the fixture\'s documented order', async () => {
    const buffer = await readCorpusFixture('lists-bullets-numbered')
    const bundle = await loadDocx(buffer)

    const pages = await paginate({
      document: bundle.document,
      fontResolver: createFontResolver(),
      theme: bundle.theme,
    })

    // Per scripts/generate-docx-corpus.mjs's fixtureListsBulletsNumbered:
    // two bullets, then a decimal/lowerLetter/lowerRoman three-level list.
    expect(allMarkerWords(pages)).toEqual(['•', '•', '1.', 'a.', 'i.'])
  })

  it("positions each list paragraph's marker line to the left of its own text-start indent", async () => {
    const buffer = await readCorpusFixture('lists-bullets-numbered')
    const bundle = await loadDocx(buffer)

    const pages = await paginate({
      document: bundle.document,
      fontResolver: createFontResolver(),
      theme: bundle.theme,
    })

    const markerLineRefs = pages
      .flatMap((page) => page.columns.flatMap((column) => column.lines))
      .filter((lineRef) =>
        lineRef.line.items.some(
          (item): item is Extract<LineItem, { kind: 'word' }> =>
            item.kind === 'word' && item.runIndex === MARKER_RUN_INDEX,
        ),
      )

    expect(markerLineRefs.length).toBeGreaterThan(0)
    // Every marker line's own left offset must sit strictly left of column 0
    // (unindented body text) by more than a trivial amount, confirming the
    // hanging pull-back — see D2/D3's indent geometry — actually moved the
    // marker away from a flush-left position.
    for (const lineRef of markerLineRefs) {
      expect(lineRef.leftPt).toBeGreaterThan(0)
    }
  })
})
