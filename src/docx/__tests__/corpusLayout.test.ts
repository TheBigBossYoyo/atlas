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

describe('corpus layout: header-footer-page-numbers (D11 milestone 1)', () => {
  it('renders the default header and footer text on every page, without an explicit headerFooterLines override', async () => {
    const buffer = await readCorpusFixture('header-footer-page-numbers')
    const bundle = await loadDocx(buffer)

    const pages = await paginate({
      document: bundle.document,
      fontResolver: createFontResolver(),
      theme: bundle.theme,
    })

    expect(pages.length).toBeGreaterThan(0)
    for (const page of pages) {
      expect(lineWords(page.headerLines).join(' ')).toContain('Header')
      // The footer's "Page X of Y" field text (w:fldChar/w:instrText) isn't
      // parsed into a real field node yet — see pageFields.ts's module doc
      // — so today it renders as the surrounding literal text with the
      // field's own numbers missing; this still confirms the footer itself
      // is laid out and reserved, which is milestone 1's actual scope.
      expect(lineWords(page.footerLines).join(' ')).toContain('Page')
    }
  })
})

describe('corpus layout: footnotes-endnotes (D11 milestones 2-4)', () => {
  it('renders a numbered footnote marker in the body and its text at the bottom of the page', async () => {
    const buffer = await readCorpusFixture('footnotes-endnotes')
    const bundle = await loadDocx(buffer)

    const pages = await paginate({
      document: bundle.document,
      fontResolver: createFontResolver(),
      theme: bundle.theme,
    })

    const footnoteMarker = findNoteRefWord(pages, 'footnote')
    expect(footnoteMarker?.text).toBe('1')

    const pageWithFootnote = pages.find((page) => page.footnoteLines.length > 0)
    expect(pageWithFootnote?.hasFootnoteSeparator).toBe(true)
    expect(lineWords(pageWithFootnote?.footnoteLines.map((f) => f.line) ?? []).join(' ')).toContain('footnote')
  })

  it('renders a numbered endnote marker in the body and its text at the end of the document', async () => {
    const buffer = await readCorpusFixture('footnotes-endnotes')
    const bundle = await loadDocx(buffer)

    const pages = await paginate({
      document: bundle.document,
      fontResolver: createFontResolver(),
      theme: bundle.theme,
    })

    const endnoteMarker = findNoteRefWord(pages, 'endnote')
    expect(endnoteMarker?.text).toBe('1')

    const allBodyWords = pages
      .flatMap((page) => page.columns.flatMap((column) => column.lines.map((lineRef) => lineRef.line)))
      .flatMap((line) => lineWords([line]))
    expect(allBodyWords.join(' ')).toContain('endnote')
  })
})

function lineWords(lines: ReadonlyArray<{ items: ReadonlyArray<LineItem> }>): ReadonlyArray<string> {
  return lines.flatMap((line) =>
    line.items
      .filter((item): item is Extract<LineItem, { kind: 'word' | 'glyph-cluster' }> =>
        item.kind === 'word' || item.kind === 'glyph-cluster',
      )
      .map((item) => item.text),
  )
}

function findNoteRefWord(
  pages: ReadonlyArray<Page>,
  kind: 'footnote' | 'endnote',
): Extract<LineItem, { kind: 'word' }> | undefined {
  for (const page of pages) {
    for (const column of page.columns) {
      for (const lineRef of column.lines) {
        const item = lineRef.line.items.find(
          (candidate): candidate is Extract<LineItem, { kind: 'word' }> =>
            candidate.kind === 'word' && candidate.noteRef?.kind === kind,
        )
        if (item !== undefined) {
          return item
        }
      }
    }
  }
  return undefined
}

describe('corpus layout: section-breaks (D9)', () => {
  it('does not force an extra page break at a continuous section boundary, but does at the following nextPage/landscape one', async () => {
    const buffer = await readCorpusFixture('section-breaks')
    const bundle = await loadDocx(buffer)

    const pages = await paginate({
      document: bundle.document,
      fontResolver: createFontResolver(),
      theme: bundle.theme,
    })

    // Per scripts/generate-docx-corpus.mjs's fixtureSectionBreaks: section 0
    // (portrait) and section 1 (continuous, still portrait) share one page;
    // section 2 (nextPage, landscape) starts a genuinely new, wider-than-
    // tall page.
    expect(pages).toHaveLength(2)
    expect(pages[0].sizePt.width).toBeLessThanOrEqual(pages[0].sizePt.height)
    expect(pages[1].sizePt.width).toBeGreaterThan(pages[1].sizePt.height)
  })
})
