import { describe, expect, it } from 'vitest'

import type { Document, Drawing, Paragraph, Section } from '../../model'
import { twip } from '../../model'
import {
  computeAnchorRect,
  computePageFloats,
  getLineExclusions,
  type FloatRectPt,
  type PageFloat,
} from '../floats'
import { paginate } from '../paginate'
import type { Page } from '../pageTypes'
import type { LineBox } from '../types'

const EMUS_PER_POINT = 12700

const PAGE_RECT: FloatRectPt = { leftPt: 0, topPt: 0, widthPt: 600, heightPt: 800 }
const MARGIN_RECT: FloatRectPt = { leftPt: 72, topPt: 72, widthPt: 456, heightPt: 656 }
const COLUMN_RECT: FloatRectPt = { leftPt: 72, topPt: 72, widthPt: 456, heightPt: 656 }
const PARAGRAPH_RECT: FloatRectPt = { leftPt: 90, topPt: 120, widthPt: 300, heightPt: 14 }

function emu(pt: number): number {
  return pt * EMUS_PER_POINT
}

describe('computeAnchorRect', () => {
  it('resolves relativeFrom page/margin/column for positionH with an explicit offset', () => {
    const base = {
      positionV: { relativeFrom: 'page' as const, offsetEmu: 0 },
      widthPt: 50,
      heightPt: 50,
      pageRect: PAGE_RECT,
      marginRect: MARGIN_RECT,
      columnRect: COLUMN_RECT,
      paragraphRect: PARAGRAPH_RECT,
    }

    expect(
      computeAnchorRect({ ...base, positionH: { relativeFrom: 'page', offsetEmu: emu(10) } }).leftPt,
    ).toBe(10)
    expect(
      computeAnchorRect({ ...base, positionH: { relativeFrom: 'margin', offsetEmu: emu(10) } }).leftPt,
    ).toBe(MARGIN_RECT.leftPt + 10)
    expect(
      computeAnchorRect({ ...base, positionH: { relativeFrom: 'column', offsetEmu: emu(10) } }).leftPt,
    ).toBe(COLUMN_RECT.leftPt + 10)
  })

  it('approximates relativeFrom character (H) and paragraph/line (V) with the anchoring line frame', () => {
    const rect = computeAnchorRect({
      positionH: { relativeFrom: 'character', offsetEmu: 0 },
      positionV: { relativeFrom: 'paragraph', offsetEmu: 0 },
      widthPt: 20,
      heightPt: 20,
      pageRect: PAGE_RECT,
      marginRect: MARGIN_RECT,
      columnRect: COLUMN_RECT,
      paragraphRect: PARAGRAPH_RECT,
    })

    expect(rect.leftPt).toBe(PARAGRAPH_RECT.leftPt)
    expect(rect.topPt).toBe(PARAGRAPH_RECT.topPt)
  })

  it('resolves align=left/center/right against the frame width, honoring the drawing size', () => {
    const frame = { leftPt: 100, widthPt: 200 }
    const drawingWidthPt = 40

    const left = computeAnchorRect({
      positionH: { relativeFrom: 'margin', align: 'left' },
      positionV: undefined,
      widthPt: drawingWidthPt,
      heightPt: 10,
      pageRect: PAGE_RECT,
      marginRect: { ...MARGIN_RECT, leftPt: frame.leftPt, widthPt: frame.widthPt },
      columnRect: COLUMN_RECT,
      paragraphRect: PARAGRAPH_RECT,
    })
    const center = computeAnchorRect({
      positionH: { relativeFrom: 'margin', align: 'center' },
      positionV: undefined,
      widthPt: drawingWidthPt,
      heightPt: 10,
      pageRect: PAGE_RECT,
      marginRect: { ...MARGIN_RECT, leftPt: frame.leftPt, widthPt: frame.widthPt },
      columnRect: COLUMN_RECT,
      paragraphRect: PARAGRAPH_RECT,
    })
    const right = computeAnchorRect({
      positionH: { relativeFrom: 'margin', align: 'right' },
      positionV: undefined,
      widthPt: drawingWidthPt,
      heightPt: 10,
      pageRect: PAGE_RECT,
      marginRect: { ...MARGIN_RECT, leftPt: frame.leftPt, widthPt: frame.widthPt },
      columnRect: COLUMN_RECT,
      paragraphRect: PARAGRAPH_RECT,
    })

    expect(left.leftPt).toBe(frame.leftPt)
    expect(center.leftPt).toBe(frame.leftPt + (frame.widthPt - drawingWidthPt) / 2)
    expect(right.leftPt).toBe(frame.leftPt + (frame.widthPt - drawingWidthPt))
  })

  it('resolves align=top/center/bottom against the vertical frame height', () => {
    const frame = { topPt: 50, heightPt: 100 }
    const drawingHeightPt = 20

    const top = computeAnchorRect({
      positionH: undefined,
      positionV: { relativeFrom: 'margin', align: 'top' },
      widthPt: 10,
      heightPt: drawingHeightPt,
      pageRect: PAGE_RECT,
      marginRect: { ...MARGIN_RECT, topPt: frame.topPt, heightPt: frame.heightPt },
      columnRect: COLUMN_RECT,
      paragraphRect: PARAGRAPH_RECT,
    })
    const bottom = computeAnchorRect({
      positionH: undefined,
      positionV: { relativeFrom: 'margin', align: 'bottom' },
      widthPt: 10,
      heightPt: drawingHeightPt,
      pageRect: PAGE_RECT,
      marginRect: { ...MARGIN_RECT, topPt: frame.topPt, heightPt: frame.heightPt },
      columnRect: COLUMN_RECT,
      paragraphRect: PARAGRAPH_RECT,
    })

    expect(top.topPt).toBe(frame.topPt)
    expect(bottom.topPt).toBe(frame.topPt + (frame.heightPt - drawingHeightPt))
  })

  it('defaults to the frame start (offset 0) when neither align nor offset is present', () => {
    const rect = computeAnchorRect({
      positionH: { relativeFrom: 'page' },
      positionV: { relativeFrom: 'page' },
      widthPt: 10,
      heightPt: 10,
      pageRect: PAGE_RECT,
      marginRect: MARGIN_RECT,
      columnRect: COLUMN_RECT,
      paragraphRect: PARAGRAPH_RECT,
    })

    expect(rect.leftPt).toBe(PAGE_RECT.leftPt)
    expect(rect.topPt).toBe(PAGE_RECT.topPt)
  })

  it('supports a negative posOffset (image placed left of/above its reference frame)', () => {
    const rect = computeAnchorRect({
      positionH: { relativeFrom: 'column', offsetEmu: emu(-5) },
      positionV: { relativeFrom: 'paragraph', offsetEmu: emu(-5) },
      widthPt: 10,
      heightPt: 10,
      pageRect: PAGE_RECT,
      marginRect: MARGIN_RECT,
      columnRect: COLUMN_RECT,
      paragraphRect: PARAGRAPH_RECT,
    })

    expect(rect.leftPt).toBe(COLUMN_RECT.leftPt - 5)
    expect(rect.topPt).toBe(PARAGRAPH_RECT.topPt - 5)
  })
})

function makeFloat(overrides: Partial<PageFloat> = {}): PageFloat {
  return {
    drawing: { kind: 'drawing', layout: 'anchor' } as Drawing,
    blockIndex: 0,
    rect: { leftPt: 100, topPt: 100, widthPt: 50, heightPt: 50 },
    behindDoc: false,
    wrap: { mode: 'square', side: 'bothSides' },
    ...overrides,
  }
}

describe('getLineExclusions', () => {
  it('excludes the float bounding box (± distance) for a square-wrapped float overlapping the line', () => {
    const float = makeFloat({
      rect: { leftPt: 100, topPt: 100, widthPt: 50, heightPt: 50 },
      wrap: { mode: 'square', side: 'bothSides', distLEmu: emu(5), distREmu: emu(5) },
    })

    const exclusions = getLineExclusions([float], 110, 12)

    expect(exclusions).toEqual([{ leftPt: 95, rightPt: 155 }])
  })

  it('returns no exclusion when the line does not vertically overlap the float', () => {
    const float = makeFloat({ rect: { leftPt: 100, topPt: 100, widthPt: 50, heightPt: 50 } })

    expect(getLineExclusions([float], 0, 20)).toEqual([]) // line ends at 20, float starts at 100
    expect(getLineExclusions([float], 150, 20)).toEqual([]) // line starts at 150, float ends at 150
  })

  it('imposes no exclusion for wrap mode none', () => {
    const float = makeFloat({ wrap: { mode: 'none' } })

    expect(getLineExclusions([float], 100, 10)).toEqual([])
  })

  it('excludes the full width for topAndBottom wrap', () => {
    const float = makeFloat({ wrap: { mode: 'topAndBottom' } })

    expect(getLineExclusions([float], 110, 10)).toEqual([
      { leftPt: Number.NEGATIVE_INFINITY, rightPt: Number.POSITIVE_INFINITY },
    ])
  })

  it('excludes out to +Infinity for wrapText=left and from -Infinity for wrapText=right', () => {
    const base = { rect: { leftPt: 100, topPt: 100, widthPt: 50, heightPt: 50 } }

    const left = getLineExclusions(
      [makeFloat({ ...base, wrap: { mode: 'square', side: 'left' } })],
      110,
      10,
    )
    const right = getLineExclusions(
      [makeFloat({ ...base, wrap: { mode: 'square', side: 'right' } })],
      110,
      10,
    )

    expect(left).toEqual([{ leftPt: 100, rightPt: Number.POSITIVE_INFINITY }])
    expect(right).toEqual([{ leftPt: Number.NEGATIVE_INFINITY, rightPt: 150 }])
  })

  it('treats tight/through the same as square (no contour data)', () => {
    const base = { rect: { leftPt: 100, topPt: 100, widthPt: 50, heightPt: 50 } }
    const tight = getLineExclusions([makeFloat({ ...base, wrap: { mode: 'tight' } })], 110, 10)
    const through = getLineExclusions([makeFloat({ ...base, wrap: { mode: 'through' } })], 110, 10)

    expect(tight).toEqual([{ leftPt: 100, rightPt: 150 }])
    expect(through).toEqual([{ leftPt: 100, rightPt: 150 }])
  })

  it('combines exclusions from multiple overlapping floats', () => {
    const a = makeFloat({ rect: { leftPt: 0, topPt: 0, widthPt: 50, heightPt: 50 } })
    const b = makeFloat({ rect: { leftPt: 400, topPt: 0, widthPt: 50, heightPt: 50 } })

    expect(getLineExclusions([a, b], 10, 10)).toEqual([
      { leftPt: 0, rightPt: 50 },
      { leftPt: 400, rightPt: 450 },
    ])
  })
})

function makeLineBox(widthPt: number, lineHeight: number): LineBox {
  return {
    items: [],
    width: widthPt,
    ascent: lineHeight * 0.8,
    descent: lineHeight * 0.2,
    lineHeight,
    isJustified: false,
    justificationStretch: 0,
  }
}

function makeDocumentWithAnchor(drawing: Drawing): Document {
  const paragraph: Paragraph = {
    kind: 'paragraph',
    props: {},
    children: [{ kind: 'run', props: {}, children: [drawing] }],
  }
  const section: Section = {
    kind: 'section',
    props: {
      pgSz: { w: twip(12000), h: twip(16000) },
      pgMar: {
        top: twip(1440),
        right: twip(1440),
        bottom: twip(1440),
        left: twip(1440),
        header: twip(720),
        footer: twip(720),
        gutter: twip(0),
      },
    },
    blocks: [paragraph],
  }

  return {
    kind: 'document',
    sections: [section],
    styles: new Map(),
    numbering: new Map(),
    headers: new Map(),
    footers: new Map(),
    comments: new Map(),
    footnotes: new Map(),
    endnotes: new Map(),
  }
}

function makePage(overrides: Partial<Page> = {}): Page {
  return {
    sectionIndex: 0,
    pageIndex: 0,
    sizePt: { width: 600, height: 800 },
    marginsPt: { top: 72, right: 72, bottom: 72, left: 72, header: 36, footer: 36, gutter: 0 },
    columns: [],
    headerLines: [],
    footerLines: [],
    footnoteLines: [],
    hasFootnoteSeparator: false,
    ...overrides,
  }
}

describe('computePageFloats', () => {
  const anchorDrawing: Drawing = {
    kind: 'drawing',
    layout: 'anchor',
    relationshipId: 'rId1',
    extent: { cx: emu(40), cy: emu(30) },
    behindDoc: true,
    positionH: { relativeFrom: 'margin', offsetEmu: emu(10) },
    positionV: { relativeFrom: 'paragraph', offsetEmu: emu(5) },
    wrap: { mode: 'square', side: 'bothSides' },
  }

  it('places a float using the first laid-out line of its anchoring paragraph', () => {
    const document = makeDocumentWithAnchor(anchorDrawing)
    const page = makePage({
      columns: [
        {
          widthPt: 456,
          leftPt: 72,
          tables: [],
          // `topPt: 72` (not 0) matches what `paginate()` actually produces
          // for a page's first line: `PageLineRef.topPt` is page-absolute,
          // already including `marginsPt.top` (see the dedicated
          // "matches real paginate() output" test below) — a fixture with
          // `topPt: 0` here would silently hide a regression that
          // re-introduces double-counting the top margin.
          lines: [
            { paragraphPath: [0], lineIndex: 0, line: makeLineBox(200, 14), topPt: 72, leftPt: 72 },
          ],
        },
      ],
    })

    const floats = computePageFloats(page, document)

    expect(floats).toHaveLength(1)
    expect(floats[0]).toMatchObject({ blockIndex: 0, behindDoc: true })
    // positionH relativeFrom=margin, offset 10pt -> marginRect.leftPt (72) + 10
    expect(floats[0].rect.leftPt).toBe(82)
    // positionV relativeFrom=paragraph, offset 5pt -> paragraph top
    // (lineRef.topPt (72), already page-absolute — no marginsPt.top added
    // again) + 5
    expect(floats[0].rect.topPt).toBe(77)
    expect(floats[0].rect.widthPt).toBe(40)
    expect(floats[0].rect.heightPt).toBe(30)
  })

  it('produces no float when the anchoring paragraph is not on this page', () => {
    const document = makeDocumentWithAnchor(anchorDrawing)
    const page = makePage({ columns: [{ widthPt: 456, leftPt: 72, tables: [], lines: [] }] })

    expect(computePageFloats(page, document)).toEqual([])
  })

  it('ignores an inline drawing entirely (no float)', () => {
    const inlineDrawing: Drawing = { kind: 'drawing', layout: 'inline', relationshipId: 'rId2' }
    const document = makeDocumentWithAnchor(inlineDrawing)
    const page = makePage({
      columns: [
        {
          widthPt: 456,
          leftPt: 72,
          tables: [],
          lines: [
            { paragraphPath: [0], lineIndex: 0, line: makeLineBox(200, 14), topPt: 0, leftPt: 72 },
          ],
        },
      ],
    })

    expect(computePageFloats(page, document)).toEqual([])
  })

  it('falls back to the default drawing size when extent is missing', () => {
    const drawing: Drawing = {
      kind: 'drawing',
      layout: 'anchor',
      relationshipId: 'rId3',
      positionH: { relativeFrom: 'page', offsetEmu: 0 },
      positionV: { relativeFrom: 'page', offsetEmu: 0 },
    }
    const document = makeDocumentWithAnchor(drawing)
    const page = makePage({
      columns: [
        {
          widthPt: 456,
          leftPt: 72,
          tables: [],
          lines: [
            { paragraphPath: [0], lineIndex: 0, line: makeLineBox(200, 14), topPt: 0, leftPt: 72 },
          ],
        },
      ],
    })

    const floats = computePageFloats(page, document)
    expect(floats[0].rect.widthPt).toBe(96)
    expect(floats[0].rect.heightPt).toBe(96)
  })

  // Regression test for a real double-top-margin bug this task found:
  // `PageLineRef.topPt` (from real `paginate()`, not a hand-built `Page`) is
  // page-absolute already, and an earlier version of `computePageFloats`
  // added `page.marginsPt.top` to it a second time — placing every
  // `relativeFrom: 'paragraph'|'line'` float `marginsPt.top` too far down.
  // Exercising this against actual pagination output (rather than a
  // fixture that could encode the same wrong assumption the buggy code
  // did) is the only way this class of bug gets caught by a test at all.
  it('places a float using the real paginate() line geometry, honoring a nonzero top margin', async () => {
    const drawing: Drawing = {
      kind: 'drawing',
      layout: 'anchor',
      relationshipId: 'rId9',
      extent: { cx: emu(40), cy: emu(30) },
      positionH: { relativeFrom: 'column', offsetEmu: 0 },
      positionV: { relativeFrom: 'paragraph', offsetEmu: 0 },
    }
    const paragraph: Paragraph = {
      kind: 'paragraph',
      props: {},
      children: [{ kind: 'run', props: {}, children: [drawing] }],
    }
    const section: Section = {
      kind: 'section',
      props: {
        pgSz: { w: twip(12000), h: twip(16000) },
        pgMar: {
          top: twip(1440), // 72pt
          right: twip(1440),
          bottom: twip(1440),
          left: twip(1440),
          header: twip(720),
          footer: twip(720),
          gutter: twip(0),
        },
      },
      blocks: [paragraph],
    }
    const document: Document = {
      kind: 'document',
      sections: [section],
      styles: new Map(),
      numbering: new Map(),
      headers: new Map(),
      footers: new Map(),
      comments: new Map(),
      footnotes: new Map(),
      endnotes: new Map(),
    }
    const fontResolver = async () => ({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      lineGap: 0,
      xHeight: 500,
      capHeight: 700,
      advanceWidth: () => 500,
      hasGlyph: () => true,
    })

    const pages = await paginate({ document, fontResolver })
    const floats = computePageFloats(pages[0], document)

    // The float's offset is 0 relative to its anchor paragraph's own first
    // line — it must land at exactly that line's rendered top (72pt, the
    // page's top margin, since this is the first line on the page), not
    // 144pt (72 counted twice).
    const firstLineTopPt = pages[0].columns[0].lines[0].topPt
    expect(firstLineTopPt).toBe(72)
    expect(floats[0].rect.topPt).toBe(firstLineTopPt)
  })
})
