import { describe, expect, it } from 'vitest'

import type { FontMetrics } from '../../fonts'
import { halfPoint, twip } from '../../model'
import type { Paragraph, ParaProps, Run, RunProps } from '../../model'
import { breakLines } from '../breakLines'
import type { FontResolver, LineBreakInput, LineItem, TabStop } from '../types'

const FAKE_METRICS: FontMetrics = {
  unitsPerEm: 10,
  ascender: 8,
  descender: -2,
  lineGap: 1,
  xHeight: 5,
  capHeight: 7,
  advanceWidth: () => 5,
  hasGlyph: () => true,
}

const fontResolver: FontResolver = async () => FAKE_METRICS

const DEFAULT_RUN_PROPS: RunProps = {
  sz: halfPoint(20),
  rFonts: {
    ascii: 'Fake Latin',
    eastAsia: 'Fake CJK',
  },
}

describe('breakLines', () => {
  it('keeps a short line in one box', async () => {
    const lines = await breakLines(createInput([wrapTextRun('hello')], {}, 40))

    expect(lines).toHaveLength(1)
    expect(lines[0]?.width).toBe(25)
    expect(lines[0]?.items.map((item) => item.kind)).toEqual(['word'])
  })

  it('wraps at the previous word boundary when the next word would overflow', async () => {
    const lines = await breakLines(createInput([wrapTextRun('hello world')], {}, 30))

    expect(lines).toHaveLength(2)
    expect(lines[0]?.width).toBe(25)
    expect(lines[0]?.items.map((item) => item.kind)).toEqual(['word', 'space'])
    expect(lines[1]?.width).toBe(25)
    expect(lines[1]?.items.map((item) => item.kind)).toEqual(['word'])
  })

  it('allows an overlong first word to occupy its own line', async () => {
    const lines = await breakLines(createInput([wrapTextRun('hello')], {}, 10))

    expect(lines).toHaveLength(1)
    expect(lines[0]?.width).toBe(25)
  })

  it('drops trailing spaces from the reported line width', async () => {
    const lines = await breakLines(createInput([wrapTextRun('hello   ')], {}, 100))

    expect(lines).toHaveLength(1)
    expect(lines[0]?.width).toBe(25)
    expect(lines[0]?.items).toHaveLength(4)
  })

  it('uses soft hyphen opportunities when a fragment fits but the next one does not', async () => {
    const lines = await breakLines(createInput([wrapTextRun('ab\u00adcdef')], {}, 15))

    expect(lines).toHaveLength(2)
    expect(lines[0]?.width).toBe(15)
    expect(lines[0]?.items.map((item) => item.kind)).toEqual(['word', 'hyphen-opportunity'])
    expect(lines[1]?.width).toBe(20)
    expect(lines[1]?.items.map((item) => item.kind)).toEqual(['word'])
  })

  it('computes justification stretch for non-final justified lines', async () => {
    const lines = await breakLines(createInput([wrapTextRun('a a a a')], { jc: 'both' }, 20))

    expect(lines).toHaveLength(2)
    expect(lines[0]?.width).toBe(15)
    expect(lines[0]?.isJustified).toBe(true)
    expect(lines[0]?.justificationStretch).toBe(5)
  })

  it('does not justify the last line of a both-aligned paragraph', async () => {
    const lines = await breakLines(createInput([wrapTextRun('a a a a')], { jc: 'both' }, 20))

    expect(lines[1]?.isJustified).toBe(false)
    expect(lines[1]?.justificationStretch).toBe(0)
  })

  it('keeps the last distribute-aligned line justified', async () => {
    const lines = await breakLines(createInput([wrapTextRun('a a')], { jc: 'distribute' }, 20))

    expect(lines).toHaveLength(1)
    expect(lines[0]?.isJustified).toBe(true)
    expect(lines[0]?.justificationStretch).toBe(5)
  })

  it('does not justify lines without stretchable spaces', async () => {
    const lines = await breakLines(createInput([wrapTextRun('abcd')], { jc: 'both' }, 40))

    expect(lines).toHaveLength(1)
    expect(lines[0]?.isJustified).toBe(false)
    expect(lines[0]?.justificationStretch).toBe(0)
  })

  it('advances tabs to the next explicit tab stop', async () => {
    const lines = await breakLines(
      createInput(
        [
          wrapRun([
            { kind: 'text', value: 'aa' },
            { kind: 'tab' },
            { kind: 'text', value: 'b' },
          ]),
        ],
        {},
        60,
        [{ positionPt: 30, alignment: 'left', leader: 'none' }],
      ),
    )

    expect(lines).toHaveLength(1)
    expect(findTab(lines[0]?.items)?.width).toBe(20)
    expect(lines[0]?.width).toBe(35)
  })

  it('falls back to 36pt tab increments when no explicit stop applies', async () => {
    const lines = await breakLines(
      createInput(
        [
          wrapRun([
            { kind: 'text', value: 'aa' },
            { kind: 'tab' },
            { kind: 'text', value: 'b' },
          ]),
        ],
        {},
        80,
      ),
    )

    expect(lines).toHaveLength(1)
    expect(findTab(lines[0]?.items)?.width).toBe(26)
    expect(lines[0]?.width).toBe(41)
  })

  it('reduces first-line width by the first-line indent', async () => {
    const lines = await breakLines(
      createInput([wrapTextRun('aa aa aa')], { ind: { firstLine: twip(300) } }, 50),
    )

    expect(lines).toHaveLength(2)
    expect(lines.map((line) => line.width)).toEqual([25, 10])
  })

  it('pulls the first line left under a hanging indent, leaving continuation lines at the base indent (D2/DXL-02)', async () => {
    // `hanging` gives the FIRST line (where a list marker sits, see D3) more
    // room by pulling it back past the base left indent; continuation lines
    // sit at the (unindented-by-hanging) base — the opposite of `firstLine`.
    const lines = await breakLines(
      createInput([wrapTextRun('aa aa aa aa aa aa aa aa')], { ind: { hanging: twip(200) } }, 50),
    )

    expect(lines).toHaveLength(3)
    // Line 0's limit is boosted by the 10pt hanging pull-back (60pt vs the
    // continuation lines' base 50pt), so it fits one more "aa" than they do.
    expect(lines.map((line) => line.width)).toEqual([55, 40, 10])
  })

  it('combines left and first-line indent when computing the first line limit', async () => {
    const lines = await breakLines(
      createInput(
        [wrapTextRun('aa aa aa aa')],
        { ind: { left: twip(200), firstLine: twip(200) } },
        50,
      ),
    )

    expect(lines).toHaveLength(2)
    expect(lines.map((line) => line.width)).toEqual([25, 25])
  })

  it('honors exact line spacing in points', async () => {
    const lines = await breakLines(
      createInput([wrapTextRun('a')], { spacing: { lineRule: 'exact', line: twip(240) } }, 20),
    )

    expect(lines[0]?.lineHeight).toBe(12)
  })

  it('scales natural line height for auto line spacing', async () => {
    const lines = await breakLines(
      createInput([wrapTextRun('a')], { spacing: { lineRule: 'auto', line: twip(480) } }, 20),
    )

    expect(lines[0]?.lineHeight).toBe(22)
  })

  it('uses the natural line height when auto spacing is unspecified', async () => {
    const lines = await breakLines(createInput([wrapTextRun('a')], {}, 20))

    expect(lines[0]?.lineHeight).toBe(11)
  })

  it('enforces a minimum line height for atLeast spacing', async () => {
    const lines = await breakLines(
      createInput([wrapTextRun('a')], { spacing: { lineRule: 'atLeast', line: twip(200) } }, 20),
    )

    expect(lines[0]?.lineHeight).toBe(11)
  })

  it('uses the requested minimum when atLeast exceeds natural height', async () => {
    const lines = await breakLines(
      createInput([wrapTextRun('a')], { spacing: { lineRule: 'atLeast', line: twip(300) } }, 20),
    )

    expect(lines[0]?.lineHeight).toBe(15)
  })

  it('splits line boxes around page breaks inside a paragraph', async () => {
    const lines = await breakLines(
      createInput(
        [
          wrapRun([
            { kind: 'text', value: 'aa' },
            { kind: 'break', breakType: 'page' },
            { kind: 'text', value: 'bb' },
          ]),
        ],
        {},
        40,
      ),
    )

    expect(lines).toHaveLength(2)
    expect(lines[0]?.width).toBe(10)
    expect(lines[0]?.items.map((item) => item.kind)).toEqual(['word', 'break'])
    expect(lines[1]?.width).toBe(10)
    expect(lines[1]?.items.map((item) => item.kind)).toEqual(['word'])
  })

  it('splits line boxes around explicit line breaks', async () => {
    const lines = await breakLines(
      createInput(
        [
          wrapRun([
            { kind: 'text', value: 'aa' },
            { kind: 'break', breakType: 'line' },
            { kind: 'text', value: 'bb' },
          ]),
        ],
        {},
        40,
      ),
    )

    expect(lines).toHaveLength(2)
    expect(lines[0]?.items.map((item) => item.kind)).toEqual(['word', 'break'])
    expect(lines[1]?.items.map((item) => item.kind)).toEqual(['word'])
  })

  it('emits a trailing empty line when a break terminates the paragraph', async () => {
    const lines = await breakLines(
      createInput(
        [
          wrapRun([
            { kind: 'text', value: 'aa' },
            { kind: 'break', breakType: 'page' },
          ]),
        ],
        {},
        40,
      ),
    )

    expect(lines).toHaveLength(2)
    expect(lines[0]?.width).toBe(10)
    expect(lines[1]?.width).toBe(0)
  })

  it('wraps CJK grapheme clusters one cluster at a time', async () => {
    const lines = await breakLines(createInput([wrapTextRun('日本語')], {}, 12))

    expect(lines).toHaveLength(2)
    expect(lines[0]?.items.map(getItemText)).toEqual(['日', '本'])
    expect(lines[1]?.items.map(getItemText)).toEqual(['語'])
  })

  // FAKE_METRICS at 10pt: ascent 8, descent 2, lineGap 1 -> natural height 11,
  // text baseline sits 0.5 (half-leading) + 8 = 8.5pt below the line top.
  it('grows the line box to fit an inline drawing taller than the text ascent', async () => {
    const lines = await breakLines(createInput([wrapRun([drawingNode(20, 20)])], {}, 100))

    expect(lines).toHaveLength(1)
    expect(lines[0]?.items.map((item) => item.kind)).toEqual(['drawing'])
    expect(lines[0]?.width).toBe(20)
    expect(lines[0]?.drawingClearancePt).toBe(11.5)
    expect(lines[0]?.lineHeight).toBe(22.5)
  })

  it('does not add clearance for a drawing shorter than the text ascent', async () => {
    const lines = await breakLines(
      createInput([wrapRun([{ kind: 'text', value: 'ab' }, drawingNode(5, 5)])], {}, 100),
    )

    expect(lines[0]?.lineHeight).toBe(11)
    expect(lines[0]?.drawingClearancePt).toBeUndefined()
  })

  it('applies drawing clearance on top of scaled auto line spacing', async () => {
    const lines = await breakLines(
      createInput([wrapRun([drawingNode(20, 20)])], { spacing: { lineRule: 'auto', line: twip(480) } }, 100),
    )

    // Double spacing: text height 22, baseline at 6 + 8 = 14 -> 6pt raise.
    expect(lines[0]?.drawingClearancePt).toBe(6)
    expect(lines[0]?.lineHeight).toBe(28)
  })

  it('keeps exact line spacing fixed even when a drawing is taller', async () => {
    const lines = await breakLines(
      createInput([wrapRun([drawingNode(20, 20)])], { spacing: { lineRule: 'exact', line: twip(240) } }, 100),
    )

    expect(lines[0]?.lineHeight).toBe(12)
    expect(lines[0]?.drawingClearancePt).toBeUndefined()
  })

  it('only raises the line that actually contains the drawing', async () => {
    const lines = await breakLines(
      createInput([wrapRun([{ kind: 'text', value: 'hello ' }, drawingNode(20, 20)])], {}, 30),
    )

    expect(lines).toHaveLength(2)
    expect(lines[0]?.items.map((item) => item.kind)).toEqual(['word', 'space'])
    expect(lines[0]?.lineHeight).toBe(11)
    expect(lines[1]?.items.map((item) => item.kind)).toEqual(['drawing'])
    expect(lines[1]?.lineHeight).toBe(22.5)
  })

  it('creates an empty line box for an empty paragraph', async () => {
    const lines = await breakLines(createInput([wrapRun([])], {}, 40))

    expect(lines).toHaveLength(1)
    expect(lines[0]?.width).toBe(0)
    expect(lines[0]?.items).toEqual([])
  })
})

function createInput(
  runs: ReadonlyArray<WrappedRun>,
  paraProps: ParaProps = {},
  availableWidth: number = 100,
  tabStops: ReadonlyArray<TabStop> = [],
): LineBreakInput {
  const paragraph: Paragraph = {
    kind: 'paragraph',
    props: paraProps,
    children: runs.map((wrappedRun) => wrappedRun.run),
  }

  return {
    paragraph,
    paraProps,
    runs,
    availableWidth,
    fontResolver,
    tabStops,
  }
}

function wrapTextRun(text: string, runProps: RunProps = DEFAULT_RUN_PROPS): WrappedRun {
  return wrapRun([{ kind: 'text', value: text }], runProps)
}

function wrapRun(children: Run['children'], runProps: RunProps = DEFAULT_RUN_PROPS): WrappedRun {
  return {
    run: {
      kind: 'run',
      children,
    },
    runProps,
  }
}

const EMUS_PER_POINT = 12700

function drawingNode(widthPt: number, heightPt: number): Run['children'][number] {
  return {
    kind: 'drawing',
    layout: 'inline',
    relationshipId: 'rId1',
    extent: { cx: widthPt * EMUS_PER_POINT, cy: heightPt * EMUS_PER_POINT },
  }
}

function findTab(
  items: ReadonlyArray<LineItem> | undefined,
): Extract<LineItem, { kind: 'tab' }> | undefined {
  return items?.find((item): item is Extract<LineItem, { kind: 'tab' }> => item.kind === 'tab')
}

function getItemText(item: LineItem): string {
  if (item.kind === 'word' || item.kind === 'glyph-cluster') {
    return item.text
  }

  return item.kind
}

type WrappedRun = {
  run: Run
  runProps: RunProps
}
