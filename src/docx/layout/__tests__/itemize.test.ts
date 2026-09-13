import { describe, expect, it } from 'vitest'

import type { FontMetrics } from '../../fonts'
import { halfPoint } from '../../model'
import type { Drawing, Run, RunProps } from '../../model'
import { itemizeRuns } from '../itemize'
import type { FontResolver } from '../types'

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

describe('itemizeRuns', () => {
  it('returns no items for an empty run', async () => {
    const items = await itemizeRuns([wrapRun([])], fontResolver)
    expect(items).toEqual([])
  })

  it('emits a single word item for a single word', async () => {
    const items = await itemizeRuns([wrapTextRun('hello')], fontResolver)

    expect(items).toHaveLength(1)
    expect(items[0]).toMatchObject({
      kind: 'word',
      text: 'hello',
      width: 25,
      runIndex: 0,
      charStart: 0,
      charEnd: 5,
    })
  })

  it('splits words on whitespace while preserving spaces', async () => {
    const items = await itemizeRuns([wrapTextRun('hello world')], fontResolver)

    expect(items).toHaveLength(3)
    expect(items.map((item) => item.kind)).toEqual(['word', 'space', 'word'])
    expect(items[1]).toMatchObject({ kind: 'space', width: 5, charOffset: 5 })
  })

  it('keeps no-break spaces inside the surrounding word', async () => {
    const items = await itemizeRuns([wrapTextRun('hello\u00a0world')], fontResolver)

    expect(items).toHaveLength(1)
    expect(items[0]).toMatchObject({
      kind: 'word',
      text: 'hello\u00a0world',
      width: 55,
    })
  })

  it('emits separate items for repeated spaces', async () => {
    const items = await itemizeRuns([wrapTextRun('a  b')], fontResolver)

    expect(items).toHaveLength(4)
    expect(items.map((item) => item.kind)).toEqual(['word', 'space', 'space', 'word'])
  })

  it('emits tab items for tab nodes', async () => {
    const items = await itemizeRuns([wrapRun([{ kind: 'tab' }])], fontResolver)

    expect(items).toEqual([
      {
        kind: 'tab',
        width: 0,
        runIndex: 0,
        charOffset: 0,
      },
    ])
  })

  it('normalizes line breaks into break items', async () => {
    const items = await itemizeRuns(
      [wrapRun([{ kind: 'break', breakType: 'textWrapping' }])],
      fontResolver,
    )

    expect(items).toEqual([
      {
        kind: 'break',
        breakKind: 'line',
        runIndex: 0,
      },
    ])
  })

  it('preserves page breaks distinctly', async () => {
    const items = await itemizeRuns([wrapRun([{ kind: 'break', breakType: 'page' }])], fontResolver)

    expect(items).toEqual([
      {
        kind: 'break',
        breakKind: 'page',
        runIndex: 0,
      },
    ])
  })

  it('emits soft hyphen opportunities between word fragments', async () => {
    const items = await itemizeRuns([wrapTextRun('co\u00adoperate')], fontResolver)

    expect(items).toHaveLength(3)
    expect(items.map((item) => item.kind)).toEqual(['word', 'hyphen-opportunity', 'word'])
    expect(items[0]).toMatchObject({ kind: 'word', text: 'co', width: 10 })
    expect(items[1]).toMatchObject({ kind: 'hyphen-opportunity', penaltyWidth: 5, charOffset: 2 })
    expect(items[2]).toMatchObject({ kind: 'word', text: 'operate', width: 35, charStart: 3 })
  })

  it('does not emit a hyphen opportunity when the soft hyphen is terminal', async () => {
    const items = await itemizeRuns([wrapTextRun('co\u00ad')], fontResolver)

    expect(items).toHaveLength(1)
    expect(items[0]).toMatchObject({ kind: 'word', text: 'co', width: 10 })
  })

  it('emits glyph clusters for CJK text', async () => {
    const items = await itemizeRuns([wrapTextRun('日本')], fontResolver)

    expect(items).toHaveLength(2)
    expect(items.map((item) => item.kind)).toEqual(['glyph-cluster', 'glyph-cluster'])
    expect(items).toMatchObject([
      { kind: 'glyph-cluster', text: '日', width: 5, charStart: 0, charEnd: 1 },
      { kind: 'glyph-cluster', text: '本', width: 5, charStart: 1, charEnd: 2 },
    ])
  })

  it('emits a drawing item carrying the source node and intrinsic size', async () => {
    const drawing: Drawing = {
      kind: 'drawing',
      layout: 'inline',
      relationshipId: 'rId7',
      description: 'Company logo',
      extent: { cx: 254000, cy: 127000 },
    }

    const items = await itemizeRuns([wrapRun([drawing])], fontResolver)

    expect(items).toHaveLength(1)
    expect(items[0]).toMatchObject({
      kind: 'drawing',
      width: 20,
      height: 10,
      runIndex: 0,
      charStart: 0,
      charEnd: 1,
    })
    expect(items[0]?.kind === 'drawing' ? items[0].drawing : undefined).toBe(drawing)
  })

  it('never collapses a drawing into a placeholder text word', async () => {
    const items = await itemizeRuns(
      [wrapRun([{ kind: 'drawing', layout: 'inline', relationshipId: 'rId1' }])],
      fontResolver,
    )

    expect(items.some((item) => item.kind === 'word' && item.text.includes('￼'))).toBe(false)
  })

  it('itemizes anchored drawings as drawing items too', async () => {
    const items = await itemizeRuns(
      [wrapRun([{ kind: 'drawing', layout: 'anchor', relationshipId: 'rId2', extent: { cx: 127000, cy: 254000 } }])],
      fontResolver,
    )

    expect(items[0]).toMatchObject({ kind: 'drawing', width: 10, height: 20 })
    expect(items[0]?.kind === 'drawing' ? items[0].drawing.layout : undefined).toBe('anchor')
  })

  it('falls back to a default drawing size when the intrinsic extent is missing', async () => {
    const items = await itemizeRuns([wrapRun([{ kind: 'drawing', layout: 'inline' }])], fontResolver)

    expect(items).toHaveLength(1)
    expect(items[0]).toMatchObject({ kind: 'drawing', width: 96, height: 96 })
  })

  it('keeps character offsets of surrounding text stable around a drawing', async () => {
    const items = await itemizeRuns(
      [
        wrapRun([
          { kind: 'text', value: 'ab' },
          { kind: 'drawing', layout: 'inline', extent: { cx: 127000, cy: 127000 } },
          { kind: 'text', value: 'cd' },
        ]),
      ],
      fontResolver,
    )

    expect(items).toMatchObject([
      { kind: 'word', text: 'ab', charStart: 0, charEnd: 2 },
      { kind: 'drawing', charStart: 2, charEnd: 3 },
      { kind: 'word', text: 'cd', charStart: 3, charEnd: 5 },
    ])
  })
})

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

type WrappedRun = {
  run: Run
  runProps: RunProps
}
