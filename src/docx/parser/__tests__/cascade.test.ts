/**
 * Tests for src/docx/parser/cascade.ts
 */

import { describe, expect, it } from 'vitest'

import { halfPoint, hexColor, twip } from '../../model'
import type { Style } from '../../model'
import { resolveParaProps, resolveRunProps } from '../cascade'
import { DocxParseError } from '../unzip'

function makeStyles(styles: ReadonlyArray<Style>): ReadonlyMap<string, Style> {
  const map = new Map<string, Style>()
  for (const style of styles) {
    map.set(style.id, style)
  }
  return map
}

describe('resolveRunProps', () => {
  it('returns doc defaults when only doc defaults exist', () => {
    expect(
      resolveRunProps(undefined, undefined, new Map(), {
        rPr: { bold: true, sz: halfPoint(24) },
      }),
    ).toEqual({ bold: true, sz: halfPoint(24) })
  })

  it('returns style props when only a style exists', () => {
    const styles = makeStyles([
      {
        id: 'Emphasis',
        type: 'character',
        run: { italic: true, color: hexColor('AA0000') },
      },
    ])

    expect(resolveRunProps(undefined, 'Emphasis', styles, {})).toEqual({
      italic: true,
      color: 'AA0000',
    })
  })

  it('returns direct props when only direct formatting exists', () => {
    expect(resolveRunProps({ bold: true }, undefined, new Map(), {})).toEqual({
      bold: true,
    })
  })

  it('lets styles override doc defaults', () => {
    const styles = makeStyles([
      {
        id: 'Accent',
        type: 'character',
        run: { color: hexColor('222222'), italic: true },
      },
    ])

    expect(
      resolveRunProps(undefined, 'Accent', styles, {
        rPr: { color: hexColor('111111'), bold: true },
      }),
    ).toEqual({
      color: '222222',
      bold: true,
      italic: true,
    })
  })

  it('lets direct props override style and doc defaults', () => {
    const styles = makeStyles([
      {
        id: 'Accent',
        type: 'character',
        run: { color: hexColor('222222'), italic: true },
      },
    ])

    expect(
      resolveRunProps({ color: hexColor('333333') }, 'Accent', styles, {
        rPr: { color: hexColor('111111'), bold: true },
      }),
    ).toEqual({
      color: '333333',
      bold: true,
      italic: true,
    })
  })

  it('resolves three-level basedOn chains', () => {
    const styles = makeStyles([
      {
        id: 'Base',
        type: 'character',
        run: { color: hexColor('111111') },
      },
      {
        id: 'Mid',
        type: 'character',
        basedOn: 'Base',
        run: { bold: true },
      },
      {
        id: 'Leaf',
        type: 'character',
        basedOn: 'Mid',
        run: { underline: { style: 'single' } },
      },
    ])

    expect(resolveRunProps(undefined, 'Leaf', styles, {})).toEqual({
      color: '111111',
      bold: true,
      underline: { style: 'single' },
    })
  })

  it('throws when a basedOn cycle exceeds the recursion cap', () => {
    const styles = makeStyles([
      {
        id: 'CycleA',
        type: 'character',
        basedOn: 'CycleB',
      },
      {
        id: 'CycleB',
        type: 'character',
        basedOn: 'CycleA',
      },
    ])

    expect(() => resolveRunProps(undefined, 'CycleA', styles, {})).toThrow(
      DocxParseError,
    )
  })

  it('merges linked character styles when resolving run props from a paragraph style', () => {
    const styles = makeStyles([
      {
        id: 'Heading1',
        type: 'paragraph',
        linked: 'Heading1Char',
        run: { bold: true, underline: { style: 'single' } },
      },
      {
        id: 'Heading1Char',
        type: 'character',
        linked: 'Heading1',
        run: {
          italic: true,
          underline: { style: 'single', color: hexColor('00FF00') },
        },
      },
    ])

    expect(
      resolveRunProps({ underline: { style: 'double' } }, 'Heading1', styles, {
        rPr: { color: 'auto' },
      }),
    ).toEqual({
      color: 'auto',
      bold: true,
      italic: true,
      underline: {
        style: 'double',
        color: '00FF00',
      },
    })
  })

  it('returns base plus direct props when style id is missing', () => {
    expect(
      resolveRunProps({ italic: true }, 'Missing', new Map(), {
        rPr: { bold: true },
      }),
    ).toEqual({
      bold: true,
      italic: true,
    })
  })

  it('overrides per property instead of replacing the whole object', () => {
    const styles = makeStyles([
      {
        id: 'Accent',
        type: 'character',
        run: {
          bold: true,
          underline: { style: 'single', color: hexColor('AA0000') },
        },
      },
    ])

    expect(
      resolveRunProps({ italic: true, underline: { style: 'single' } }, 'Accent', styles, {}),
    ).toEqual({
      bold: true,
      italic: true,
      underline: {
        style: 'single',
        color: 'AA0000',
      },
    })
  })
})

describe('resolveParaProps', () => {
  it('returns doc defaults when only paragraph doc defaults exist', () => {
    expect(
      resolveParaProps(undefined, undefined, new Map(), {
        pPr: { jc: 'center', spacing: { before: twip(120) } },
      }),
    ).toEqual({
      jc: 'center',
      spacing: { before: twip(120) },
    })
  })

  it('returns style props when only a paragraph style exists', () => {
    const styles = makeStyles([
      {
        id: 'BodyText',
        type: 'paragraph',
        paragraph: { ind: { left: twip(720) }, keepLines: true },
      },
    ])

    expect(resolveParaProps(undefined, 'BodyText', styles, {})).toEqual({
      ind: { left: twip(720) },
      keepLines: true,
    })
  })

  it('merges doc defaults, style props, and direct props per nested property', () => {
    const styles = makeStyles([
      {
        id: 'BodyText',
        type: 'paragraph',
        paragraph: {
          spacing: { after: twip(180) },
          ind: { left: twip(720) },
          keepNext: true,
        },
      },
    ])

    expect(
      resolveParaProps(
        {
          spacing: { before: twip(60) },
          ind: { hanging: twip(360) },
        },
        'BodyText',
        styles,
        {
          pPr: {
            spacing: { before: twip(120), line: twip(240) },
            jc: 'start',
          },
        },
      ),
    ).toEqual({
      jc: 'start',
      keepNext: true,
      spacing: {
        before: twip(60),
        after: twip(180),
        line: twip(240),
      },
      ind: {
        left: twip(720),
        hanging: twip(360),
      },
    })
  })

  it('returns base plus direct paragraph props when style id is missing', () => {
    expect(
      resolveParaProps(
        { keepLines: true },
        'Missing',
        new Map(),
        { pPr: { pageBreakBefore: true } },
      ),
    ).toEqual({
      pageBreakBefore: true,
      keepLines: true,
    })
  })
})
