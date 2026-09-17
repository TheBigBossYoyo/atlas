/**
 * DEFER-2 — right-to-left paragraphs (`w:bidi`): the direction reaches the
 * line boxes, and `w:jc`'s logical start/end flip with it.
 */
import { describe, expect, it } from 'vitest'

import type { FontMetrics } from '../../fonts'
import type { Document, ParaProps, Paragraph, Section } from '../../model'
import { twip } from '../../model'
import { paginate } from '../paginate'
import type { FontResolver } from '../types'

const ARABIC = 'مرحبا بالعالم'

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

function paragraph(text: string, props: ParaProps = {}): Paragraph {
  return {
    kind: 'paragraph',
    props,
    children: [{ kind: 'run', children: [{ kind: 'text', value: text }] }],
  } as Paragraph
}

function documentWith(paragraphs: ReadonlyArray<Paragraph>): Document {
  const section: Section = {
    kind: 'section',
    props: {
      pgSz: { w: twip(400 * 20), h: twip(200 * 20) },
      pgMar: {
        top: twip(0),
        right: twip(0),
        bottom: twip(0),
        left: twip(0),
        header: twip(0),
        footer: twip(0),
        gutter: twip(0),
      },
      cols: { num: 1, space: twip(0), col: [] },
    },
    blocks: paragraphs,
  }
  return {
    kind: 'document',
    sections: [section],
    styles: new Map(),
    numbering: new Map(),
    comments: new Map(),
    footnotes: new Map(),
    endnotes: new Map(),
    headers: new Map(),
    footers: new Map(),
  }
}

async function firstLineLeftPt(paragraphs: ReadonlyArray<Paragraph>): Promise<number> {
  const pages = await paginate({ document: documentWith(paragraphs), fontResolver: createFontResolver() })
  const [line] = pages[0].columns[0].lines
  return line.leftPt
}

describe('right-to-left paragraphs', () => {
  it('marks a bidi paragraph line as rtl, and leaves a plain one alone', async () => {
    const fontResolver = createFontResolver()
    const rtlPages = await paginate({ document: documentWith([paragraph(ARABIC, { bidi: true })]), fontResolver })
    const ltrPages = await paginate({ document: documentWith([paragraph('hello')]), fontResolver })

    expect(rtlPages[0].columns[0].lines[0].line.rtl).toBe(true)
    expect(ltrPages[0].columns[0].lines[0].line.rtl).toBeUndefined()
  })

  it('starts a bidi paragraph at the right edge and a plain one at the left', async () => {
    const ltr = await firstLineLeftPt([paragraph('hello')])
    const rtl = await firstLineLeftPt([paragraph(ARABIC, { bidi: true })])

    expect(ltr).toBeCloseTo(0, 3)
    expect(rtl).toBeGreaterThan(0)
  })

  it('flips w:jc end to the left edge in a bidi paragraph', async () => {
    const endAligned = await firstLineLeftPt([paragraph(ARABIC, { bidi: true, jc: 'end' })])
    expect(endAligned).toBeCloseTo(0, 3)
  })

  it('still right-aligns an end-aligned left-to-right paragraph', async () => {
    const endAligned = await firstLineLeftPt([paragraph('hello', { jc: 'end' })])
    expect(endAligned).toBeGreaterThan(0)
  })
})
