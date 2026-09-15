/** USR-15 — text boxes keep PowerPoint's insets/anchoring and never clip their own glyphs. */
import { describe, expect, it } from 'vitest'

import { resolveTextBody } from '../shape'
import { textBodyToCss } from '../../../shared/slideStyleHelpers'

function bodyPr(attributes: string): Element {
  const xml = `<a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ${attributes}/>`
  return new DOMParser().parseFromString(xml, 'application/xml').documentElement
}

describe('resolveTextBody', () => {
  it('uses the OOXML default insets and top anchor for a plain body', () => {
    expect(resolveTextBody([bodyPr('')], null)).toEqual({
      insets: { top: 4.8, right: 9.6, bottom: 4.8, left: 9.6 },
      anchor: 'top',
      wrap: true,
    })
  })

  it('lets the slide override the layout, and the layout fill in what the slide leaves unset', () => {
    const slide = bodyPr('lIns="0" wrap="none"')
    const layout = bodyPr('lIns="190500" anchor="b" tIns="95250"')
    const body = resolveTextBody([slide, layout, null], 'body')
    expect(body.insets).toEqual({ top: 10, right: 9.6, bottom: 4.8, left: 0 })
    expect(body.anchor).toBe('bottom')
    expect(body.wrap).toBe(false)
  })

  it('centers titles vertically unless an anchor is set', () => {
    expect(resolveTextBody([null], 'title').anchor).toBe('middle')
    expect(resolveTextBody([bodyPr('anchor="t"')], 'ctrTitle').anchor).toBe('top')
  })
})

describe('textBodyToCss', () => {
  it('maps insets to padding, anchor to flex alignment, and never hides overflowing text', () => {
    const css = textBodyToCss({ insets: { top: 1, right: 2, bottom: 3, left: 4 }, anchor: 'middle', wrap: false })
    expect(css).toMatchObject({
      padding: '1px 2px 3px 4px',
      justifyContent: 'center',
      overflow: 'visible',
      whiteSpace: 'pre',
    })
  })

  it('returns no overrides for a text box without body properties (ODP)', () => {
    expect(textBodyToCss(undefined)).toEqual({})
  })
})
