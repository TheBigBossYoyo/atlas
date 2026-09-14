import { describe, expect, it } from 'vitest'

import { parseXml } from '../../shared/xmlUtils'
import { parseOdpTextBody } from '../text'
import type { OdpStyleIndex } from '../styles'

const EMPTY_INDEX: OdpStyleIndex = {
  byName: new Map(),
  listStyles: new Map(),
  masterPages: new Map(),
  pageLayouts: new Map(),
}

/** Builds a standalone `draw:text-box` element from an inner-XML fragment, for unit-testing `parseOdpTextBody` directly. */
function textBoxFrom(innerXml: string): Element {
  const xml = `<draw:text-box
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:xlink="http://www.w3.org/1999/xlink">${innerXml}</draw:text-box>`
  return parseXml(xml).documentElement
}

describe('parseOdpTextBody', () => {
  // Regression: the paragraph walker used to only unwrap one level of
  // <text:span>, and had no branch at all for <text:a> — so a hyperlink's
  // text (or a span nested inside another span) was silently dropped instead
  // of rendered inline.
  it('recurses into a text:a hyperlink instead of dropping its text', () => {
    const container = textBoxFrom(
      '<text:p>See <text:a xlink:href="https://example.com">our site</text:a> for details.</text:p>',
    )

    const { text } = parseOdpTextBody(container, EMPTY_INDEX)
    expect(text).toBe('See our site for details.')
  })

  it('recurses into a text:span nested inside another text:span', () => {
    const container = textBoxFrom(
      '<text:p><text:span text:style-name="Outer">outer <text:span text:style-name="Inner">inner</text:span> tail</text:span></text:p>',
    )

    const { text } = parseOdpTextBody(container, EMPTY_INDEX)
    expect(text).toBe('outer inner tail')
  })

  it('still resolves a single-level span\'s text (no regression on the common case)', () => {
    const container = textBoxFrom('<text:p><text:span>plain</text:span></text:p>')

    const { paragraphs } = parseOdpTextBody(container, EMPTY_INDEX)
    expect(paragraphs[0]?.runs).toEqual([{ text: 'plain' }])
  })
})
