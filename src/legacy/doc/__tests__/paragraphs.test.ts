import { describe, expect, it } from 'vitest'

import { splitIntoParagraphs } from '../paragraphs'

describe('splitIntoParagraphs', () => {
  it('splits on the paragraph mark (0x0D)', () => {
    expect(splitIntoParagraphs('First\rSecond\rThird')).toEqual(['First', 'Second', 'Third'])
  })

  it('treats a table cell/row mark (0x07) as a paragraph break too', () => {
    expect(splitIntoParagraphs('Cell ACell B\rNext row')).toEqual(['Cell A', 'Cell B', 'Next row'])
  })

  it('does not emit a trailing empty paragraph for a story ending in a paragraph mark', () => {
    expect(splitIntoParagraphs('Only paragraph\r')).toEqual(['Only paragraph'])
  })

  it('keeps a final paragraph with no trailing mark', () => {
    expect(splitIntoParagraphs('First\rUnterminated')).toEqual(['First', 'Unterminated'])
  })

  it('converts a line break (0x0B) into an embedded soft newline within the same paragraph', () => {
    expect(splitIntoParagraphs('Line oneLine two\r')).toEqual(['Line one\nLine two'])
  })

  it('converts a page break (0x0C) into an embedded soft newline within the same paragraph', () => {
    expect(splitIntoParagraphs('BeforeAfter\r')).toEqual(['Before\nAfter'])
  })

  it('strips other C0 control characters (e.g. an annotation reference mark) but keeps tabs', () => {
    expect(splitIntoParagraphs('AB\tC\r')).toEqual(['AB\tC'])
  })

  it('collapses a field to just its cached result text, dropping the instruction', () => {
    const withField = 'See page  PAGE ' + '3' + '' + ' for details.\r'
    expect(splitIntoParagraphs(withField)).toEqual(['See page 3 for details.'])
  })

  it('drops a field with no cached result (no 0x14 separator) entirely, keeping surrounding text', () => {
    const withField = 'Before  REF bm1 after.\r'
    expect(splitIntoParagraphs(withField)).toEqual(['Before after.'])
  })

  it('drops a nested field along with its parent instructions', () => {
    // Outer field's instruction contains a whole nested field; only the
    // outer field's own cached result should survive.
    const nested = ' outer  inner innerResult outerResult'
    expect(splitIntoParagraphs(`X${nested}Y\r`)).toEqual(['XouterResultY'])
  })

  it('returns an empty array for an empty story', () => {
    expect(splitIntoParagraphs('')).toEqual([])
  })

  it('handles plain text with no special characters as a single paragraph', () => {
    expect(splitIntoParagraphs('Just plain text')).toEqual(['Just plain text'])
  })

  it('produces one paragraph per mark for consecutive paragraph marks (blank paragraphs)', () => {
    expect(splitIntoParagraphs('A\r\rB\r')).toEqual(['A', '', 'B'])
  })
})
