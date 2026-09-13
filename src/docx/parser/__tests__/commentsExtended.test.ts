/**
 * Tests for src/docx/parser/commentsExtended.ts (D16 / DXS-11)
 */
import { describe, expect, it } from 'vitest'

import { parseCommentsExtended } from '../commentsExtended'
import { DocxParseError } from '../unzip'

const NAMESPACE = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'

function commentsExtendedXml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx ${NAMESPACE}>${body}</w15:commentsEx>`
}

describe('parseCommentsExtended', () => {
  it('returns an empty map for an empty part', () => {
    const result = parseCommentsExtended(commentsExtendedXml(''))
    expect(result.size).toBe(0)
  })

  it('parses a single resolved (done) entry keyed by paraId', () => {
    const result = parseCommentsExtended(
      commentsExtendedXml('<w15:commentEx w15:paraId="12AB34CD" w15:done="1"/>'),
    )

    expect(result.get('12AB34CD')).toBe(true)
  })

  it('parses a single unresolved entry', () => {
    const result = parseCommentsExtended(
      commentsExtendedXml('<w15:commentEx w15:paraId="12AB34CD" w15:done="0"/>'),
    )

    expect(result.get('12AB34CD')).toBe(false)
  })

  it('treats a missing w15:done as not resolved', () => {
    const result = parseCommentsExtended(commentsExtendedXml('<w15:commentEx w15:paraId="12AB34CD"/>'))

    expect(result.get('12AB34CD')).toBe(false)
  })

  it('parses multiple entries', () => {
    const result = parseCommentsExtended(
      commentsExtendedXml(
        '<w15:commentEx w15:paraId="AAAA0001" w15:done="1"/>'
          + '<w15:commentEx w15:paraId="AAAA0002" w15:done="0"/>',
      ),
    )

    expect(result.get('AAAA0001')).toBe(true)
    expect(result.get('AAAA0002')).toBe(false)
    expect(result.size).toBe(2)
  })

  it('skips an entry with no paraId', () => {
    const result = parseCommentsExtended(commentsExtendedXml('<w15:commentEx w15:done="1"/>'))

    expect(result.size).toBe(0)
  })

  it('wraps a fast-xml-parser failure in DocxParseError (D28 / DXP-16 consistency)', () => {
    expect(() => parseCommentsExtended('<<< not xml <<<')).toThrow(DocxParseError)
  })
})
