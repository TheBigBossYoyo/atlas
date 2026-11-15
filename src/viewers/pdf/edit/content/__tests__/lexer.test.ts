import { describe, expect, it } from 'vitest'

import { formatNumber, parseContentStream, tokenize } from '../lexer'
import { parseCMap, splitCodes } from '../cmap'

const ops = (src: string) => parseContentStream(src).map((op) => op.operator)

describe('parseContentStream', () => {
  it('splits operators and keeps each one\'s exact byte span', () => {
    const src = 'BT /F1 12 Tf 72 700 Td (Hi) Tj ET'
    const parsed = parseContentStream(src)
    expect(parsed.map((op) => op.operator)).toEqual(['BT', 'Tf', 'Td', 'Tj', 'ET'])
    expect(parsed.map((op) => src.slice(op.start, op.end))).toEqual(['BT', '/F1 12 Tf', '72 700 Td', '(Hi) Tj', 'ET'])
  })

  it('reads literal strings with nesting, escapes, octal codes and line continuations', () => {
    const [op] = parseContentStream('(a(b)c\\)\\n\\101\\\nd) Tj')
    expect(op.operands[0]).toEqual({ type: 'string', bytes: 'a(b)c)\nAd', hex: false })
  })

  it('reads a raw end of line inside a literal string as one LF, as readers do', () => {
    const cr = String.fromCharCode(13)
    const lf = String.fromCharCode(10)
    const [op] = parseContentStream(`(a${cr}${lf}b${cr}c) Tj`)
    expect(op.operands[0]).toEqual({ type: 'string', bytes: `a${lf}b${lf}c`, hex: false })
  })

  it('reads hex strings, padding an odd final digit and skipping whitespace', () => {
    const [op] = parseContentStream('<48 65 6C6C 6F2> Tj')
    expect(op.operands[0]).toEqual({ type: 'string', bytes: 'Hello ', hex: true })
  })

  it('reads TJ arrays of strings and adjustments', () => {
    const [op] = parseContentStream('[(A) -120 (B) 3.5] TJ')
    expect(op.operands[0]).toEqual({
      type: 'array',
      items: [
        { type: 'string', bytes: 'A', hex: false },
        { type: 'number', value: -120 },
        { type: 'string', bytes: 'B', hex: false },
        { type: 'number', value: 3.5 },
      ],
    })
  })

  it('treats an inline image as one operator, binary data and all', () => {
    // Binary image data that happens to contain operator-like bytes.
    const data = `${String.fromCharCode(0)}Tj ET${String.fromCharCode(0xff)}`
    const src = `q BI /W 2 /H 1 /BPC 8 /CS /G ID ${data} EI Q (after) Tj`
    expect(ops(src)).toEqual(['q', 'BI', 'Q', 'Tj'])
  })

  it('decodes #xx escapes in names and skips comments', () => {
    const [op] = parseContentStream('% a comment\n/A#20B 10 Tf')
    expect(op.operands[0]).toEqual({ type: 'name', value: 'A B' })
  })

  it('accepts the malformed numbers real producers write', () => {
    const [op] = parseContentStream('--3 4- Td')
    expect(op.operands).toEqual([{ type: 'number', value: -3 }, { type: 'number', value: 4 }])
  })
})

describe('formatNumber', () => {
  it('writes plain decimals without exponents or trailing zeros', () => {
    expect(formatNumber(12)).toBe('12')
    expect(formatNumber(0.5)).toBe('0.5')
    expect(formatNumber(-1234.56789)).toBe('-1234.5679')
    expect(formatNumber(1e-9)).toBe('0')
  })
})

describe('CMaps', () => {
  const toUnicode = `
    /CIDInit /ProcSet findresource begin
    1 begincodespacerange <0000> <FFFF> endcodespacerange
    2 beginbfchar <0003> <0020> <0024> <00660069> endbfchar
    1 beginbfrange <0010> <0012> <0041> endbfrange
    1 beginbfrange <0020> <0021> [<00E9> <00C9>] endbfrange
    endcmap`

  it('reads bfchar (including ligatures) and both bfrange forms', () => {
    const cmap = parseCMap(toUnicode)
    expect(cmap.toUnicode.get(0x0003)).toBe(' ')
    expect(cmap.toUnicode.get(0x0024)).toBe('fi')
    expect([0x10, 0x11, 0x12].map((code) => cmap.toUnicode.get(code))).toEqual(['A', 'B', 'C'])
    expect(cmap.toUnicode.get(0x0021)).toBe('É')
    expect(cmap.codeSpace).toEqual([{ length: 2, low: 0, high: 0xffff }])
  })

  it('reads cid ranges and flags a usecmap it cannot resolve', () => {
    const cmap = parseCMap('/UniGB-UCS2-H usecmap 1 begincidrange <20> <7E> 1 endcidrange')
    expect(cmap.usesExternal).toBe(true)
    expect(cmap.cidRanges).toEqual([{ low: 0x20, high: 0x7e, cid: 1 }])
  })

  it('splits mixed-width codes by the code space', () => {
    const codeSpace = [
      { length: 1, low: 0x00, high: 0x80 },
      { length: 2, low: 0x8140, high: 0x9ffc },
    ]
    expect(splitCodes('A\u0081@B', codeSpace)).toEqual([
      { code: 0x41, length: 1 },
      { code: 0x8140, length: 2 },
      { code: 0x42, length: 1 },
    ])
  })

  it('tokenizes a CMap like a content stream', () => {
    expect(tokenize('<0001> <0041>').map((token) => token.type)).toEqual(['string', 'string'])
  })
})
