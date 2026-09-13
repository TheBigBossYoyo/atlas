import { describe, expect, it } from 'vitest'

import { decodeTextBuffer, isValidUtf8 } from '../../../electron/lib/textDecoding.cjs'

describe('isValidUtf8', () => {
  it('accepts plain ASCII', () => {
    expect(isValidUtf8(Buffer.from('hello world', 'ascii'))).toBe(true)
  })

  it('accepts valid multi-byte UTF-8 (accented characters and emoji)', () => {
    expect(isValidUtf8(Buffer.from('héllo wörld 🎉', 'utf-8'))).toBe(true)
  })

  it('rejects a lone continuation byte', () => {
    expect(isValidUtf8(Buffer.from([0x80]))).toBe(false)
  })

  it('rejects a truncated multi-byte sequence', () => {
    expect(isValidUtf8(Buffer.from([0xc3]))).toBe(false)
  })

  it('rejects a lone Windows-1252 accented byte (0xE9 with no continuation)', () => {
    // 0xE9 alone looks like the lead byte of a 3-byte UTF-8 sequence but has
    // no continuation bytes — this is the exact "café" .csv scenario from
    // RUN-05 where naive UTF-8 decoding produced U+FFFD mojibake.
    expect(isValidUtf8(Buffer.from([0x63, 0x61, 0x66, 0xe9]))).toBe(false)
  })

  it('rejects an overlong encoding', () => {
    // Overlong 2-byte encoding of NUL (U+0000), which must be rejected.
    expect(isValidUtf8(Buffer.from([0xc0, 0x80]))).toBe(false)
  })

  it('rejects an encoded UTF-16 surrogate', () => {
    // U+D800 encoded as a (technically well-formed) 3-byte UTF-8 sequence.
    expect(isValidUtf8(Buffer.from([0xed, 0xa0, 0x80]))).toBe(false)
  })
})

describe('decodeTextBuffer', () => {
  it('decodes a UTF-8 BOM file, stripping the BOM', () => {
    const buffer = Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from('hello', 'utf-8')])
    expect(decodeTextBuffer(buffer)).toBe('hello')
  })

  it('decodes a UTF-16 LE BOM file', () => {
    const content = Buffer.from('hello', 'utf16le')
    const buffer = Buffer.concat([Buffer.from([0xff, 0xfe]), content])
    expect(decodeTextBuffer(buffer)).toBe('hello')
  })

  it('decodes a UTF-16 BE BOM file', () => {
    const le = Buffer.from('hello', 'utf16le')
    const be = Buffer.alloc(le.length)
    for (let i = 0; i + 1 < le.length; i += 2) {
      be[i] = le[i + 1]
      be[i + 1] = le[i]
    }
    const buffer = Buffer.concat([Buffer.from([0xfe, 0xff]), be])
    expect(decodeTextBuffer(buffer)).toBe('hello')
  })

  it('decodes valid no-BOM UTF-8 byte-for-byte identically to today\'s plain decode', () => {
    const original = '# Heading\n\nSome **markdown** with héllo wörld and 🎉 emoji.\n'
    const buffer = Buffer.from(original, 'utf-8')

    expect(decodeTextBuffer(buffer)).toBe(buffer.toString('utf-8'))
    expect(decodeTextBuffer(buffer)).toBe(original)
  })

  it('falls back to Windows-1252 for invalid-UTF-8 accented content instead of U+FFFD mojibake', () => {
    // "café" with é encoded as the single Windows-1252 byte 0xE9.
    const buffer = Buffer.from([0x63, 0x61, 0x66, 0xe9])
    const decoded = decodeTextBuffer(buffer)
    expect(decoded).toBe('café')
    expect(decoded).not.toContain('�')
  })

  it('maps Windows-1252 curly quotes and em dash correctly', () => {
    // 0x93 “, 0x94 ”, 0x97 — alongside a byte (0xE9) that forces the
    // invalid-UTF-8 fallback path.
    const buffer = Buffer.from([0x93, 0x41, 0x94, 0x20, 0x97, 0xe9])
    expect(decodeTextBuffer(buffer)).toBe('“A” —é')
  })
})
