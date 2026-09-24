import { describe, expect, it } from 'vitest'

import {
  decodeTextBuffer,
  decodeTextBufferWithMeta,
  detectNewline,
  getTextFileMeta,
  setTextFileMeta,
  carryTextFileMeta,
  DEFAULT_TEXT_FILE_META,
  isValidUtf8,
} from '../textDecoding'

function bytes(...values: number[]): Uint8Array {
  return new Uint8Array(values)
}

describe('isValidUtf8', () => {
  it('accepts plain ASCII', () => {
    expect(isValidUtf8(new TextEncoder().encode('hello world'))).toBe(true)
  })

  it('accepts valid multi-byte UTF-8 (accented characters and emoji)', () => {
    expect(isValidUtf8(new TextEncoder().encode('héllo wörld 🎉'))).toBe(true)
  })

  it('rejects a lone continuation byte', () => {
    expect(isValidUtf8(bytes(0x80))).toBe(false)
  })

  it('rejects a truncated multi-byte sequence', () => {
    expect(isValidUtf8(bytes(0xc3))).toBe(false)
  })

  it('rejects a lone Windows-1252 accented byte (0xE9 with no continuation)', () => {
    expect(isValidUtf8(bytes(0x63, 0x61, 0x66, 0xe9))).toBe(false)
  })

  it('rejects an overlong encoding', () => {
    expect(isValidUtf8(bytes(0xc0, 0x80))).toBe(false)
  })

  it('rejects an encoded UTF-16 surrogate', () => {
    expect(isValidUtf8(bytes(0xed, 0xa0, 0x80))).toBe(false)
  })
})

describe('decodeTextBuffer', () => {
  it('decodes a UTF-8 BOM file, stripping the BOM', () => {
    const content = new TextEncoder().encode('hello')
    const buffer = new Uint8Array([0xef, 0xbb, 0xbf, ...content]).buffer
    expect(decodeTextBuffer(buffer)).toBe('hello')
  })

  it('decodes a UTF-16 LE BOM file', () => {
    const le = bytes(0x68, 0x00, 0x69, 0x00) // "hi"
    const buffer = new Uint8Array([0xff, 0xfe, ...le]).buffer
    expect(decodeTextBuffer(buffer)).toBe('hi')
  })

  it('decodes a UTF-16 BE BOM file', () => {
    const be = bytes(0x00, 0x68, 0x00, 0x69) // "hi"
    const buffer = new Uint8Array([0xfe, 0xff, ...be]).buffer
    expect(decodeTextBuffer(buffer)).toBe('hi')
  })

  it("decodes valid no-BOM UTF-8 byte-for-byte identically to today's plain TextDecoder", () => {
    const original = '# Heading\n\nSome **markdown** with héllo wörld and 🎉 emoji.\n'
    const buffer = new TextEncoder().encode(original).buffer

    expect(decodeTextBuffer(buffer)).toBe(new TextDecoder('utf-8').decode(buffer))
    expect(decodeTextBuffer(buffer)).toBe(original)
  })

  it('falls back to Windows-1252 for invalid-UTF-8 accented content instead of U+FFFD mojibake', () => {
    const buffer = bytes(0x63, 0x61, 0x66, 0xe9).buffer as ArrayBuffer // "café" with é as a raw cp1252 byte
    const decoded = decodeTextBuffer(buffer)
    expect(decoded).toBe('café')
    expect(decoded).not.toContain('�')
  })
})

// NIGHT/text-roundtrip — renderer-side mirror of the electron-lib metadata
// tests. Covers the `prefetchedBuffer` decode path in `useFileHandler.ts`
// (the Open dialog reads bytes in the renderer, not main) and the path-keyed
// meta registry `App.tsx`/`CodeViewer.tsx` read from at save time.
describe('detectNewline', () => {
  it('reports "crlf" for an all-CRLF file', () => {
    expect(detectNewline('a\r\nb\r\nc')).toBe('crlf')
  })
  it('reports "lf" for an all-LF file', () => {
    expect(detectNewline('a\nb\nc')).toBe('lf')
  })
})

describe('decodeTextBufferWithMeta', () => {
  it('normalizes CRLF content to LF and reports meta.newline "crlf"', () => {
    const buffer = new TextEncoder().encode('a\r\nb\r\nc').buffer
    const { content, meta } = decodeTextBufferWithMeta(buffer)
    expect(content).toBe('a\nb\nc')
    expect(meta.newline).toBe('crlf')
  })

  it('detects a UTF-8 BOM and reports it in meta', () => {
    const body = new TextEncoder().encode('hello')
    const buffer = new Uint8Array([0xef, 0xbb, 0xbf, ...body]).buffer
    const { content, meta } = decodeTextBufferWithMeta(buffer)
    expect(content).toBe('hello')
    expect(meta).toEqual({ encoding: 'utf-8', bom: true, newline: 'lf' })
  })

  it('detects Windows-1252 fallback content', () => {
    const buffer = bytes(0x63, 0x61, 0x66, 0xe9).buffer as ArrayBuffer // "café"
    const { content, meta } = decodeTextBufferWithMeta(buffer)
    expect(content).toBe('café')
    expect(meta.encoding).toBe('windows-1252')
    expect(meta.bom).toBe(false)
  })
})

describe('text file meta registry', () => {
  it('returns today\'s defaults for an untracked path', () => {
    expect(getTextFileMeta('/never/seen.md')).toEqual(DEFAULT_TEXT_FILE_META)
  })

  it('remembers a path\'s meta after setTextFileMeta', () => {
    setTextFileMeta('/tracked/a.md', { encoding: 'utf-16le', bom: true, newline: 'crlf' })
    expect(getTextFileMeta('/tracked/a.md')).toEqual({ encoding: 'utf-16le', bom: true, newline: 'crlf' })
  })

  it('carries meta forward to a new path on Save As, leaving the old path\'s entry untouched', () => {
    setTextFileMeta('/tracked/b.md', { encoding: 'windows-1252', bom: false, newline: 'lf' })
    carryTextFileMeta('/tracked/b.md', '/tracked/b-renamed.md')
    expect(getTextFileMeta('/tracked/b-renamed.md')).toEqual({ encoding: 'windows-1252', bom: false, newline: 'lf' })
    expect(getTextFileMeta('/tracked/b.md')).toEqual({ encoding: 'windows-1252', bom: false, newline: 'lf' })
  })

  it('is a no-op when the old path was never tracked', () => {
    carryTextFileMeta('/never/tracked.md', '/tracked/c.md')
    expect(getTextFileMeta('/tracked/c.md')).toEqual(DEFAULT_TEXT_FILE_META)
  })
})
