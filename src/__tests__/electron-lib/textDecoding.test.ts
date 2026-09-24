import { describe, expect, it } from 'vitest'

import {
  decodeTextBuffer,
  decodeTextBufferWithMeta,
  encodeTextBuffer,
  isValidUtf8,
  detectNewline,
} from '../../../electron/lib/textDecoding.cjs'

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

// NIGHT/text-roundtrip — SHELL-1 (CRLF -> LF on save)/SHELL-2 (BOM stripped
// on save), plus the wider "UTF-16/Windows-1252 silently re-encoded as
// UTF-8" gap the coordinator flagged in the task brief. Reproduced against
// the UNFIXED code first (see report): `decodeTextBuffer` alone never
// reported an encoding/BOM/newline, so nothing downstream of it COULD have
// reapplied any of this — these tests exercise the fix,
// `decodeTextBufferWithMeta`/`encodeTextBuffer`.
describe('detectNewline', () => {
  it('reports "crlf" for an all-CRLF file', () => {
    expect(detectNewline('a\r\nb\r\nc')).toBe('crlf')
  })
  it('reports "lf" for an all-LF file', () => {
    expect(detectNewline('a\nb\nc')).toBe('lf')
  })
  it('reports "lf" for a file with no line breaks at all', () => {
    expect(detectNewline('just one line')).toBe('lf')
  })
  it('picks the more frequent convention for a mixed file', () => {
    expect(detectNewline('a\r\nb\r\nc\nd\r\ne')).toBe('crlf') // 3 CRLF vs 1 LF
    expect(detectNewline('a\nb\nc\r\nd\ne')).toBe('lf') // 3 LF vs 1 CRLF
  })
  it('breaks an exact tie in favor of "lf" (documented rule)', () => {
    expect(detectNewline('a\r\nb\nc')).toBe('lf') // 1 CRLF, 1 LF
  })
})

describe('decodeTextBufferWithMeta / encodeTextBuffer round trip', () => {
  it('SHELL-1: preserves CRLF line endings through decode -> edit -> encode', () => {
    const original = Buffer.from('# Title\r\n\r\nLine one.\r\nLine two.\r\n', 'utf-8')
    const { content, meta } = decodeTextBufferWithMeta(original)
    expect(meta.newline).toBe('crlf')
    // The editor's own string model is always `\n`-normalized...
    expect(content).toBe('# Title\n\nLine one.\nLine two.\n')
    // ...but an untouched save reproduces the original CRLF bytes exactly.
    const { buffer: reEncoded } = encodeTextBuffer(content, meta)
    expect(reEncoded.equals(original)).toBe(true)
    // A real edit (appending text) still gets the file's CRLF convention,
    // not just whatever the DOM textarea/CodeMirror normalized it to.
    const edited = content + 'Extra line appended.';
    const { buffer: editedBuffer } = encodeTextBuffer(edited, meta)
    expect(editedBuffer.toString('utf-8')).toBe(
      '# Title\r\n\r\nLine one.\r\nLine two.\r\nExtra line appended.',
    )
  })

  it('SHELL-2: preserves a leading UTF-8 BOM through decode -> encode', () => {
    const original = Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from('# Title\n', 'utf-8')])
    const { content, meta } = decodeTextBufferWithMeta(original)
    expect(meta).toEqual({ encoding: 'utf-8', bom: true, newline: 'lf' })
    expect(content).toBe('# Title\n')
    const { buffer } = encodeTextBuffer(content, meta)
    expect(buffer.equals(original)).toBe(true)
    expect(buffer.subarray(0, 3)).toEqual(Buffer.from([0xef, 0xbb, 0xbf]))
  })

  it('round-trips a UTF-16 LE BOM file byte-for-byte unedited', () => {
    const body = Buffer.from('héllo\r\nwörld', 'utf16le')
    const original = Buffer.concat([Buffer.from([0xff, 0xfe]), body])
    const { content, meta } = decodeTextBufferWithMeta(original)
    expect(meta).toEqual({ encoding: 'utf-16le', bom: true, newline: 'crlf' })
    const { buffer } = encodeTextBuffer(content, meta)
    expect(buffer.equals(original)).toBe(true)
  })

  it('round-trips a UTF-16 BE BOM file byte-for-byte unedited', () => {
    const le = Buffer.from('héllo', 'utf16le')
    const be = Buffer.alloc(le.length)
    for (let i = 0; i + 1 < le.length; i += 2) {
      be[i] = le[i + 1]
      be[i + 1] = le[i]
    }
    const original = Buffer.concat([Buffer.from([0xfe, 0xff]), be])
    const { content, meta } = decodeTextBufferWithMeta(original)
    expect(meta.encoding).toBe('utf-16be')
    expect(meta.bom).toBe(true)
    const { buffer } = encodeTextBuffer(content, meta)
    expect(buffer.equals(original)).toBe(true)
  })

  it('re-encodes a Windows-1252 file back to cp1252 bytes (not UTF-8)', () => {
    // "café" with é as the single cp1252 byte 0xE9 — same fixture as the
    // existing decodeTextBuffer Windows-1252 test above.
    const original = Buffer.from([0x63, 0x61, 0x66, 0xe9])
    const { content, meta } = decodeTextBufferWithMeta(original)
    expect(meta.encoding).toBe('windows-1252')
    expect(meta.bom).toBe(false)
    expect(content).toBe('café')
    const { buffer, encodingFallback } = encodeTextBuffer(content, meta)
    expect(encodingFallback).toBe(false)
    expect(buffer.equals(original)).toBe(true)
  })

  it('falls back to honest UTF-8 when an edit adds a character cp1252 cannot represent', () => {
    const original = Buffer.from([0x63, 0x61, 0x66, 0xe9]) // "café"
    const { content, meta } = decodeTextBufferWithMeta(original)
    const edited = content + '🎉' // not representable in Windows-1252
    const { buffer, encodingFallback } = encodeTextBuffer(edited, meta)
    expect(encodingFallback).toBe(true)
    expect(buffer.toString('utf-8')).toBe('café🎉')
  })

  it('a plain no-BOM UTF-8 file with no meta.bom round-trips with no BOM added', () => {
    const original = Buffer.from('plain utf-8, no bom\n', 'utf-8')
    const { content, meta } = decodeTextBufferWithMeta(original)
    expect(meta).toEqual({ encoding: 'utf-8', bom: false, newline: 'lf' })
    const { buffer } = encodeTextBuffer(content, meta)
    expect(buffer.equals(original)).toBe(true)
  })
})
