import { describe, expect, it } from 'vitest'

import { BinaryReader } from '../../binaryReader'
import { LegacyFormatError } from '../../errors'
import { decodeStoryText } from '../textDecode'
import type { Piece } from '../pieceTable'

function wordDocumentWith(
  compressedText: string,
  unicodeText: string,
): { readonly reader: BinaryReader; readonly compressedOffset: number; readonly unicodeOffset: number } {
  const compressedBytes = Uint8Array.from(compressedText, (c) => c.charCodeAt(0))
  const unicodeEncoded = new Uint16Array([...unicodeText].map((c) => c.charCodeAt(0)))
  const unicodeLeBytes = new Uint8Array(unicodeEncoded.buffer)

  const compressedOffset = 0
  const unicodeOffset = compressedBytes.length + 16 // padding gap, so a bug reading the wrong offset would fail loudly

  const total = new Uint8Array(unicodeOffset + unicodeLeBytes.length + 16)
  total.set(compressedBytes, compressedOffset)
  total.set(unicodeLeBytes, unicodeOffset)

  return { reader: new BinaryReader(total), compressedOffset, unicodeOffset }
}

describe('decodeStoryText', () => {
  it('decodes a single compressed (cp1252) piece', () => {
    const { reader, compressedOffset } = wordDocumentWith('Hello', '')
    const pieces: Piece[] = [{ cpStart: 0, cpEnd: 5, isCompressed: true, byteOffset: compressedOffset }]

    expect(decodeStoryText(reader, pieces, 5)).toBe('Hello')
  })

  it('decodes a single Unicode (UTF-16LE) piece', () => {
    const { reader, unicodeOffset } = wordDocumentWith('', 'World')
    const pieces: Piece[] = [{ cpStart: 0, cpEnd: 5, isCompressed: false, byteOffset: unicodeOffset }]

    expect(decodeStoryText(reader, pieces, 5)).toBe('World')
  })

  it('concatenates mixed compressed and Unicode pieces in order', () => {
    const { reader, compressedOffset, unicodeOffset } = wordDocumentWith('Hi ', 'there')
    const pieces: Piece[] = [
      { cpStart: 0, cpEnd: 3, isCompressed: true, byteOffset: compressedOffset },
      { cpStart: 3, cpEnd: 8, isCompressed: false, byteOffset: unicodeOffset },
    ]

    expect(decodeStoryText(reader, pieces, 8)).toBe('Hi there')
  })

  it('decodes high cp1252 bytes (smart quotes) correctly, not as Latin-1', () => {
    // 0x93/0x94 are cp1252 curly double-quotes, U+201C/U+201D.
    const bytes = new Uint8Array([0x93, 0x41, 0x94])
    const pieces: Piece[] = [{ cpStart: 0, cpEnd: 3, isCompressed: true, byteOffset: 0 }]

    expect(decodeStoryText(new BinaryReader(bytes), pieces, 3)).toBe('“A”')
  })

  it('stops decoding once ccpLimit (main story length) is reached, excluding later pieces', () => {
    const { reader, compressedOffset, unicodeOffset } = wordDocumentWith('MainBody', 'Footnote')
    const pieces: Piece[] = [
      { cpStart: 0, cpEnd: 8, isCompressed: true, byteOffset: compressedOffset }, // main story
      { cpStart: 8, cpEnd: 16, isCompressed: false, byteOffset: unicodeOffset }, // footnote story, same CP space
    ]

    expect(decodeStoryText(reader, pieces, 8)).toBe('MainBody')
  })

  it('truncates a piece that straddles the ccpLimit boundary', () => {
    const { reader, compressedOffset } = wordDocumentWith('ABCDE', '')
    const pieces: Piece[] = [{ cpStart: 0, cpEnd: 5, isCompressed: true, byteOffset: compressedOffset }]

    expect(decodeStoryText(reader, pieces, 3)).toBe('ABC')
  })

  it('throws LegacyFormatError when a piece points outside the WordDocument stream', () => {
    const reader = new BinaryReader(new Uint8Array(4))
    const pieces: Piece[] = [{ cpStart: 0, cpEnd: 100, isCompressed: true, byteOffset: 0 }]

    expect(() => decodeStoryText(reader, pieces, 100)).toThrow(LegacyFormatError)
  })
})
