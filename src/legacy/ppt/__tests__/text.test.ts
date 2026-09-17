import { describe, expect, it } from 'vitest'

import { BinaryReader } from '../../binaryReader'
import { collectSlideTextGroups } from '../text'

function writeRecordHeader(view: DataView, offset: number, recVer: number, type: number, payloadLength: number): number {
  view.setUint16(offset, recVer & 0x000f, true)
  view.setUint16(offset + 2, type, true)
  view.setUint32(offset + 4, payloadLength, true)
  return offset + 8
}

function utf16leBytes(text: string): Uint8Array {
  return new Uint8Array(new Uint16Array([...text].map((c) => c.charCodeAt(0))).buffer)
}

const RT_TEXT_HEADER_ATOM = 3999
const RT_TEXT_CHARS_ATOM = 4000
const RT_TEXT_BYTES_ATOM = 4008

describe('collectSlideTextGroups', () => {
  it('pairs a TextHeaderAtom with the TextCharsAtom that follows it', () => {
    const text = 'Title text'
    const textBytes = utf16leBytes(text)
    const buffer = new ArrayBuffer(8 + 4 + 8 + textBytes.length)
    const view = new DataView(buffer)

    let offset = writeRecordHeader(view, 0, 0x0, RT_TEXT_HEADER_ATOM, 4)
    view.setInt32(offset, 0, true) // textType = 0 (Title)
    offset += 4
    offset = writeRecordHeader(view, offset, 0x0, RT_TEXT_CHARS_ATOM, textBytes.length)
    new Uint8Array(buffer, offset, textBytes.length).set(textBytes)
    offset += textBytes.length

    const groups = collectSlideTextGroups(new BinaryReader(new Uint8Array(buffer)), 0, offset)

    expect(groups).toEqual([{ textType: 0, text: 'Title text' }])
  })

  it('decodes a TextBytesAtom as cp1252', () => {
    const bodyBytes = new Uint8Array([0x42, 0x6f, 0x64, 0x79]) // "Body"
    const buffer = new ArrayBuffer(8 + 4 + 8 + bodyBytes.length)
    const view = new DataView(buffer)

    let offset = writeRecordHeader(view, 0, 0x0, RT_TEXT_HEADER_ATOM, 4)
    view.setInt32(offset, 1, true) // textType = 1 (Body)
    offset += 4
    offset = writeRecordHeader(view, offset, 0x0, RT_TEXT_BYTES_ATOM, bodyBytes.length)
    new Uint8Array(buffer, offset, bodyBytes.length).set(bodyBytes)
    offset += bodyBytes.length

    const groups = collectSlideTextGroups(new BinaryReader(new Uint8Array(buffer)), 0, offset)

    expect(groups).toEqual([{ textType: 1, text: 'Body' }])
  })

  it('collects multiple text groups from nested containers, in document order', () => {
    const titleBytes = utf16leBytes('Title')
    const bodyBytes = utf16leBytes('Body')

    // Outer container wraps: [TextHeaderAtom(0), TextCharsAtom("Title")], then a nested
    // container wrapping [TextHeaderAtom(1), TextCharsAtom("Body")].
    const innerPayloadLength = 8 + 4 + 8 + bodyBytes.length
    const totalLength =
      8 + // outer container header
      (8 + 4) + // title header atom
      (8 + titleBytes.length) + // title chars atom
      8 + // inner container header
      innerPayloadLength

    const buffer = new ArrayBuffer(totalLength)
    const view = new DataView(buffer)

    let offset = writeRecordHeader(view, 0, 0x0f, 9999, totalLength - 8) // outer container

    offset = writeRecordHeader(view, offset, 0x0, RT_TEXT_HEADER_ATOM, 4)
    view.setInt32(offset, 0, true)
    offset += 4

    offset = writeRecordHeader(view, offset, 0x0, RT_TEXT_CHARS_ATOM, titleBytes.length)
    new Uint8Array(buffer, offset, titleBytes.length).set(titleBytes)
    offset += titleBytes.length

    offset = writeRecordHeader(view, offset, 0x0f, 8888, innerPayloadLength) // inner container
    offset = writeRecordHeader(view, offset, 0x0, RT_TEXT_HEADER_ATOM, 4)
    view.setInt32(offset, 1, true)
    offset += 4
    offset = writeRecordHeader(view, offset, 0x0, RT_TEXT_CHARS_ATOM, bodyBytes.length)
    new Uint8Array(buffer, offset, bodyBytes.length).set(bodyBytes)
    offset += bodyBytes.length

    const groups = collectSlideTextGroups(new BinaryReader(new Uint8Array(buffer)), 0, offset)

    expect(groups).toEqual([
      { textType: 0, text: 'Title' },
      { textType: 1, text: 'Body' },
    ])
  })

  it('leaves textType undefined for a text atom with no preceding TextHeaderAtom', () => {
    const textBytes = utf16leBytes('Orphan')
    const buffer = new ArrayBuffer(8 + textBytes.length)
    const view = new DataView(buffer)

    const offset = writeRecordHeader(view, 0, 0x0, RT_TEXT_CHARS_ATOM, textBytes.length)
    new Uint8Array(buffer, offset, textBytes.length).set(textBytes)

    const groups = collectSlideTextGroups(new BinaryReader(new Uint8Array(buffer)), 0, offset + textBytes.length)

    expect(groups).toEqual([{ textType: undefined, text: 'Orphan' }])
  })

  it('returns an empty list for a slide subtree with no text atoms', () => {
    const buffer = new ArrayBuffer(8)
    const view = new DataView(buffer)
    writeRecordHeader(view, 0, 0x0, 12345, 0)

    const groups = collectSlideTextGroups(new BinaryReader(new Uint8Array(buffer)), 0, 8)
    expect(groups).toEqual([])
  })
})
