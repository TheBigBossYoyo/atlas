import { describe, expect, it } from 'vitest'

import { parseFib } from '../fib'
import { LegacyFormatError } from '../../errors'

/**
 * Builds a minimal, valid-enough "WordDocument" stream: just the FIB fields
 * `parseFib` actually reads, at their real fixed byte offsets ([MS-DOC]
 * 2.5.1), zero-filled everywhere else. Real Word-written files carry a lot
 * more (fonts, styles, ...) between/after these, none of which `parseFib`
 * touches.
 */
function buildFibBytes(options: {
  readonly flags1?: number
  readonly ccpText?: number
  readonly fcClx?: number
  readonly lcbClx?: number
  readonly magic?: number
}): Uint8Array {
  const buffer = new ArrayBuffer(512)
  const view = new DataView(buffer)

  view.setUint16(0, options.magic ?? 0xa5ec, true) // wIdent
  view.setUint16(0x0a, options.flags1 ?? 0, true) // flags1 (fWhichTblStm lives here)
  view.setUint32(76, options.ccpText ?? 100, true) // fibRgLw97.ccpText
  view.setUint32(418, options.fcClx ?? 200, true) // fcClx
  view.setUint32(422, options.lcbClx ?? 50, true) // lcbClx

  return new Uint8Array(buffer)
}

describe('parseFib', () => {
  it('reads ccpText/fcClx/lcbClx at their fixed offsets', () => {
    const fib = parseFib(buildFibBytes({ ccpText: 1234, fcClx: 4096, lcbClx: 300 }))

    expect(fib.ccpText).toBe(1234)
    expect(fib.fcClx).toBe(4096)
    expect(fib.lcbClx).toBe(300)
  })

  it('selects 0Table when fWhichTblStm (bit 0x0200) is clear', () => {
    const fib = parseFib(buildFibBytes({ flags1: 0x0000 }))
    expect(fib.tableStreamName).toBe('0Table')
  })

  it('selects 1Table when fWhichTblStm (bit 0x0200) is set', () => {
    const fib = parseFib(buildFibBytes({ flags1: 0x0200 }))
    expect(fib.tableStreamName).toBe('1Table')
  })

  it('ignores unrelated flag bits when deciding the table stream', () => {
    const fib = parseFib(buildFibBytes({ flags1: 0xfdff })) // every bit set except 0x0200
    expect(fib.tableStreamName).toBe('0Table')
  })

  it('throws LegacyFormatError for a bad FIB signature', () => {
    expect(() => parseFib(buildFibBytes({ magic: 0x1234 }))).toThrow(LegacyFormatError)
  })

  it('throws LegacyFormatError for a stream too short to hold a FIB', () => {
    expect(() => parseFib(new Uint8Array(10))).toThrow(LegacyFormatError)
  })
})
