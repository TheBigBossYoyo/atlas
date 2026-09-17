/**
 * End-to-end test for extractLegacyDocText: builds a minimal but complete
 * in-memory .doc (see `../../__tests__/fixtures.ts`) and verifies the whole
 * doc/fib/pieceTable/textDecode/paragraphs pipeline produces the right
 * paragraphs — rather than re-testing each module's internals again (see
 * their own dedicated test files for that).
 */
import { describe, expect, it } from 'vitest'

import { extractLegacyDocText } from '../index'
import { LegacyFormatError } from '../../errors'
import { buildCfbBytes, buildMinimalDocBytes } from '../../__tests__/fixtures'

describe('extractLegacyDocText', () => {
  it('extracts paragraphs from a minimal well-formed .doc', () => {
    const result = extractLegacyDocText(buildMinimalDocBytes('Hello world\rSecond paragraph\r'))
    expect(result.paragraphs).toEqual(['Hello world', 'Second paragraph'])
  })

  it('throws LegacyFormatError when the WordDocument stream is missing', () => {
    const bytes = buildCfbBytes([['SomethingElse', new Uint8Array([1])]])
    expect(() => extractLegacyDocText(bytes)).toThrow(LegacyFormatError)
  })

  it('throws LegacyFormatError for a file that is not a CFB container at all', () => {
    expect(() => extractLegacyDocText(new Uint8Array([1, 2, 3, 4]))).toThrow(LegacyFormatError)
  })

  it('throws LegacyFormatError when the Table stream the FIB points at is missing', () => {
    // A WordDocument stream whose FIB claims a non-empty CLX ("0Table"), but
    // with no Table stream at all in the container to hold it.
    const buffer = new ArrayBuffer(512)
    const view = new DataView(buffer)
    view.setUint16(0, 0xa5ec, true) // wIdent
    view.setUint32(422, 20, true) // lcbClx — non-empty, but "0Table" won't exist

    const bytes = buildCfbBytes([['WordDocument', new Uint8Array(buffer)]])

    expect(() => extractLegacyDocText(bytes)).toThrow(LegacyFormatError)
  })
})
