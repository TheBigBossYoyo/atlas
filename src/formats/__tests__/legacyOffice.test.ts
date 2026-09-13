import { describe, expect, it } from 'vitest'

import { guessLegacyOfficeKind, isLegacyOfficeMagic, legacyOfficeMessage } from '../legacyOffice'

function cfbBuffer(extraBytes: ReadonlyArray<number> = []): ArrayBuffer {
  const magic = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]
  return new Uint8Array([...magic, ...extraBytes]).buffer
}

describe('isLegacyOfficeMagic', () => {
  it('detects the OLE2/CFB signature', () => {
    expect(isLegacyOfficeMagic(cfbBuffer())).toBe(true)
  })

  it('detects the signature with trailing content', () => {
    expect(isLegacyOfficeMagic(cfbBuffer([0x01, 0x02, 0x03]))).toBe(true)
  })

  it('returns false for a ZIP-based document (docx)', () => {
    expect(isLegacyOfficeMagic(new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0, 0, 0, 0]).buffer)).toBe(false)
  })

  it('returns false for a buffer shorter than the magic', () => {
    expect(isLegacyOfficeMagic(new Uint8Array([0xd0, 0xcf, 0x11]).buffer)).toBe(false)
  })

  it('returns false for an empty buffer', () => {
    expect(isLegacyOfficeMagic(new ArrayBuffer(0))).toBe(false)
  })
})

describe('guessLegacyOfficeKind', () => {
  it.each([
    ['report.doc', 'doc'],
    ['template.dot', 'doc'],
    ['budget.xls', 'xls'],
    ['template.xlt', 'xls'],
    ['deck.ppt', 'ppt'],
    ['template.pot', 'ppt'],
    ['show.pps', 'ppt'],
  ] as const)('guesses %s -> %s', (path, expected) => {
    expect(guessLegacyOfficeKind(path)).toBe(expected)
  })

  it('returns unknown for an extension outside the legacy Office family', () => {
    expect(guessLegacyOfficeKind('mystery.bin')).toBe('unknown')
  })

  it('returns unknown for a path with no extension', () => {
    expect(guessLegacyOfficeKind('README')).toBe('unknown')
  })

  it('matches case-insensitively', () => {
    expect(guessLegacyOfficeKind('REPORT.DOC')).toBe('doc')
  })
})

describe('legacyOfficeMessage', () => {
  it('names the specific legacy format and its modern replacement', () => {
    expect(legacyOfficeMessage('doc')).toMatch(/\.doc\b/)
    expect(legacyOfficeMessage('doc')).toMatch(/\.docx/)
    expect(legacyOfficeMessage('xls')).toMatch(/\.xls\b/)
    expect(legacyOfficeMessage('xls')).toMatch(/\.xlsx/)
    expect(legacyOfficeMessage('ppt')).toMatch(/\.ppt\b/)
    expect(legacyOfficeMessage('ppt')).toMatch(/\.pptx/)
  })

  it('still gives an actionable message when the kind cannot be guessed', () => {
    expect(legacyOfficeMessage('unknown')).toMatch(/docx.*xlsx.*pptx/)
  })
})
