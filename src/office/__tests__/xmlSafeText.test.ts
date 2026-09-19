import { describe, expect, it } from 'vitest'

import { parseXmlPart, serializeXmlPart, xmlSafeText } from '../ooxmlDom'

const c = (code: number): string => String.fromCharCode(code)

describe('xmlSafeText', () => {
  it('drops the control characters XML 1.0 forbids and keeps tab, newline and carriage return', () => {
    expect(xmlSafeText(`a${c(0)}b${c(8)}c${c(11)}d${c(12)}e${c(27)}f${c(0xfffe)}g`)).toBe('abcdefg')
    expect(xmlSafeText(`x${c(9)}y${c(10)}z${c(13)}`)).toBe(`x${c(9)}y${c(10)}z${c(13)}`)
    expect(xmlSafeText('café 日本 <&>')).toBe('café 日本 <&>')
  })

  it('keeps a part well-formed when pasted text carried a control character', () => {
    const doc = parseXmlPart('<root><t/></root>')
    doc.getElementsByTagName('t')[0].textContent = xmlSafeText(`pasted${c(1)}text`)
    expect(() => parseXmlPart(serializeXmlPart(doc))).not.toThrow()
  })
})
