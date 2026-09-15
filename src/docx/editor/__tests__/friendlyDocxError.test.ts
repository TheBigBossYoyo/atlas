import { describe, expect, it } from 'vitest'

import { DocxParseError } from '../../parser/unzip'
import { DocxSaveError } from '../../serializer/validateDocxPackage'
import { friendlyDocxErrorMessage } from '../friendlyDocxError'

describe('friendlyDocxErrorMessage', () => {
  it('wraps a JSZip corrupted-archive error on open', () => {
    const error = new DocxParseError("Failed to unzip DOCX: Can't find end of central directory")
    const message = friendlyDocxErrorMessage(error, 'open')

    expect(message).toContain('valid Word document')
    expect(message).toContain("Can't find end of central directory")
    expect(message).not.toMatch(/^Failed to unzip/)
  })

  it('wraps a JSZip corrupted-archive error on save', () => {
    const error = new Error('Corrupted zip: invalid signature')
    const message = friendlyDocxErrorMessage(error, 'save')

    expect(message).toContain('corrupted')
    expect(message).toContain('invalid signature')
  })

  it('frames an existing zip-bomb guard message without re-explaining it', () => {
    const error = new DocxParseError(
      'Entry "word/document.xml" is 999,999,999 bytes uncompressed, over the 200-byte per-entry limit; refusing to extract (possible zip bomb).',
    )
    const message = friendlyDocxErrorMessage(error, 'open')

    expect(message).toContain('not opened')
    expect(message).toContain('possible zip bomb')
  })

  it('wraps an XML corruption error', () => {
    const error = new Error('Unclosed tag "w:p" at position 4021')
    const message = friendlyDocxErrorMessage(error, 'open')

    expect(message).toContain('internal structure')
    expect(message).toContain('Unclosed tag')
  })

  it('wraps a DocxSaveError (post-save validation failure)', () => {
    const error = new DocxSaveError('word/document.xml: <w:tbl> missing required <w:tblGrid>')
    const message = friendlyDocxErrorMessage(error, 'save')

    expect(message).toContain('could not verify')
    expect(message).toContain('tblGrid')
  })

  it('falls back to a generic friendly message for an unrecognized error on open', () => {
    const error = new Error('boom')
    const message = friendlyDocxErrorMessage(error, 'open')

    expect(message).toContain("Couldn't open")
    expect(message).toContain('boom')
  })

  it('falls back to a generic friendly message for an unrecognized error on save', () => {
    const error = new Error('boom')
    const message = friendlyDocxErrorMessage(error, 'save')

    expect(message).toContain("Couldn't save")
    expect(message).toContain('not been written')
  })

  it('handles a non-Error thrown value', () => {
    const message = friendlyDocxErrorMessage('a plain string throw', 'open')

    expect(message).toContain('a plain string throw')
  })
})
