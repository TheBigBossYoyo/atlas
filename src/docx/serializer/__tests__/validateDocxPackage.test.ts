import { describe, expect, it } from 'vitest'

import { DocxSaveError, validateDocxPackage } from '../validateDocxPackage'

const CONTENT_TYPES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '</Types>'

const DOCUMENT_RELS_XML =
  '<?xml version="1.0" encoding="UTF-8"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
  '</Relationships>'

function documentXml(bodyInner: string): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${bodyInner}</w:body>` +
    '</w:document>'
  )
}

function baseParts(documentBody: string): Map<string, string | Uint8Array> {
  return new Map<string, string | Uint8Array>([
    ['[Content_Types].xml', CONTENT_TYPES_XML],
    ['word/document.xml', documentXml(documentBody)],
    ['word/_rels/document.xml.rels', DOCUMENT_RELS_XML],
    ['word/styles.xml', '<w:styles xmlns:w="x"/>'],
  ])
}

describe('validateDocxPackage', () => {
  it('accepts a well-formed package with a valid w:tbl and resolvable relationships', () => {
    const parts = baseParts(
      '<w:tbl><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>',
    )

    expect(() => validateDocxPackage(parts)).not.toThrow()
  })

  it('accepts a document with no tables at all', () => {
    const parts = baseParts('<w:p><w:r><w:t>Hello</w:t></w:r></w:p>')

    expect(() => validateDocxPackage(parts)).not.toThrow()
  })

  it('throws DocxSaveError when a w:tbl has no w:tblGrid (DXS-02 regression guard)', () => {
    const parts = baseParts('<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>')

    expect(() => validateDocxPackage(parts)).toThrow(DocxSaveError)
    expect(() => validateDocxPackage(parts)).toThrow(/tblGrid/)
  })

  it('catches a missing w:tblGrid on a table nested inside a table cell', () => {
    const outerWithGrid =
      '<w:tbl><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc>' +
      '<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>' +
      '</w:tc></w:tr></w:tbl>'
    const parts = baseParts(outerWithGrid)

    expect(() => validateDocxPackage(parts)).toThrow(DocxSaveError)
  })

  it('throws DocxSaveError with a well-formedness message for malformed XML', () => {
    const parts = baseParts('')
    parts.set('word/document.xml', '<w:document attr="unterminated><w:body/></w:document>')

    expect(() => validateDocxPackage(parts)).toThrow(DocxSaveError)
    expect(() => validateDocxPackage(parts)).toThrow(/not well-formed/)
  })

  it('throws DocxSaveError when a relationship targets a part that was not written', () => {
    const parts = baseParts('<w:p/>')
    parts.set(
      'word/_rels/document.xml.rels',
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>' +
        '</Relationships>',
    )

    expect(() => validateDocxPackage(parts)).toThrow(DocxSaveError)
    expect(() => validateDocxPackage(parts)).toThrow(/comments\.xml/)
  })

  it('throws DocxSaveError when a referenced part has no content-type entry (DXS-06/DXS-07 regression guard)', () => {
    const parts = baseParts('<w:p/>')
    parts.set('word/comments.xml', '<w:comments xmlns:w="x"/>')
    parts.set(
      'word/_rels/document.xml.rels',
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>' +
        '</Relationships>',
    )
    // Note: no Override for /word/comments.xml, and Content_Types has no
    // generic Default for the "xml" extension in this test's fixture either
    // — swap in one without the generic default to force the failure.
    parts.set(
      '[Content_Types].xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '</Types>',
    )

    expect(() => validateDocxPackage(parts)).toThrow(DocxSaveError)
    expect(() => validateDocxPackage(parts)).toThrow(/content-type/)
  })

  it('does not require an existing part for an External relationship target', () => {
    const parts = baseParts('<w:p/>')
    parts.set(
      'word/_rels/document.xml.rels',
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>' +
        '</Relationships>',
    )

    expect(() => validateDocxPackage(parts)).not.toThrow()
  })

  it('treats a URL-shaped target as external even without an explicit TargetMode', () => {
    const parts = baseParts('<w:p/>')
    parts.set(
      'word/_rels/document.xml.rels',
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com"/>' +
        '</Relationships>',
    )

    expect(() => validateDocxPackage(parts)).not.toThrow()
  })

  it('resolves header rels relative to word/, matching real relationship semantics', () => {
    const parts = baseParts('<w:p/>')
    parts.set('word/header1.xml', '<w:hdr xmlns:w="x"/>')
    parts.set(
      'word/_rels/header1.xml.rels',
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/>' +
        '</Relationships>',
    )
    parts.set('word/media/logo.png', new Uint8Array([1, 2, 3]))
    parts.set(
      '[Content_Types].xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="png" ContentType="image/png"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '</Types>',
    )

    expect(() => validateDocxPackage(parts)).not.toThrow()
  })
})
