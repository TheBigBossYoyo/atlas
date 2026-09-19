/**
 * Unit tests for `scripts/lib/officeValidator.mjs` itself (D-VALID —
 * document fidelity audit): proves each check actually fires on a
 * synthetic package built to violate exactly one rule, and that a clean
 * minimal package produces zero issues (the negative control — without it,
 * an overly eager check could pass every "real" test in this repo for the
 * wrong reason: everything it's run against already happens to dodge it).
 *
 * Every package here is built from scratch with `JSZip`, independent of any
 * Atlas reader/writer — these tests are about the VALIDATOR's own
 * correctness, not about anything Atlas saves (see the various
 * `officeFileValidation.test.ts` suites under `src/docx`, `src/viewers/
 * spreadsheet`, and `src/viewers/slides` for that).
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { validateOfficeFile } from '../../../scripts/lib/officeValidator.mjs'

const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

type ContentTypesOptions = { readonly includePngDefault?: boolean }

function contentTypesXml(opts: ContentTypesOptions = {}): string {
  const png = opts.includePngDefault ? '<Default Extension="png" ContentType="image/png"/>' : ''
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${CT_NS}">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>${png}` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
    `</Types>`
  )
}

const PACKAGE_RELS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
  `</Relationships>`

const DOCUMENT_RELS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
  `</Relationships>`

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W_NS}"/>`

type DocumentBodyOptions = {
  readonly pPrFirst?: boolean
  readonly tblOrderOk?: boolean
  readonly includeTblGrid?: boolean
  readonly rId?: string
}

function documentXml(opts: DocumentBodyOptions = {}): string {
  const { pPrFirst = true, tblOrderOk = true, includeTblGrid = true, rId = 'rId1' } = opts
  const pPr = '<w:pPr><w:jc w:val="center"/></w:pPr>'
  const run = '<w:r><w:t>Hello</w:t></w:r>'
  const paragraph = pPrFirst ? `<w:p>${pPr}${run}</w:p>` : `<w:p>${run}${pPr}</w:p>`

  const tblPr = '<w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>'
  const tblGrid = includeTblGrid ? '<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' : ''
  const row = '<w:tr><w:tc><w:p/></w:tc></w:tr>'
  const table = tblOrderOk ? `<w:tbl>${tblPr}${tblGrid}${row}</w:tbl>` : `<w:tbl>${tblGrid}${row}${tblPr}</w:tbl>`

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="${W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>` +
    `${paragraph}${table}` +
    `<w:sectPr r:id="${rId}"/>` +
    `</w:body></w:document>`
  )
}

type PackageOptions = ContentTypesOptions &
  DocumentBodyOptions & {
    readonly documentRelsXml?: string
    readonly extraFiles?: Readonly<Record<string, string>>
  }

async function buildDocxBuffer(opts: PackageOptions = {}): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', contentTypesXml(opts))
  zip.file('_rels/.rels', PACKAGE_RELS_XML)
  zip.file('word/document.xml', documentXml(opts))
  zip.file('word/_rels/document.xml.rels', opts.documentRelsXml ?? DOCUMENT_RELS_XML)
  zip.file('word/styles.xml', STYLES_XML)
  for (const [path, content] of Object.entries(opts.extraFiles ?? {})) zip.file(path, content)
  return zip.generateAsync({ type: 'nodebuffer' })
}

function codesOf(issues: ReadonlyArray<{ readonly code: string }>): string[] {
  return issues.map((i) => i.code)
}

describe('validateOfficeFile — negative control', () => {
  it('a correctly-built minimal docx has zero issues', async () => {
    const buffer = await buildDocxBuffer()
    const result = validateOfficeFile(buffer)
    expect(result.format).toEqual({ family: 'opc', kind: 'docx' })
    expect(result.issues).toEqual([])
  })
})

describe('validateOfficeFile — OPC checks', () => {
  it('flags a part with no content-type entry', async () => {
    const buffer = await buildDocxBuffer({ extraFiles: { 'word/media/image1.png': 'x' } })
    const result = validateOfficeFile(buffer)
    expect(codesOf(result.issues)).toContain('part-missing-content-type')
  })

  it('does not flag a part once its extension has a Default entry', async () => {
    const buffer = await buildDocxBuffer({
      includePngDefault: true,
      extraFiles: { 'word/media/image1.png': 'x' },
    })
    const result = validateOfficeFile(buffer)
    expect(codesOf(result.issues)).not.toContain('part-missing-content-type')
  })

  it('flags a relationship target that does not exist in the package', async () => {
    const badRels =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="nonexistent.xml"/>` +
      `</Relationships>`
    const buffer = await buildDocxBuffer({ documentRelsXml: badRels })
    const result = validateOfficeFile(buffer)
    expect(codesOf(result.issues)).toContain('dangling-relationship-target')
  })

  it('flags an r:id referenced in a part but not declared in its .rels file', async () => {
    const buffer = await buildDocxBuffer({ rId: 'rIdMissing' })
    const result = validateOfficeFile(buffer)
    expect(codesOf(result.issues)).toContain('dangling-rid')
  })

  it('flags a missing required part', async () => {
    const zip = new JSZip()
    zip.file('[Content_Types].xml', contentTypesXml())
    zip.file('_rels/.rels', PACKAGE_RELS_XML)
    // word/document.xml deliberately omitted.
    const buffer = await zip.generateAsync({ type: 'nodebuffer' })
    const result = validateOfficeFile(buffer)
    expect(codesOf(result.issues)).toContain('missing-required-part')
  })
})

describe('validateOfficeFile — XML well-formedness / character / namespace checks', () => {
  it('flags an XML-1.0-illegal control character', async () => {
    const buffer = await buildDocxBuffer({
      extraFiles: { 'word/document.xml': documentXml().replace('Hello', 'Hello') },
    })
    const result = validateOfficeFile(buffer)
    expect(codesOf(result.issues)).toContain('illegal-xml-char')
  })

  it('flags malformed (non-well-formed) XML', async () => {
    const buffer = await buildDocxBuffer({
      extraFiles: { 'word/styles.xml': '<w:styles xmlns:w="urn:x"><w:open>' },
    })
    const result = validateOfficeFile(buffer)
    expect(codesOf(result.issues)).toContain('malformed-xml')
  })

  it('flags an undeclared namespace prefix', async () => {
    const badDoc = documentXml().replace(
      '<w:sectPr',
      '<a:graphic><a:graphicData uri="x"/></a:graphic><w:sectPr',
    )
    const buffer = await buildDocxBuffer({ extraFiles: { 'word/document.xml': badDoc } })
    const result = validateOfficeFile(buffer)
    expect(codesOf(result.issues)).toContain('namespace-not-declared')
  })
})

describe('validateOfficeFile — element order checks', () => {
  it('flags w:pPr appearing after run content instead of first', async () => {
    const buffer = await buildDocxBuffer({ pPrFirst: false })
    const result = validateOfficeFile(buffer)
    expect(codesOf(result.issues)).toContain('element-order')
  })

  it('flags a w:tbl with no w:tblGrid', async () => {
    const buffer = await buildDocxBuffer({ includeTblGrid: false })
    const result = validateOfficeFile(buffer)
    expect(codesOf(result.issues)).toContain('missing-tblgrid')
  })

  it('flags w:tblPr/w:tblGrid/w:tr out of CT_Tbl order', async () => {
    const buffer = await buildDocxBuffer({ tblOrderOk: false })
    const result = validateOfficeFile(buffer)
    expect(codesOf(result.issues)).toContain('element-order')
  })

  it('flags a:lstStyle appearing after a:p in p:txBody', async () => {
    const AN = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    const PN = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    const badTxBody = `<p:txBody xmlns:p="${PN}" xmlns:a="${AN}"><a:bodyPr/><a:p/><a:lstStyle/></p:txBody>`
    const doc =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W_NS}">` +
      `<w:body><w:p/>${badTxBody}</w:body></w:document>`
    const buffer = await buildDocxBuffer({ extraFiles: { 'word/document.xml': doc } })
    const result = validateOfficeFile(buffer)
    expect(codesOf(result.issues)).toContain('element-order')
  })

  it('flags a worksheet with <sheetData> before <dimension>', async () => {
    const SM_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    const zip = new JSZip()
    zip.file(
      '[Content_Types].xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${CT_NS}">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
        `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
        `</Types>`,
    )
    zip.file(
      '_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    )
    zip.file(
      'xl/workbook.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${SM_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    )
    zip.file(
      'xl/worksheets/sheet1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${SM_NS}"><sheetData/><dimension ref="A1"/></worksheet>`,
    )
    const buffer = await zip.generateAsync({ type: 'nodebuffer' })
    const result = validateOfficeFile(buffer)
    expect(codesOf(result.issues)).toContain('element-order')
  })
})

describe('validateOfficeFile — zip container checks', () => {
  it('flags a zip entry name starting with "/"', async () => {
    const zip = new JSZip()
    zip.file('[Content_Types].xml', contentTypesXml())
    zip.file('_rels/.rels', PACKAGE_RELS_XML)
    zip.file('word/document.xml', documentXml())
    zip.file('word/_rels/document.xml.rels', DOCUMENT_RELS_XML)
    zip.file('word/styles.xml', STYLES_XML)
    zip.file('/word/extra.xml', '<x/>')
    const buffer = await zip.generateAsync({ type: 'nodebuffer' })
    const result = validateOfficeFile(buffer)
    expect(codesOf(result.issues)).toContain('zip-leading-slash')
  })
})

describe('validateOfficeFile — ODF checks', () => {
  async function buildOdfBuffer(opts: {
    readonly mimetypeFirst?: boolean
    readonly mimetypeStored?: boolean
    readonly omitFromManifest?: boolean
  } = {}): Promise<Buffer> {
    const { mimetypeFirst = true, mimetypeStored = true, omitFromManifest = false } = opts
    const manifest =
      `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">` +
      `<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.presentation"/>` +
      (omitFromManifest ? '' : `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>`) +
      `</manifest:manifest>`
    const content = `<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.2"/>`

    const zip = new JSZip()
    const mimetypeOptions = { compression: mimetypeStored ? 'STORE' : 'DEFLATE' } as const
    if (mimetypeFirst) {
      zip.file('mimetype', 'application/vnd.oasis.opendocument.presentation', mimetypeOptions)
      zip.file('META-INF/manifest.xml', manifest)
    } else {
      zip.file('META-INF/manifest.xml', manifest)
      zip.file('mimetype', 'application/vnd.oasis.opendocument.presentation', mimetypeOptions)
    }
    zip.file('content.xml', content)
    return zip.generateAsync({ type: 'nodebuffer' })
  }

  it('a correctly-built minimal odp has zero issues', async () => {
    const buffer = await buildOdfBuffer()
    const result = validateOfficeFile(buffer)
    expect(result.format).toEqual({ family: 'odf', kind: 'odp', mimetype: 'application/vnd.oasis.opendocument.presentation' })
    expect(result.issues).toEqual([])
  })

  it('flags mimetype not being the first zip entry', async () => {
    const buffer = await buildOdfBuffer({ mimetypeFirst: false })
    const result = validateOfficeFile(buffer)
    expect(codesOf(result.issues)).toContain('odf-mimetype-not-first')
  })

  it('flags mimetype being compressed instead of stored', async () => {
    const buffer = await buildOdfBuffer({ mimetypeStored: false })
    const result = validateOfficeFile(buffer)
    expect(codesOf(result.issues)).toContain('odf-mimetype-not-stored')
  })

  it('flags a file present in the package but missing from the manifest', async () => {
    const buffer = await buildOdfBuffer({ omitFromManifest: true })
    const result = validateOfficeFile(buffer)
    expect(codesOf(result.issues)).toContain('odf-file-not-in-manifest')
  })
})
