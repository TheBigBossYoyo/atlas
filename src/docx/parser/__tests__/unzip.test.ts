/**
 * Tests for src/docx/parser/unzip.ts
 *
 * Builds a minimal in-memory DOCX via JSZip to avoid binary fixture files.
 */

import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { unzipDocx, DocxParseError } from '../unzip'

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>`

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
  </w:body>
</w:document>`

async function buildMinimalDocx(): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', CONTENT_TYPES_XML)
  zip.file('_rels/.rels', RELS_XML)
  zip.file('word/document.xml', DOCUMENT_XML)
  const buffer = await zip.generateAsync({ type: 'uint8array' })
  return buffer
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('unzipDocx', () => {
  it('returns a DocxArchive with the expected file list', async () => {
    const bytes = await buildMinimalDocx()
    const archive = await unzipDocx(bytes)

    expect(archive.files.has('[Content_Types].xml')).toBe(true)
    expect(archive.files.has('_rels/.rels')).toBe(true)
    expect(archive.files.has('word/document.xml')).toBe(true)
  })

  it('returns Uint8Array values for each entry', async () => {
    const bytes = await buildMinimalDocx()
    const archive = await unzipDocx(bytes)

    for (const [, value] of archive.files) {
      expect(value).toBeInstanceOf(Uint8Array)
    }
  })

  it('accepts an ArrayBuffer as input', async () => {
    const bytes = await buildMinimalDocx()
    const archive = await unzipDocx(bytes.buffer as ArrayBuffer)

    expect(archive.files.has('word/document.xml')).toBe(true)
  })

  it('content of word/document.xml round-trips correctly', async () => {
    const bytes = await buildMinimalDocx()
    const archive = await unzipDocx(bytes)

    const docBytes = archive.files.get('word/document.xml')!
    const decoded = new TextDecoder().decode(docBytes)
    expect(decoded).toContain('<w:t>Hello</w:t>')
  })

  it('throws DocxParseError for corrupt input', async () => {
    const corrupt = new Uint8Array([0x00, 0x01, 0x02, 0x03])
    await expect(unzipDocx(corrupt)).rejects.toBeInstanceOf(DocxParseError)
  })

  it('DocxParseError has the correct name', async () => {
    const corrupt = new Uint8Array([0xff, 0xfe])
    await expect(unzipDocx(corrupt)).rejects.toMatchObject({ name: 'DocxParseError' })
  })
})
