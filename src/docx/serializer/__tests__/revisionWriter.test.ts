import { describe, expect, it } from 'vitest'

import { parseDocument } from '../../parser'
import { writeDocumentXml } from '../documentWriter'

const DOCX_NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
].join(' ')

function documentXml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${DOCX_NAMESPACES}><w:body>${body}</w:body></w:document>`
}

function normalizeAst(value: unknown): unknown {
  if (value instanceof Map) {
    return Array.from(value.entries(), ([key, entry]) => [key, normalizeAst(entry)])
  }

  if (Array.isArray(value)) {
    return value.map((entry) => normalizeAst(entry))
  }

  if (value !== null && typeof value === 'object') {
    const record = value as Record<string, unknown>
    const normalized: Record<string, unknown> = {}

    for (const [key, entry] of Object.entries(record)) {
      if (record.kind === 'text' && key === 'preserveSpace') {
        continue
      }

      // Source-only capture (D19 / DXS-15) — see documentWriter.test.ts's
      // identical exclusion for why a round-trip legitimately gains these.
      if (record.kind === 'document' && (key === 'rootNamespaces' || key === 'mcIgnorable')) {
        continue
      }

      normalized[key] = normalizeAst(entry)
    }

    return normalized
  }

  return value
}

describe('writeDocumentXml revisions', () => {
  it('round-trips track changes with hyperlink insertions and deleted text', () => {
    const xml = documentXml(`
      <w:p>
        <w:ins w:id="1" w:author="Alice" w:date="2024-01-02T03:04:05Z">
          <w:hyperlink r:id="rId1">
            <w:r><w:t>linked</w:t></w:r>
          </w:hyperlink>
          <w:r><w:t>tail</w:t></w:r>
        </w:ins>
        <w:del w:id="2" w:author="Bob" w:date="2024-02-03T04:05:06Z">
          <w:r><w:delText>removed</w:delText></w:r>
        </w:del>
      </w:p>
      <w:sectPr/>
    `)

    const parsed = parseDocument(xml)
    const written = writeDocumentXml(parsed)
    const reparsed = parseDocument(written)

    expect(written).toContain('<w:ins w:id="1" w:author="Alice" w:date="2024-01-02T03:04:05Z">')
    expect(written).toContain('<w:delText xml:space="preserve">removed</w:delText>')
    expect(normalizeAst(reparsed)).toEqual(normalizeAst(parsed))
  })
})
