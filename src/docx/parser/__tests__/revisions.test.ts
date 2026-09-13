import { describe, expect, it } from 'vitest'

import { type Block, type Paragraph } from '../../model'
import { parseDocument } from '..'

const DOCX_NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
].join(' ')

function documentXml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DOCX_NAMESPACES}>
  <w:body>
    ${body}
  </w:body>
</w:document>`
}

function expectParagraph(block: Block): Paragraph {
  if (block.kind !== 'paragraph') {
    throw new Error('Expected paragraph')
  }
  return block
}

describe('parseDocument - track changes', () => {
  it('parses w:ins as InsRevision with id/author/date and Run children', () => {
    const xml = documentXml(
      `<w:p>
        <w:ins w:id="1" w:author="Alice" w:date="2024-01-02T03:04:05Z">
          <w:r><w:t>inserted</w:t></w:r>
        </w:ins>
      </w:p>`,
    )
    const doc = parseDocument(xml)
    const para = expectParagraph(doc.sections[0].blocks[0])
    expect(para.children).toHaveLength(1)
    const child = para.children[0]
    expect(child.kind).toBe('ins-revision')
    if (child.kind !== 'ins-revision') return
    expect(child.id).toBe('1')
    expect(child.author).toBe('Alice')
    expect(child.date).toBe('2024-01-02T03:04:05Z')
    expect(child.children).toHaveLength(1)
    const run = child.children[0]
    expect(run.children[0]).toEqual({ kind: 'text', value: 'inserted' })
  })

  it('parses w:del as DelRevision and reads w:delText as text', () => {
    const xml = documentXml(
      `<w:p>
        <w:del w:id="2" w:author="Bob" w:date="2024-02-03T04:05:06Z">
          <w:r><w:delText>removed</w:delText></w:r>
        </w:del>
      </w:p>`,
    )
    const doc = parseDocument(xml)
    const para = expectParagraph(doc.sections[0].blocks[0])
    const child = para.children[0]
    expect(child.kind).toBe('del-revision')
    if (child.kind !== 'del-revision') return
    expect(child.id).toBe('2')
    expect(child.author).toBe('Bob')
    expect(child.children[0].children[0]).toEqual({ kind: 'text', value: 'removed' })
  })
})
