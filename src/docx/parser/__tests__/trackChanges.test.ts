import { describe, expect, it } from 'vitest'

import { type Block, type Paragraph, type ParagraphChild } from '../../model'
import { parseDocument } from '../document'

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

function getRevisionChildren(
  child: Extract<Paragraph['children'][number], { kind: 'ins-revision' | 'del-revision' }>,
): ReadonlyArray<ParagraphChild> {
  return child.children as ReadonlyArray<ParagraphChild>
}

describe('parseDocument revisions', () => {
  it('parses paragraph children inside w:ins and w:del', () => {
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
    `)

    const document = parseDocument(xml)
    const paragraph = expectParagraph(document.sections[0].blocks[0])
    const insertion = paragraph.children[0]
    const deletion = paragraph.children[1]

    expect(insertion.kind).toBe('ins-revision')
    expect(deletion.kind).toBe('del-revision')

    if (insertion.kind !== 'ins-revision' || deletion.kind !== 'del-revision') {
      return
    }

    const insertionChildren = getRevisionChildren(insertion)
    expect(insertionChildren[0]?.kind).toBe('hyperlink')
    expect(insertionChildren[1]?.kind).toBe('run')

    const deletionChildren = getRevisionChildren(deletion)
    expect(deletionChildren[0]).toMatchObject({
      kind: 'run',
      children: [{ kind: 'text', value: 'removed' }],
    })
  })
})
