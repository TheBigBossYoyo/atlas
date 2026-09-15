import { describe, expect, it } from 'vitest'

import type { Document, Field, Footer, Header, Paragraph, Section } from '../../model'
import { updateFields } from '../updateFields'
import type { FieldEvaluationContext } from '../types'

function makeDocument(
  sections: ReadonlyArray<Section>,
  overrides: Partial<Document> = {},
): Document {
  return {
    kind: 'document',
    sections,
    styles: new Map(),
    numbering: new Map(),
    comments: new Map(),
    footnotes: new Map(),
    endnotes: new Map(),
    headers: new Map(),
    footers: new Map(),
    ...overrides,
  }
}

function makeSection(blocks: Section['blocks']): Section {
  return { kind: 'section', props: {}, blocks }
}

function makeFieldParagraph(field: Field): Paragraph {
  return { kind: 'paragraph', children: [field] }
}

function makeField(overrides: Partial<Field> & Pick<Field, 'fieldType' | 'instruction'>): Field {
  return { kind: 'field', result: [{ kind: 'run', children: [{ kind: 'text', value: 'stale' }] }], raw: '<stale/>', ...overrides }
}

function makeContext(overrides: Partial<FieldEvaluationContext> = {}): FieldEvaluationContext {
  return { bookmarkText: new Map(), sequenceCounters: new Map(), ...overrides }
}

function fieldAt(document: Document, sectionIndex = 0, blockIndex = 0): Field {
  const block = document.sections[sectionIndex]?.blocks[blockIndex]
  if (block?.kind !== 'paragraph' || block.children[0]?.kind !== 'field') {
    throw new Error('Expected a field at the given position')
  }
  return block.children[0]
}

describe('updateFields', () => {
  it('replaces a resolvable field\'s cached result and clears its raw passthrough', () => {
    const field = makeField({ fieldType: 'AUTHOR', instruction: 'AUTHOR' })
    const document = makeDocument([makeSection([makeFieldParagraph(field)])])

    const { document: updated, updatedCount } = updateFields(document, makeContext({ author: 'A. Author' }))

    expect(updatedCount).toBe(1)
    const updatedField = fieldAt(updated)
    expect(updatedField.raw).toBeUndefined()
    expect(updatedField.result).toEqual([{ kind: 'run', children: [{ kind: 'text', value: 'A. Author' }] }])
  })

  it('leaves an unevaluable field (missing context data) completely unchanged, raw included', () => {
    const field = makeField({ fieldType: 'REF', instruction: 'REF _Unknown' })
    const document = makeDocument([makeSection([makeFieldParagraph(field)])])

    const { document: updated, updatedCount } = updateFields(document, makeContext())

    expect(updatedCount).toBe(0)
    expect(fieldAt(updated)).toEqual(field)
  })

  it('preserves the field\'s existing run formatting (props) from its old cached result', () => {
    const field = makeField({
      fieldType: 'AUTHOR',
      instruction: 'AUTHOR',
      result: [{ kind: 'run', props: { bold: true }, children: [{ kind: 'text', value: 'old' }] }],
    })
    const document = makeDocument([makeSection([makeFieldParagraph(field)])])

    const { document: updated } = updateFields(document, makeContext({ author: 'New Author' }))

    expect(fieldAt(updated).result).toEqual([
      { kind: 'run', props: { bold: true }, children: [{ kind: 'text', value: 'New Author' }] },
    ])
  })

  it('leaves a locked (w:fldLock) field completely unchanged, raw included, matching Word\'s own Update Field(s)', () => {
    const field = makeField({ fieldType: 'AUTHOR', instruction: 'AUTHOR', locked: true })
    const document = makeDocument([makeSection([makeFieldParagraph(field)])])

    const { document: updated, updatedCount } = updateFields(document, makeContext({ author: 'A. Author' }))

    expect(updatedCount).toBe(0)
    expect(fieldAt(updated)).toEqual(field)
  })

  it('never touches a TOC field (updateTableOfContents handles those separately)', () => {
    const field = makeField({ fieldType: 'TOC', instruction: 'TOC \\o "1-3"' })
    const document = makeDocument([makeSection([makeFieldParagraph(field)])])

    const { document: updated, updatedCount } = updateFields(document, makeContext())

    expect(updatedCount).toBe(0)
    expect(fieldAt(updated)).toEqual(field)
  })

  it('evaluates a PAGE field in the body using currentPageOf, addressed by [sectionIndex, blockIndex]', () => {
    const field = makeField({ fieldType: 'PAGE', instruction: 'PAGE' })
    const document = makeDocument([
      makeSection([{ kind: 'paragraph', children: [] }, makeFieldParagraph(field)]),
    ])

    const context = makeContext({ currentPageOf: (path) => (path[0] === 0 && path[1] === 1 ? 3 : undefined) })
    const { document: updated } = updateFields(document, context)

    expect(fieldAt(updated, 0, 1).result).toEqual([{ kind: 'run', children: [{ kind: 'text', value: '3' }] }])
  })

  it('does not evaluate PAGE/NUMPAGES fields inside a header (no single fixed page number applies there)', () => {
    const pageField = makeField({ fieldType: 'PAGE', instruction: 'PAGE' })
    const header: Header = { kind: 'header', id: 'h1', blocks: [makeFieldParagraph(pageField)] }
    const document = makeDocument([makeSection([])], { headers: new Map([['h1', header]]) })

    const context = makeContext({ currentPageOf: () => 7, pageCount: 9 })
    const { document: updated, updatedCount } = updateFields(document, context)

    expect(updatedCount).toBe(0)
    const updatedHeader = updated.headers.get('h1')
    expect(updatedHeader?.blocks[0]).toEqual(makeFieldParagraph(pageField))
  })

  it('does evaluate a non-page field (e.g. AUTHOR) inside a footer', () => {
    const authorField = makeField({ fieldType: 'AUTHOR', instruction: 'AUTHOR' })
    const footer: Footer = { kind: 'footer', id: 'f1', blocks: [makeFieldParagraph(authorField)] }
    const document = makeDocument([makeSection([])], { footers: new Map([['f1', footer]]) })

    const { document: updated, updatedCount } = updateFields(document, makeContext({ author: 'Foot Author' }))

    expect(updatedCount).toBe(1)
    const block = updated.footers.get('f1')?.blocks[0]
    const updatedField = block?.kind === 'paragraph' ? block.children[0] : undefined
    expect(updatedField?.kind === 'field' && updatedField.result[0]).toEqual({
      kind: 'run',
      children: [{ kind: 'text', value: 'Foot Author' }],
    })
  })

  it('recalculates a field nested inside a hyperlink', () => {
    const field = makeField({ fieldType: 'AUTHOR', instruction: 'AUTHOR' })
    const paragraph: Paragraph = {
      kind: 'paragraph',
      children: [{ kind: 'hyperlink', anchor: 'Top', children: [field] }],
    }
    const document = makeDocument([makeSection([paragraph])])

    const { document: updated } = updateFields(document, makeContext({ author: 'Hyperlinked Author' }))

    const block = updated.sections[0].blocks[0]
    const hyperlink = block.kind === 'paragraph' ? block.children[0] : undefined
    const updatedField = hyperlink?.kind === 'hyperlink' ? hyperlink.children[0] : undefined
    expect(updatedField?.kind === 'field' && updatedField.result[0]).toEqual({
      kind: 'run',
      children: [{ kind: 'text', value: 'Hyperlinked Author' }],
    })
  })

  it('recalculates a field nested inside a tracked insertion (w:ins)', () => {
    const field = makeField({ fieldType: 'AUTHOR', instruction: 'AUTHOR' })
    const paragraph: Paragraph = {
      kind: 'paragraph',
      children: [{ kind: 'ins-revision', id: '1', children: Object.freeze([field]) as never }],
    }
    const document = makeDocument([makeSection([paragraph])])

    const { document: updated } = updateFields(document, makeContext({ author: 'Ins Author' }))

    const block = updated.sections[0].blocks[0]
    const revision = block.kind === 'paragraph' ? block.children[0] : undefined
    const revisionChildren =
      revision?.kind === 'ins-revision' ? (revision.children as ReadonlyArray<Field>) : undefined
    const updatedField = revisionChildren?.[0]
    expect(updatedField?.kind === 'field' && updatedField.result[0]).toEqual({
      kind: 'run',
      children: [{ kind: 'text', value: 'Ins Author' }],
    })
  })

  it('recalculates a field inside a table cell, but never evaluates PAGE there (no section-relative address for a cell)', () => {
    const authorField = makeField({ fieldType: 'AUTHOR', instruction: 'AUTHOR' })
    const pageField = makeField({ fieldType: 'PAGE', instruction: 'PAGE' })
    const document = makeDocument([
      makeSection([
        {
          kind: 'table',
          rows: [
            {
              kind: 'table-row',
              cells: [
                {
                  kind: 'table-cell',
                  blocks: [makeFieldParagraph(authorField), makeFieldParagraph(pageField)],
                },
              ],
            },
          ],
        },
      ]),
    ])

    const context = makeContext({ author: 'Cell Author', currentPageOf: () => 5 })
    const { document: updated, updatedCount } = updateFields(document, context)

    expect(updatedCount).toBe(1)
    const table = updated.sections[0].blocks[0]
    const cell = table.kind === 'table' && table.rows[0]?.kind === 'table-row' ? table.rows[0].cells[0] : undefined
    const cellBlocks = cell?.kind === 'table-cell' ? cell.blocks : []
    const updatedAuthorField = cellBlocks[0].kind === 'paragraph' ? cellBlocks[0].children[0] : undefined
    expect(updatedAuthorField?.kind === 'field' && updatedAuthorField.result[0]).toEqual({
      kind: 'run',
      children: [{ kind: 'text', value: 'Cell Author' }],
    })
    // PAGE is left completely untouched (no section-relative address applies inside a cell).
    expect(cellBlocks[1].kind === 'paragraph' && cellBlocks[1].children[0]).toEqual(pageField)
  })

  it('returns updatedCount 0 for a document with no fields at all', () => {
    const document = makeDocument([makeSection([{ kind: 'paragraph', children: [] }])])
    const { updatedCount } = updateFields(document, makeContext())
    expect(updatedCount).toBe(0)
  })
})
