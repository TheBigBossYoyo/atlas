/**
 * Tests for src/docx/fonts/referencedFamilies.ts (D22 / DXP-18)
 */
import { describe, expect, it } from 'vitest'

import {
  collectReferencedFontFamilies,
  resolveReferencedFontFamilies,
  resolveUnreferencedFontFamilies,
} from '../referencedFamilies'
import { FONT_FAMILIES } from '../families'
import type {
  Comment,
  Document,
  Endnote,
  Footer,
  Footnote,
  Header,
  Paragraph,
  Run,
  RunProps,
  Section,
  Style,
  Table,
} from '../../model'

function run(rFonts?: RunProps['rFonts']): Run {
  return {
    kind: 'run',
    ...(rFonts !== undefined ? { props: { rFonts } } : {}),
    children: [{ kind: 'text', value: 'x' }],
  }
}

function paragraph(...children: Paragraph['children']): Paragraph {
  return { kind: 'paragraph', children }
}

function section(...blocks: Section['blocks']): Section {
  return { kind: 'section', props: {}, blocks }
}

function baseDocument(overrides: Partial<Document> = {}): Document {
  return {
    kind: 'document',
    sections: [],
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

describe('collectReferencedFontFamilies', () => {
  it('returns nothing for a document with no font formatting anywhere', () => {
    const document = baseDocument({ sections: [section(paragraph(run()))] })
    expect(collectReferencedFontFamilies(document)).toEqual([])
  })

  it('collects a font referenced by direct run formatting', () => {
    const document = baseDocument({
      sections: [section(paragraph(run({ ascii: 'Times New Roman' })))],
    })
    expect(collectReferencedFontFamilies(document)).toEqual(['Times New Roman'])
  })

  it('collects distinct ascii/hAnsi/cs/eastAsia fonts from the same run', () => {
    const document = baseDocument({
      sections: [
        section(
          paragraph(
            run({ ascii: 'Arial', hAnsi: 'Arial', cs: 'Times New Roman', eastAsia: 'MS Mincho' }),
          ),
        ),
      ],
    })
    expect(new Set(collectReferencedFontFamilies(document))).toEqual(
      new Set(['Arial', 'Times New Roman', 'MS Mincho']),
    )
  })

  it('collects a font from a run inside a hyperlink', () => {
    const document = baseDocument({
      sections: [
        section(
          paragraph({
            kind: 'hyperlink',
            children: [run({ ascii: 'Courier New' })],
          }),
        ),
      ],
    })
    expect(collectReferencedFontFamilies(document)).toEqual(['Courier New'])
  })

  it('collects a font from a run inside an ins/del revision', () => {
    const document = baseDocument({
      sections: [
        section(
          paragraph({
            kind: 'ins-revision',
            id: '1',
            children: [run({ ascii: 'Cambria' })],
          }),
        ),
      ],
    })
    expect(collectReferencedFontFamilies(document)).toEqual(['Cambria'])
  })

  it('collects a font from a run inside a table cell', () => {
    const table: Table = {
      kind: 'table',
      rows: [
        {
          kind: 'table-row',
          cells: [
            {
              kind: 'table-cell',
              blocks: [paragraph(run({ ascii: 'Calibri' }))],
            },
          ],
        },
      ],
    }
    const document = baseDocument({ sections: [section(table)] })
    expect(collectReferencedFontFamilies(document)).toEqual(['Calibri'])
  })

  it('collects a font from document defaults', () => {
    const document = baseDocument({
      defaults: { run: { rFonts: { ascii: 'Georgia' } } },
    })
    expect(collectReferencedFontFamilies(document)).toEqual(['Georgia'])
  })

  it('collects a font from a style\'s own run formatting', () => {
    const style: Style = {
      id: 'Heading1',
      type: 'paragraph',
      run: { rFonts: { ascii: 'Verdana' } },
    }
    const document = baseDocument({ styles: new Map([['Heading1', style]]) })
    expect(collectReferencedFontFamilies(document)).toEqual(['Verdana'])
  })

  it('collects fonts from headers, footers, footnotes, endnotes, and comments', () => {
    const header: Header = { kind: 'header', id: 'h1', blocks: [paragraph(run({ ascii: 'A' }))] }
    const footer: Footer = { kind: 'footer', id: 'f1', blocks: [paragraph(run({ ascii: 'B' }))] }
    const footnote: Footnote = { kind: 'footnote', id: '1', blocks: [paragraph(run({ ascii: 'C' }))] }
    const endnote: Endnote = { kind: 'endnote', id: '1', blocks: [paragraph(run({ ascii: 'D' }))] }
    const comment: Comment = { kind: 'comment', id: '1', body: [paragraph(run({ ascii: 'E' }))] }

    const document = baseDocument({
      headers: new Map([['h1', header]]),
      footers: new Map([['f1', footer]]),
      footnotes: new Map([['1', footnote]]),
      endnotes: new Map([['1', endnote]]),
      comments: new Map([['1', comment]]),
    })

    expect(new Set(collectReferencedFontFamilies(document))).toEqual(new Set(['A', 'B', 'C', 'D', 'E']))
  })
})

describe('resolveReferencedFontFamilies', () => {
  it('resolves referenced names to their bundled substitute, deduplicated', () => {
    const document = baseDocument({
      sections: [
        section(
          paragraph(run({ ascii: 'Times New Roman' })),
          paragraph(run({ ascii: 'Times' })), // alias for the same family
        ),
      ],
    })

    expect(resolveReferencedFontFamilies(document)).toEqual([FONT_FAMILIES[3]])
  })

  it('skips a referenced font with no bundled substitute', () => {
    const document = baseDocument({
      sections: [section(paragraph(run({ ascii: 'Comic Sans MS' })))],
    })
    expect(resolveReferencedFontFamilies(document)).toEqual([])
  })

  it('returns an empty array for a document with no font formatting', () => {
    expect(resolveReferencedFontFamilies(baseDocument())).toEqual([])
  })
})

describe('resolveUnreferencedFontFamilies', () => {
  it('returns every bundled family for a document referencing none of them', () => {
    expect(resolveUnreferencedFontFamilies(baseDocument())).toEqual(FONT_FAMILIES)
  })

  it('excludes the one family a document references, keeping the rest', () => {
    const document = baseDocument({
      sections: [section(paragraph(run({ ascii: 'Calibri' })))],
    })
    const unreferenced = resolveUnreferencedFontFamilies(document)

    expect(unreferenced).not.toContain(FONT_FAMILIES[0])
    expect(unreferenced).toHaveLength(FONT_FAMILIES.length - 1)
  })
})
