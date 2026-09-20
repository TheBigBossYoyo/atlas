import { describe, expect, it } from 'vitest'

import { htmlToPasteBlocks, type PasteParagraph, type PasteTable } from '../pasteBlocks'

function paragraphsOf(html: string): ReadonlyArray<PasteParagraph> {
  return htmlToPasteBlocks(html).filter((b): b is PasteParagraph => b.kind === 'paragraph')
}

describe('htmlToPasteBlocks — paragraphs and inline formatting', () => {
  it('returns an empty array for the empty string', () => {
    expect(htmlToPasteBlocks('')).toEqual([])
  })

  it('parses two <p> blocks into two paragraphs', () => {
    const blocks = paragraphsOf('<p>Hello</p><p>World</p>')
    expect(blocks).toHaveLength(2)
    expect(blocks[0].runs).toEqual([{ kind: 'text', text: 'Hello', format: {} }])
    expect(blocks[1].runs).toEqual([{ kind: 'text', text: 'World', format: {} }])
  })

  it('captures bold/italic/underline formatting from tags', () => {
    const [p] = paragraphsOf('<p>plain <b>bold</b> <i>italic</i> <u>under</u></p>')
    expect(p.runs.find((r) => r.kind === 'text' && r.text.trim() === 'bold')).toMatchObject({
      format: { bold: true },
    })
    expect(p.runs.find((r) => r.kind === 'text' && r.text.trim() === 'italic')).toMatchObject({
      format: { italic: true },
    })
    expect(p.runs.find((r) => r.kind === 'text' && r.text.trim() === 'under')).toMatchObject({
      format: { underline: { style: 'single' } },
    })
  })

  it('treats <br> as a paragraph break inside a block', () => {
    const blocks = paragraphsOf('<p>line one<br>line two</p>')
    expect(blocks).toHaveLength(2)
    expect(blocks[0].runs[0]).toMatchObject({ text: 'line one' })
    expect(blocks[1].runs[0]).toMatchObject({ text: 'line two' })
  })

  it('skips <script> and <style> content', () => {
    const blocks = paragraphsOf('<style>.x{color:red}</style><p>visible</p><script>alert(1)</script>')
    expect(blocks).toHaveLength(1)
    expect(blocks[0].runs[0]).toMatchObject({ text: 'visible' })
  })

  it('collapses consecutive whitespace in text content', () => {
    const [p] = paragraphsOf('<p>foo   \n   bar</p>')
    expect(p.runs[0]).toMatchObject({ text: 'foo bar' })
  })
})

describe('htmlToPasteBlocks — DXE-19 color/highlight fidelity', () => {
  it('parses an inline style color into a hex run format', () => {
    const [p] = paragraphsOf('<p><span style="color: rgb(255, 0, 0)">red text</span></p>')
    // DOCX-14 — ST_HexColor never carries a leading '#'; hexColor() strips it.
    expect(p.runs[0]).toMatchObject({ format: { color: 'ff0000' } })
  })

  it('maps a background-color matching a Word highlight swatch to `highlight`', () => {
    const [p] = paragraphsOf('<p><span style="background-color: #ffff00">hi</span></p>')
    expect(p.runs[0]).toMatchObject({ format: { highlight: 'yellow' } })
  })

  it('falls back to shading for a background-color outside the highlight palette', () => {
    const [p] = paragraphsOf('<p><span style="background-color: #123456">hi</span></p>')
    expect(p.runs[0]).toMatchObject({ format: { shd: { fill: '123456' } } })
  })

  it('reads a legacy <font color> attribute', () => {
    const [p] = paragraphsOf('<p><font color="#00ff00">green</font></p>')
    expect(p.runs[0]).toMatchObject({ format: { color: '00ff00' } })
  })

  it('ignores an unparseable color value', () => {
    const [p] = paragraphsOf('<p><span style="color: currentColor">x</span></p>')
    expect(p.runs[0]).toMatchObject({ format: {} })
  })
})

describe('htmlToPasteBlocks — DXE-19 hyperlinks', () => {
  it('attaches href to runs inside a safe-scheme <a>', () => {
    const [p] = paragraphsOf('<p><a href="https://example.com/atlas">link</a></p>')
    expect(p.runs[0]).toMatchObject({ text: 'link', href: 'https://example.com/atlas' })
  })

  it('accepts a mailto: link', () => {
    const [p] = paragraphsOf('<p><a href="mailto:a@b.com">mail</a></p>')
    expect(p.runs[0]).toMatchObject({ href: 'mailto:a@b.com' })
  })

  it('drops the href for an unsafe scheme, keeping the plain text', () => {
    const [p] = paragraphsOf('<p><a href="javascript:alert(1)">bad</a></p>')
    expect(p.runs[0]).toMatchObject({ text: 'bad' })
    expect((p.runs[0] as { href?: string }).href).toBeUndefined()
  })

  it('drops the href for a relative (schemeless) link', () => {
    const [p] = paragraphsOf('<p><a href="/relative">rel</a></p>')
    expect((p.runs[0] as { href?: string }).href).toBeUndefined()
  })
})

describe('htmlToPasteBlocks — DXE-19 lists', () => {
  it('marks each <li> paragraph with its list kind', () => {
    const blocks = paragraphsOf('<ul><li>one</li><li>two</li></ul>')
    expect(blocks).toHaveLength(2)
    expect(blocks[0].list).toEqual({ ordered: false, level: 0 })
    expect(blocks[1].list).toEqual({ ordered: false, level: 0 })
  })

  it('distinguishes ordered from unordered lists', () => {
    const [p] = paragraphsOf('<ol><li>one</li></ol>')
    expect(p.list).toMatchObject({ ordered: true })
  })

  it('increments level for a nested list', () => {
    const blocks = paragraphsOf('<ul><li>outer<ul><li>inner</li></ul></li></ul>')
    const inner = blocks.find((b) => b.runs.some((r) => r.kind === 'text' && r.text === 'inner'))
    expect(inner?.list).toMatchObject({ level: 1 })
  })

  it('a plain paragraph outside any list carries no list info', () => {
    const [p] = paragraphsOf('<p>plain</p>')
    expect(p.list).toBeUndefined()
  })
})

describe('htmlToPasteBlocks — DXE-19 images', () => {
  it('parses a data:image/png <img> as an image run', () => {
    const blocks = paragraphsOf('<p><img src="data:image/png;base64,iVBORw0KGgo="></p>')
    expect(blocks[0].runs).toEqual([{ kind: 'image', dataUrl: 'data:image/png;base64,iVBORw0KGgo=' }])
  })

  it('ignores an http(s) <img src> (no filesystem/network read of untrusted clipboard HTML)', () => {
    const blocks = paragraphsOf('<p><img src="https://example.com/pic.png"></p>')
    expect(blocks).toHaveLength(0)
  })

  it('ignores a file: <img src>', () => {
    const blocks = paragraphsOf('<p><img src="file:///C:/secret.png"></p>')
    expect(blocks).toHaveLength(0)
  })

  it('ignores an unsupported data: image mime type', () => {
    const blocks = paragraphsOf('<p><img src="data:image/svg+xml;base64,PHN2Zz48L3N2Zz4="></p>')
    expect(blocks).toHaveLength(0)
  })
})

describe('htmlToPasteBlocks — DXE-19 tables', () => {
  it('parses rows/cells with their text content', () => {
    const html = '<table><tr><td>A1</td><td>B1</td></tr><tr><td>A2</td><td>B2</td></tr></table>'
    const [table] = htmlToPasteBlocks(html) as [PasteTable]
    expect(table.kind).toBe('table')
    expect(table.rows).toHaveLength(2)
    expect(table.rows[0].cells).toHaveLength(2)
    expect(table.rows[0].cells[0].paragraphs[0].runs[0]).toMatchObject({ text: 'A1' })
    expect(table.rows[1].cells[1].paragraphs[0].runs[0]).toMatchObject({ text: 'B2' })
  })

  it('reads a colspan attribute into colSpan', () => {
    const html = '<table><tr><td colspan="2">Merged</td></tr></table>'
    const [table] = htmlToPasteBlocks(html) as [PasteTable]
    expect(table.rows[0].cells[0].colSpan).toBe(2)
  })

  it('defaults colSpan to 1 when absent or invalid', () => {
    const html = '<table><tr><td>A</td><td colspan="bogus">B</td></tr></table>'
    const [table] = htmlToPasteBlocks(html) as [PasteTable]
    expect(table.rows[0].cells[0].colSpan).toBe(1)
    expect(table.rows[0].cells[1].colSpan).toBe(1)
  })

  it('finds rows inside a <tbody>/<thead>', () => {
    const html = '<table><thead><tr><th>H</th></tr></thead><tbody><tr><td>D</td></tr></tbody></table>'
    const [table] = htmlToPasteBlocks(html) as [PasteTable]
    expect(table.rows).toHaveLength(2)
    expect(table.rows[0].cells[0].paragraphs[0].runs[0]).toMatchObject({ text: 'H' })
    expect(table.rows[1].cells[0].paragraphs[0].runs[0]).toMatchObject({ text: 'D' })
  })

  it('gives an empty cell one empty paragraph rather than none', () => {
    const html = '<table><tr><td></td></tr></table>'
    const [table] = htmlToPasteBlocks(html) as [PasteTable]
    expect(table.rows[0].cells[0].paragraphs).toEqual([{ kind: 'paragraph', runs: [] }])
  })

  it('a table followed by more paragraph content parses both, in order', () => {
    const html = '<p>before</p><table><tr><td>cell</td></tr></table><p>after</p>'
    const blocks = htmlToPasteBlocks(html)
    expect(blocks.map((b) => b.kind)).toEqual(['paragraph', 'table', 'paragraph'])
  })

  it('returns nothing for an empty <table>', () => {
    expect(htmlToPasteBlocks('<table></table>')).toEqual([])
  })
})
