import { describe, expect, it } from 'vitest'

import { parseHeader } from '../../parser/headers'
import type { Document, Header, Paragraph } from '../../model'
import { writeHeaderXml } from '../headerWriter'
import { setHeaderFooterBlockText } from '../../editor/headerFooter'

function makeParagraph(text: string): Paragraph {
  return {
    kind: 'paragraph',
    children: [
      {
        kind: 'run',
        children: [{ kind: 'text', value: text }],
      },
    ],
  }
}

function makeHeader(blocks: ReadonlyArray<Paragraph>): Header {
  return {
    kind: 'header',
    id: 'rId3',
    blocks,
  }
}

describe('writeHeaderXml', () => {
  it('writes an empty header part', () => {
    const xml = writeHeaderXml(makeHeader([]))

    expect(xml.startsWith('<?xml')).toBe(true)
    expect(xml).toContain('<w:hdr')
    expect(xml).not.toContain('<w:p')
  })

  it('writes a header paragraph with text', () => {
    const xml = writeHeaderXml(makeHeader([makeParagraph('Page Header')]))

    expect(xml).toContain('<w:hdr')
    expect(xml).toContain('<w:p>')
    expect(xml).toContain('Page Header')
  })

  it('writes multiple header paragraphs', () => {
    const xml = writeHeaderXml(makeHeader([makeParagraph('First'), makeParagraph('Second')]))

    expect(xml).toContain('First')
    expect(xml).toContain('Second')
  })

  it('declares the full standard namespace set, not just xmlns:w (D19 / DXS-08)', () => {
    const xml = writeHeaderXml(makeHeader([]))

    // Previously only xmlns:w (+ a hardcoded xmlns:r) was declared, so a
    // drawing/hyperlink/shape inside a header emitted an undeclared
    // namespace prefix — an XML well-formedness violation.
    for (const prefix of ['w', 'r', 'wp', 'a', 'pic', 'v', 'mc', 'w14']) {
      expect(xml).toContain(`xmlns:${prefix}=`)
    }
  })

  it(
    'substitutes real XML back in for a nested unrecognized node (e.g. w:proofErr) instead of '
      + 'leaking an unrestored atlas-raw-unknown placeholder',
    () => {
      // w:proofErr is not modeled (it becomes an UnknownNode run child) but
      // is ubiquitous in real Word-authored documents (inserted around
      // nearly every word the spell-checker flags) — a header/footer/
      // footnote/endnote/comment containing one previously round-tripped
      // to a literal, un-substituted `<atlas-raw-unknown data-id="..."/>`
      // element instead of the original raw XML, because buildParagraph's
      // placeholder/restoration mechanism only ran for the main document
      // body, never for these standalone parts.
      const sourceXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        + '<w:p><w:r><w:proofErr w:type="spellStart"/><w:t>Helo</w:t><w:proofErr w:type="spellEnd"/></w:r></w:p>'
        + '</w:hdr>'

      const header = parseHeader(sourceXml, 'rId1')
      const written = writeHeaderXml(header)

      expect(written).not.toContain('atlas-raw-unknown')
      expect(written).toContain('<w:proofErr w:type="spellStart"/>')
      expect(written).toContain('<w:proofErr w:type="spellEnd"/>')
    },
  )

  // D29 follow-up — the original plain-text header/footer editor rewrote a
  // whole part from scratch on any edit, so a header holding an image and a
  // PAGE field lost both the moment its text paragraph was touched. This
  // exercises the full round trip: parse a header with all three (image,
  // field, bold text), edit only the text paragraph through the same
  // `setHeaderFooterBlockText` the panel's commit path uses, and check the
  // re-serialized XML still has the drawing and the field.
  it('keeps a drawing and a PAGE field intact when the header text paragraph is edited', () => {
    const sourceXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
      + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
      + ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
      + ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
      + ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
      + '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/>'
      + '<wp:docPr id="1" name="Logo"/><a:graphic><a:graphicData>'
      + '<pic:pic><pic:blipFill><a:blip r:embed="rIdLogo1"/></pic:blipFill></pic:pic>'
      + '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
      + '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Confidential</w:t></w:r></w:p>'
      + '<w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>'
      + '</w:hdr>'

    const header = parseHeader(sourceXml, 'rId3')
    expect(header.blocks).toHaveLength(3)

    const stubDocument: Document = {
      kind: 'document',
      sections: [{ kind: 'section', props: {}, blocks: [] }],
      styles: new Map(),
      numbering: new Map(),
      comments: new Map(),
      footnotes: new Map(),
      endnotes: new Map(),
      headers: new Map([['rId3', header]]),
      footers: new Map(),
    }
    const edited = setHeaderFooterBlockText(stubDocument, 'header', 'rId3', 1, 'CONFIDENTIAL')

    const written = writeHeaderXml(edited.headers.get('rId3')!)

    // The edited paragraph's new text, still bold — re-parsed rather than
    // grepped for the literal substring, since the run-level diff may (as it
    // does here, an edit that only changes case) split it across more than
    // one `<w:t>`/`<w:r>` when only part of the text actually changed.
    const roundTripped = parseHeader(written, 'rId3')
    const editedParagraph = roundTripped.blocks[1] as Paragraph
    const editedRuns = editedParagraph.children as ReadonlyArray<{ kind: string; props?: { bold?: boolean }; children: ReadonlyArray<{ value: string }> }>
    expect(editedRuns.map((run) => run.children[0]?.value).join('')).toBe('CONFIDENTIAL')
    expect(editedRuns.every((run) => run.props?.bold === true)).toBe(true)
    // The drawing (image) is untouched.
    expect(written).toContain('rIdLogo1')
    expect(written).toContain('wp:inline')
    // The PAGE field is untouched.
    expect(written).toContain('w:fldSimple')
    expect(written).toContain('PAGE')
  })
})
