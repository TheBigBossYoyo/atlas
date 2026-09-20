import { describe, expect, it } from 'vitest'

import { parseHeader } from '../../parser/headers'
import type { Document, Header, Paragraph } from '../../model'
import { writeHeaderXml } from '../headerWriter'
import { buildHeaderFooterSegmentEdit, listHeaderFooterParts, setHeaderFooterBlockText } from '../../editor/headerFooter'
import { applyCommand } from '../../editor/commands'

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

  // D29 follow-up 2 — the header/footer editor now edits the TEXT SEGMENTS
  // of a paragraph that mixes plain text with a drawing/field/tab, rather
  // than treating the whole paragraph as one untouchable placeholder. This
  // is the real round trip that matters: one paragraph holding a logo
  // drawing, a text run, a tab, and a complex (`w:fldChar`/`w:instrText`)
  // PAGE field, with the text edited — the drawing, the field's own
  // instruction-text runs (in order), and the tab must come out
  // byte-identical, and only the text may change.
  it('editing the text segments of a mixed paragraph (drawing + text + tab + complex PAGE field) leaves the drawing, the field, and the tab byte-identical', () => {
    // The complex field form: begin/instrText/separate/result/end as SIBLING
    // `w:r` elements (see `groupComplexFieldRuns`'s doc comment in the
    // parser) — the shape a real Word document uses, and the one this
    // editor must never come apart.
    const fieldXml =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
      + '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
      + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
      + '<w:r><w:t>1</w:t></w:r>'
      + '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

    const sourceXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
      + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
      + ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
      + ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
      + ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
      + '<w:p>'
      + '<w:r><w:drawing><wp:inline><wp:extent cx="190500" cy="190500"/>'
      + '<wp:docPr id="1" name="Logo"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
      + '<pic:pic><pic:blipFill><a:blip r:embed="rIdLogo1"/></pic:blipFill></pic:pic>'
      + '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
      + '<w:r><w:t xml:space="preserve">Chapter Title</w:t></w:r>'
      + '<w:r><w:tab/></w:r>'
      + fieldXml
      + '</w:p>'
      + '</w:hdr>'

    const header = parseHeader(sourceXml, 'rId3')
    expect(header.blocks).toHaveLength(1)

    // The drawing's CANONICAL serialized form — extracted from the
    // unedited header's own re-serialization rather than hand-copied from
    // `sourceXml`, since the writer legitimately re-emits `a:graphicData`'s
    // spec-required `uri` attribute even for an untouched drawing (DXS-09
    // canonicalization, unrelated to this feature). What must hold is that
    // this exact substring survives editing the paragraph's text — not that
    // it matches the original bytes on disk.
    const beforeXml = writeHeaderXml(header)
    const drawingXml = beforeXml.slice(beforeXml.indexOf('<w:drawing>'), beforeXml.indexOf('</w:drawing>') + '</w:drawing>'.length)
    expect(drawingXml).toContain('rIdLogo1')

    const stubDocument: Document = {
      kind: 'document',
      sections: [{ kind: 'section', props: { headerReference: [{ id: 'rId3', type: 'default' }] }, blocks: [] }],
      styles: new Map(),
      numbering: new Map(),
      comments: new Map(),
      footnotes: new Map(),
      endnotes: new Map(),
      headers: new Map([['rId3', header]]),
      footers: new Map(),
    }

    // Sanity check the reading side first: this is a 'mixed' row (drawing,
    // then editable text, then the field), not a placeholder.
    const rowsBefore = listHeaderFooterParts(stubDocument)[0].rows
    expect(rowsBefore).toEqual([
      {
        kind: 'mixed',
        blockIndex: 0,
        segments: [
          { kind: 'atom', label: '[Image]' },
          { kind: 'text', segmentIndex: 0, text: 'Chapter Title\t' },
          { kind: 'atom', label: '[Page number]' },
        ],
      },
    ])

    const command = buildHeaderFooterSegmentEdit(stubDocument, 'header', 'rId3', 0, 0, 'Executive Summary\t')!
    const applied = applyCommand(stubDocument, command)
    const written = writeHeaderXml(applied.document.headers.get('rId3')!)

    // The text changed...
    expect(written).toContain('Executive Summary')
    expect(written).not.toContain('Chapter Title')
    // ...and everything else is byte-identical to the source.
    expect(written).toContain(drawingXml)
    expect(written).toContain(fieldXml)
    expect(written).toContain('<w:tab/>')

    // Undo restores the original paragraph object exactly (one undo step).
    const undone = applyCommand(applied.document, applied.inverse)
    expect(undone.document.headers.get('rId3')!.blocks[0]).toBe(header.blocks[0])
    expect(writeHeaderXml(undone.document.headers.get('rId3')!)).toContain(drawingXml)
    expect(writeHeaderXml(undone.document.headers.get('rId3')!)).toContain('Chapter Title')
  })
})
