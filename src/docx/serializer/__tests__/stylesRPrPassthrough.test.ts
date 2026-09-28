/**
 * RPR-STYLES-1 — `styles.xml` and `numbering.xml` round-trip a `w:rPr` child
 * Atlas has no field for.
 *
 * Silent data loss on every save, not a hypothetical: `saveDocx` rebuilds
 * `word/styles.xml` from the model in full, so anything the parser did not
 * carry simply ceased to exist. `w:noProof` and `w:snapToGrid` appear on Word's
 * own built-in styles, and `w14:`/`w15:` run properties on anything written by
 * a current Word — none of which the direct-formatting path lost, because
 * `parser/document.ts` has had this passthrough since DOCX-12. Only the styles
 * path was missing it.
 *
 * These assert on the SERIALIZED XML rather than on the model, because the model
 * carrying a fragment proves nothing about whether the file gets it back.
 */
import { describe, expect, it } from 'vitest'

import { parseNumbering } from '../../parser/numbering'
import { parseStyles } from '../../parser/styles'
import { writeNumberingXml } from '../numberingWriter'
import { writeStylesXml } from '../stylesWriter'

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml'

function stylesXml(body: string, rootAttributes = ''): string {
  return `${DECL}<w:styles xmlns:w="${W_NS}"${rootAttributes}>${body}</w:styles>`
}

function roundTripStyles(xml: string): string {
  return writeStylesXml(parseStyles(xml))
}

/** A character style whose `w:rPr` is exactly `body`. */
function characterStyle(body: string): string {
  return `<w:style w:type="character" w:styleId="S"><w:rPr>${body}</w:rPr></w:style>`
}

describe('styles.xml w:rPr passthrough', () => {
  it('keeps an unmodeled run property on a named style', () => {
    const out = roundTripStyles(stylesXml(characterStyle('<w:b/><w:noProof/>')))

    expect(out).toContain('<w:noProof/>')
  })

  it('puts the unmodeled child back in schema order, not at the end', () => {
    // `CT_RPr` is a strict xsd:sequence, and `w:snapToGrid` belongs between
    // `w:position` and `w:sz`. Word reports a mis-ordered sequence as
    // unreadable content, so "kept, but in the wrong place" is not a fix.
    const out = roundTripStyles(
      stylesXml(characterStyle('<w:b/><w:position w:val="4"/><w:snapToGrid/><w:sz w:val="24"/>')),
    )

    expect(out).toContain('<w:position w:val="4"/><w:snapToGrid/><w:sz w:val="24"/>')
  })

  it('keeps several unmodeled children, each at its own position', () => {
    const out = roundTripStyles(
      stylesXml(characterStyle('<w:noProof/><w:b/><w:snapToGrid/><w:sz w:val="20"/><w:specVanish/>')),
    )

    expect(out).toContain('<w:noProof/><w:b/><w:snapToGrid/><w:sz w:val="20"/><w:specVanish/>')
  })

  it('keeps attributes and nesting of a structured unmodeled child', () => {
    const out = roundTripStyles(
      stylesXml(characterStyle('<w:eastAsianLayout w:id="1" w:combine="1" w:combineBrackets="round"/>')),
    )

    expect(out).toContain('<w:eastAsianLayout w:id="1" w:combine="1" w:combineBrackets="round"/>')
  })

  it('emits a w:rPr whose ONLY children are unmodeled ones', () => {
    // The emptiness check runs on the spliced result. Run before the splice, as
    // it used to be, and this whole element vanishes along with its content.
    const out = roundTripStyles(stylesXml(characterStyle('<w:noProof/>')))

    expect(out).toContain('<w:rPr><w:noProof/></w:rPr>')
  })

  it('keeps an unmodeled child on docDefaults and on a table conditional format', () => {
    const out = roundTripStyles(
      stylesXml(
        '<w:docDefaults><w:rPrDefault><w:rPr><w:noProof/></w:rPr></w:rPrDefault></w:docDefaults>' +
          '<w:style w:type="table" w:styleId="T"><w:tblStylePr w:type="firstRow">' +
          '<w:rPr><w:b/><w:snapToGrid/></w:rPr></w:tblStylePr></w:style>',
      ),
    )

    expect(out).toContain('<w:rPrDefault><w:rPr><w:noProof/></w:rPr></w:rPrDefault>')
    expect(out).toContain('<w:b/><w:snapToGrid/>')
  })

  it('carries the root namespace declarations through, so a prefixed fragment stays well-formed', () => {
    // Why this is part of the same fix: the writer used to hardcode `xmlns:w`
    // and `xmlns:r` and nothing else. Re-emitting `<w14:ligatures>` under a root
    // that never declares `w14` is not well-formed XML — Word rejects the whole
    // document as unreadable content, which is worse than the data loss this
    // change set out to fix.
    const out = roundTripStyles(
      stylesXml(
        characterStyle('<w14:ligatures w14:val="standard"/>'),
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
          ` xmlns:w14="${W14_NS}"` +
          ' mc:Ignorable="w14"',
      ),
    )

    expect(out).toContain('<w14:ligatures w14:val="standard"/>')
    expect(out).toContain(`xmlns:w14="${W14_NS}"`)
    expect(out).toContain('mc:Ignorable="w14"')
  })

  it('still emits the default declarations for a StylesPart that was never parsed', () => {
    const out = writeStylesXml({ docDefaults: {}, styles: new Map() })

    expect(out).toContain(`xmlns:w="${W_NS}"`)
    expect(out).toContain('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
  })

  it('leaves no placeholder tag in the output', () => {
    const out = roundTripStyles(stylesXml(characterStyle('<w:noProof/>')))

    expect(out).not.toContain('atlas-raw-')
  })

  it('survives a $ in a passthrough attribute value', () => {
    // `String.replace` with a string replacement reads `$&`/`$1` as substitution
    // syntax; the restore step passes a function for exactly this reason.
    const out = roundTripStyles(stylesXml(characterStyle('<w:custom w:val="a$&amp;b$1c"/>')))

    expect(out).toContain('<w:custom w:val="a$&amp;b$1c"/>')
  })
})

describe('numbering.xml w:rPr passthrough', () => {
  // A `w:lvl`'s `w:rPr` goes through the same `parseRunPropsNode`, so numbering
  // inherited both the capture and the obligation to restore the placeholders.
  // Without the restore step in `writeNumberingXml`, this part would ship an
  // `<atlas-raw-0/>` element to Word.
  function numberingXml(body: string, rootAttributes = ''): string {
    return `${DECL}<w:numbering xmlns:w="${W_NS}"${rootAttributes}>${body}</w:numbering>`
  }

  it('keeps an unmodeled run property on an abstract numbering level', () => {
    const out = writeNumberingXml(
      parseNumbering(
        numberingXml(
          '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">' +
            '<w:rPr><w:b/><w:noProof/></w:rPr></w:lvl></w:abstractNum>',
        ),
      ),
    )

    expect(out).toContain('<w:noProof/>')
    expect(out).not.toContain('atlas-raw-')
  })

  it('keeps an unmodeled run property on a level override', () => {
    const out = writeNumberingXml(
      parseNumbering(
        numberingXml(
          '<w:num w:numId="1"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0">' +
            '<w:lvl w:ilvl="0"><w:rPr><w:snapToGrid/></w:rPr></w:lvl>' +
            '</w:lvlOverride></w:num>',
        ),
      ),
    )

    expect(out).toContain('<w:snapToGrid/>')
    expect(out).not.toContain('atlas-raw-')
  })

  it('carries the root namespace declarations through', () => {
    const out = writeNumberingXml(
      parseNumbering(
        numberingXml(
          '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">' +
            '<w:rPr><w14:ligatures w14:val="none"/></w:rPr></w:lvl></w:abstractNum>',
          ` xmlns:w14="${W14_NS}"`,
        ),
      ),
    )

    expect(out).toContain('<w14:ligatures w14:val="none"/>')
    expect(out).toContain(`xmlns:w14="${W14_NS}"`)
  })
})
