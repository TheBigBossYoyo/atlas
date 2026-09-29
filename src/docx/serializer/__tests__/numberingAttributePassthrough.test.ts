/**
 * FID-NUM-1 — `numbering.xml` round-trips the attributes Atlas does not model.
 *
 * Two silent, unconditional losses on every save of any document with a list:
 *
 *   - `w:abstractNum/@w15:restartNumberingAfterBreak` was modelled nowhere, so
 *     every `w:abstractNum` that has one lost it. Every fixture in this repo's
 *     corpus has one.
 *   - `w:lvl/@w15:tentative` was read as `getAttr(node, 'w:tentative')` — the
 *     wrong namespace. Real files write `w15:tentative` (nine times per file
 *     across this corpus; not one writes `w:tentative`), so the lookup never
 *     matched and the value went every time, under either prefix.
 *
 * Both were found by running the extended fidelity detector corpus-wide, not by
 * reading the code — the `w:tentative` lookup looks perfectly correct in
 * isolation, and is why a warning detector earns its keep over inspection.
 *
 * These assert the serialized XML, because the model holding an attribute says
 * nothing about the file getting it back.
 */
import { describe, expect, it } from 'vitest'

import { parseNumbering } from '../../parser/numbering'
import { writeNumberingXml } from '../numberingWriter'

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml'

/** Shaped exactly like Word's own output — see this file's header. */
function numberingXml(body: string, rootAttributes = ` xmlns:w15="${W15_NS}"`): string {
  return `${DECL}<w:numbering xmlns:w="${W_NS}"${rootAttributes}>${body}</w:numbering>`
}

function roundTrip(xml: string): string {
  return writeNumberingXml(parseNumbering(xml))
}

describe('numbering.xml attribute passthrough', () => {
  it('keeps w15:restartNumberingAfterBreak on an abstract numbering definition', () => {
    const out = roundTrip(
      numberingXml(
        '<w:abstractNum w:abstractNumId="1" w15:restartNumberingAfterBreak="0">' +
          '<w:multiLevelType w:val="hybridMultilevel"/></w:abstractNum>',
      ),
    )

    expect(out).toContain('w15:restartNumberingAfterBreak="0"')
  })

  it('keeps w15:tentative on a level', () => {
    const out = roundTrip(
      numberingXml(
        '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0" w15:tentative="1">' +
          '<w:numFmt w:val="bullet"/></w:lvl></w:abstractNum>',
      ),
    )

    expect(out).toContain('w15:tentative="1"')
  })

  it('still writes a spec-conformant w:tentative through its own modelled field', () => {
    // `LvlDef.tentative` remains the `w`-namespace field it always was, so a file
    // that follows the base schema rather than Word's habit keeps working. The two
    // are deliberately separate: folding them into one field would mean guessing
    // which namespace to write back.
    const out = roundTrip(
      numberingXml('<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0" w:tentative="1"/></w:abstractNum>'),
    )

    expect(out).toContain('w:tentative="1"')
    expect(out).not.toContain('w15:tentative')
  })

  it('keeps an attribute it has never heard of, on either node', () => {
    // The point of a map rather than a field per attribute: the next `w16:`
    // extension attribute must not be lost the same way these two were.
    const out = roundTrip(
      numberingXml(
        '<w:abstractNum w:abstractNumId="1" w16:invented="a">' +
          '<w:lvl w:ilvl="0" w16:alsoInvented="b"/></w:abstractNum>',
        ` xmlns:w15="${W15_NS}" xmlns:w16="http://example.invalid/w16"`,
      ),
    )

    expect(out).toContain('w16:invented="a"')
    expect(out).toContain('w16:alsoInvented="b"')
  })

  it('keeps the attributes on a level nested inside a level override', () => {
    const out = roundTrip(
      numberingXml(
        '<w:num w:numId="1"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0">' +
          '<w:lvl w:ilvl="0" w15:tentative="1"/></w:lvlOverride></w:num>',
      ),
    )

    expect(out).toContain('w15:tentative="1"')
  })

  it('adds no attributes to a node that had none unusual', () => {
    const out = roundTrip(
      numberingXml('<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum>'),
    )

    expect(out).toContain('<w:abstractNum w:abstractNumId="1">')
    expect(out).toContain('<w:lvl w:ilvl="0">')
  })

  it('does not treat the modelled attributes as unknown and write them twice', () => {
    const out = roundTrip(
      numberingXml('<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0" w:tentative="1"/></w:abstractNum>'),
    )

    expect(out.match(/w:abstractNumId="1"/g)).toHaveLength(1)
    expect(out.match(/w:ilvl="0"/g)).toHaveLength(1)
    expect(out.match(/w:tentative="1"/g)).toHaveLength(1)
  })
})
