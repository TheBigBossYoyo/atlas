/**
 * FID-DRAW-1 — a saved drawing keeps the attributes Atlas does not model.
 *
 * Three silent, unconditional losses, all from the same cause: the drawing writer
 * rebuilt each element's attributes from a handful of modelled fields plus
 * literals, so anything else the source carried simply ceased to exist.
 *
 *   - `wp:inline` was given NO attributes at all, so an inline image's
 *     `distT`/`distB`/`distL`/`distR` — the space around it — went on every save.
 *   - `wp:anchor`'s `distT`/`distB`/`distL`/`distR`/`simplePos`/`relativeHeight`/
 *     `locked`/`layoutInCell` were HARDCODED literals, read from nothing.
 *     `relativeHeight` is the floating image's Z-ORDER, so two anchored images
 *     stacked in a particular order both came back as `relativeHeight="0"` and
 *     which one sat in front was silently flattened.
 *   - `a:blip/@cstate` and `pic:spPr/@bwMode` were never emitted.
 *
 * Found by the corpus fidelity detector rather than by reading the code, and
 * invisible to the LibreOffice text comparison in
 * `src/docx/__tests__/libreOfficeRoundTrip.test.ts` — none of this changes a
 * single character of text, which is exactly why it needed its own detector.
 *
 * These assert the serialized XML: the model holding an attribute says nothing
 * about the file getting it back.
 */
import { describe, expect, it } from 'vitest'

import { parseDocument } from '../../parser/document'
import { writeDocumentXml } from '../documentWriter'

const BASE_NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
].join(' ')

function documentXml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${BASE_NAMESPACES}><w:body>${body}</w:body></w:document>`
}

function roundTrip(body: string): string {
  return writeDocumentXml(parseDocument(documentXml(body)))
}

/** A picture subtree, with whatever attributes the caller wants on blip/spPr. */
function picture(blipAttrs = '', spPrAttrs = ''): string {
  return (
    '<a:graphic><a:graphicData><pic:pic>' +
    `<pic:blipFill><a:blip r:embed="rId9"${blipAttrs}/></pic:blipFill>` +
    `<pic:spPr${spPrAttrs}><a:xfrm><a:off x="0" y="0"/></a:xfrm></pic:spPr>` +
    '</pic:pic></a:graphicData></a:graphic>'
  )
}

describe('drawing attribute passthrough', () => {
  it('keeps an anchored image z-order instead of flattening it to 0', () => {
    // The whole point of the fix. Two images at different heights must not both
    // come back at 0, or their front-to-back order is gone.
    const out = roundTrip(
      '<w:p><w:r><w:drawing>' +
        '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="3"' +
        ' behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
        '<wp:extent cx="914400" cy="457200"/><wp:docPr id="1"/>' +
        picture() +
        '</wp:anchor></w:drawing></w:r>' +
        '<w:r><w:drawing>' +
        '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="7"' +
        ' behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
        '<wp:extent cx="914400" cy="457200"/><wp:docPr id="2"/>' +
        picture() +
        '</wp:anchor></w:drawing></w:r></w:p>',
    )

    expect(out).toContain('relativeHeight="3"')
    expect(out).toContain('relativeHeight="7"')
  })

  it('keeps the space around an anchored image', () => {
    const out = roundTrip(
      '<w:p><w:r><w:drawing>' +
        '<wp:anchor distT="114300" distB="228600" distL="342900" distR="457200" simplePos="0"' +
        ' relativeHeight="0" behindDoc="0" locked="1" layoutInCell="0" allowOverlap="1">' +
        '<wp:extent cx="914400" cy="457200"/><wp:docPr id="1"/>' +
        picture() +
        '</wp:anchor></w:drawing></w:r></w:p>',
    )

    expect(out).toContain('distT="114300"')
    expect(out).toContain('distB="228600"')
    expect(out).toContain('distL="342900"')
    expect(out).toContain('distR="457200"')
    expect(out).toContain('locked="1"')
    expect(out).toContain('layoutInCell="0"')
  })

  it('keeps the space around an INLINE image, which used to get no attributes at all', () => {
    const out = roundTrip(
      '<w:p><w:r><w:drawing>' +
        '<wp:inline distT="114300" distB="114300" distL="228600" distR="228600">' +
        '<wp:extent cx="914400" cy="457200"/><wp:docPr id="1"/>' +
        picture() +
        '</wp:inline></w:drawing></w:r></w:p>',
    )

    expect(out).toContain('distT="114300"')
    expect(out).toContain('distL="228600"')
  })

  it('does not invent dist attributes on an inline image that never had them', () => {
    // `wp:inline` has no required attributes, unlike `wp:anchor`. A save must not
    // add `dist*="0"` to a file that did not carry them.
    const out = roundTrip(
      '<w:p><w:r><w:drawing><wp:inline>' +
        '<wp:extent cx="914400" cy="457200"/><wp:docPr id="1"/>' +
        picture() +
        '</wp:inline></w:drawing></w:r></w:p>',
    )

    expect(out).toContain('<wp:inline>')
    expect(out).not.toContain('distT')
  })

  it('keeps a:blip/@cstate and pic:spPr/@bwMode', () => {
    const out = roundTrip(
      '<w:p><w:r><w:drawing><wp:inline>' +
        '<wp:extent cx="914400" cy="457200"/><wp:docPr id="1"/>' +
        picture(' cstate="print"', ' bwMode="auto"') +
        '</wp:inline></w:drawing></w:r></w:p>',
    )

    expect(out).toContain('cstate="print"')
    expect(out).toContain('bwMode="auto"')
  })

  it('still writes the attributes CT_Anchor requires when the source omitted them', () => {
    // A `Drawing` can also be built by code rather than parsed, and an anchor
    // without these is not schema-valid, so the literals remain as a fallback.
    const out = roundTrip(
      '<w:p><w:r><w:drawing><wp:anchor>' +
        '<wp:extent cx="914400" cy="457200"/><wp:docPr id="1"/>' +
        picture() +
        '</wp:anchor></w:drawing></w:r></w:p>',
    )

    for (const required of ['distT', 'distB', 'distL', 'distR', 'simplePos', 'relativeHeight', 'behindDoc', 'locked', 'layoutInCell', 'allowOverlap']) {
      expect(out, `wp:anchor must still carry ${required}`).toContain(`${required}=`)
    }
  })

  it('lets the modelled behindDoc/allowOverlap win over a captured copy', () => {
    // Those two have typed fields the layout engine reads, so they are applied
    // after the captured map — otherwise a stale copy could contradict the model
    // the renderer is using.
    const out = roundTrip(
      '<w:p><w:r><w:drawing>' +
        '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0"' +
        ' behindDoc="1" locked="0" layoutInCell="1" allowOverlap="0">' +
        '<wp:extent cx="914400" cy="457200"/><wp:docPr id="1"/>' +
        picture() +
        '</wp:anchor></w:drawing></w:r></w:p>',
    )

    expect(out.match(/behindDoc="1"/g)).toHaveLength(1)
    expect(out.match(/allowOverlap="0"/g)).toHaveLength(1)
  })
})
