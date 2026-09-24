/**
 * Wires `scripts/lib/officeValidator.mjs` (the spec-level OOXML validation
 * harness) into CI against Atlas's real PPTX editing save path
 * (`office/officePackage.ts`'s `writeOfficePackage`, driven through a
 * realistic sequence of `pptxEdits.ts`/`pptxSlideOps.ts` edits) — checking
 * things `pptxEditing.test.ts` doesn't: `[Content_Types].xml` coverage,
 * dangling `r:id`s, XML-1.0-illegal characters, namespace prefixes
 * declared, and `p:txBody`'s `a:bodyPr`/`a:lstStyle`/`a:p` child order.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { loadOfficePackage, writeOfficePackage, type OfficePackage } from '../../../../../office/officePackage'
import { addSlide, deleteSlide, duplicateSlide, moveSlide } from '../pptxSlideOps'
import { deleteShape, insertTextBox, setShapeBox, setShapeText, setSlideNotes } from '../pptxEdits'
import { buildEditableDeck } from './editableDeck'

import { validateOfficeFile } from '../../../../../../scripts/lib/officeValidator.mjs'

const NS = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`

/** SHELL-5 — a package with no slideLayout/slideMaster/theme parts at all. */
async function buildLayoutlessDeck(): Promise<ArrayBuffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`)
  zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation ${NS}>
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>`)
  zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`)
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
  <p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>Only slide</a:t></a:r></a:p></p:txBody></p:sp>
</p:spTree></p:cSld></p:sld>`)
  return zip.generateAsync({ type: 'arraybuffer' })
}

const SLIDE_1 = 'ppt/slides/slide1.xml'
const SLIDE_2 = 'ppt/slides/slide2.xml'

function errorsOnly(issues: ReadonlyArray<{ readonly severity: string }>) {
  return issues.filter((i) => i.severity === 'error')
}

async function validate(pkg: OfficePackage) {
  const bytes = await writeOfficePackage(pkg)
  return validateOfficeFile(Buffer.from(bytes))
}

describe('PPTX editing saves pass spec-level OOXML validation', () => {
  it('a realistic sequence of edits produces a spec-clean package', async () => {
    let pkg = await loadOfficePackage(await buildEditableDeck())

    pkg = setShapeText(pkg, SLIDE_1, '2', 'Annual review\nSecond line')
    pkg = setShapeBox(pkg, SLIDE_1, '2', { x: 10, y: 20, w: 300, h: 40 })
    const inserted = insertTextBox(pkg, SLIDE_2, { x: 100, y: 100, w: 400, h: 50 }, 'New <text> & "quoted" more')
    pkg = deleteShape(inserted.pkg, SLIDE_2, inserted.sourceId!)
    pkg = setSlideNotes(pkg, SLIDE_1, 'Thank the team')
    pkg = addSlide(pkg, 0).pkg
    pkg = duplicateSlide(pkg, 1)
    pkg = moveSlide(pkg, 0, 2)
    pkg = deleteSlide(pkg, 0)

    const result = await validate(pkg)
    expect(result.format).toEqual({ family: 'opc', kind: 'pptx' })
    expect(errorsOnly(result.issues)).toEqual([])
  })

  it('an added slide alone is spec-clean', async () => {
    const pkg = await loadOfficePackage(await buildEditableDeck())
    const { pkg: withNewSlide } = addSlide(pkg, 0)
    const result = await validate(withNewSlide)
    expect(errorsOnly(result.issues)).toEqual([])
  })

  it('a duplicated + moved + deleted slide sequence is spec-clean', async () => {
    let pkg = await loadOfficePackage(await buildEditableDeck())
    pkg = duplicateSlide(pkg, 0)
    pkg = moveSlide(pkg, 2, 0)
    pkg = deleteSlide(pkg, 1)
    const result = await validate(pkg)
    expect(errorsOnly(result.issues)).toEqual([])
  })

  // SHELL-5 — the synthesized fallback slideMaster/slideLayout/theme (added
  // when a package has no slideLayout at all, so "New Slide" isn't a silent
  // no-op) must itself be a spec-clean save, not just "good enough for
  // Atlas's own lenient parser".
  it('a slide added via the synthesized fallback layout is spec-clean', async () => {
    const pkg = await loadOfficePackage(await buildLayoutlessDeck())
    const { pkg: withNewSlide } = addSlide(pkg, 0)
    const result = await validate(withNewSlide)
    expect(errorsOnly(result.issues)).toEqual([])
  })
})
