/**
 * Wires `scripts/lib/officeValidator.mjs` (the spec-level OOXML validation
 * harness) into CI against Atlas's real PPTX editing save path
 * (`office/officePackage.ts`'s `writeOfficePackage`, driven through a
 * realistic sequence of `pptxEdits.ts`/`pptxSlideOps.ts` edits) — checking
 * things `pptxEditing.test.ts` doesn't: `[Content_Types].xml` coverage,
 * dangling `r:id`s, XML-1.0-illegal characters, namespace prefixes
 * declared, and `p:txBody`'s `a:bodyPr`/`a:lstStyle`/`a:p` child order.
 */
import { describe, expect, it } from 'vitest'

import { loadOfficePackage, writeOfficePackage, type OfficePackage } from '../../../../../office/officePackage'
import { addSlide, deleteSlide, duplicateSlide, moveSlide } from '../pptxSlideOps'
import { deleteShape, insertTextBox, setShapeBox, setShapeText, setSlideNotes } from '../pptxEdits'
import { buildEditableDeck } from './editableDeck'

import { validateOfficeFile } from '../../../../../../scripts/lib/officeValidator.mjs'

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
})
