/**
 * Wires `scripts/lib/officeValidator.mjs` (the spec-level ODF validation
 * harness) into CI against Atlas's real ODP editing save path
 * (`office/officePackage.ts`'s `writeOfficePackage`, driven through a
 * realistic sequence of `odpEdits.ts` edits) — checking things
 * `odpEditing.test.ts` doesn't: `mimetype` first-and-stored,
 * `META-INF/manifest.xml` listing every file, and every part's XML being
 * well-formed with no illegal characters or undeclared namespace prefixes.
 *
 * `odpFixture.ts` (the shared parser fixture this borrows its content from)
 * has no `META-INF/manifest.xml` at all — harmless for the PARSER tests
 * that use it (nothing there reads the manifest), but not a valid ODF
 * package on its own, so a manifest is added here to make it one before
 * exercising the actual save path.
 */
import { describe, expect, it } from 'vitest'

import { loadOfficePackage, writeOfficePackage, type OfficePackage } from '../../../../../office/officePackage'
import { buildOdpFixtureZip } from '../../__tests__/odpFixture'
import { addSlide, deleteShape, deleteSlide, duplicateSlide, insertTextBox, moveSlide, setShapeBox, setShapeText, setSlideNotes } from '../odpEdits'

import { validateOfficeFile } from '../../../../../../scripts/lib/officeValidator.mjs'

const MANIFEST_XML = `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.presentation"/>
<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
<manifest:file-entry manifest:full-path="Pictures/image1.png" manifest:media-type="image/png"/>
</manifest:manifest>`

async function buildCompletePackage(): Promise<OfficePackage> {
  const zip = buildOdpFixtureZip()
  zip.file('META-INF/manifest.xml', MANIFEST_XML)
  const buffer = await zip.generateAsync({ type: 'arraybuffer' })
  return loadOfficePackage(buffer)
}

function errorsOnly(issues: ReadonlyArray<{ readonly severity: string }>) {
  return issues.filter((i) => i.severity === 'error')
}

async function validate(pkg: OfficePackage) {
  const bytes = await writeOfficePackage(pkg)
  return validateOfficeFile(Buffer.from(bytes))
}

describe('ODP editing saves pass spec-level ODF validation', () => {
  it('a realistic sequence of edits produces a spec-clean package', async () => {
    let pkg = await buildCompletePackage()

    pkg = setShapeText(pkg, 0, '0', 'Rewritten title\nSecond line')
    pkg = setShapeBox(pkg, 0, '0', { x: 100, y: 50, w: 300, h: 120 })
    const inserted = insertTextBox(pkg, 0, { x: 10, y: 10, w: 200, h: 50 }, 'Added <text> & "quoted"')
    pkg = deleteShape(inserted.pkg, 0, inserted.sourceId!)
    pkg = setSlideNotes(pkg, 0, 'Remember the EMEA numbers')
    pkg = addSlide(pkg, 0).pkg
    pkg = duplicateSlide(pkg, 1)
    pkg = moveSlide(pkg, 0, 2)
    pkg = deleteSlide(pkg, 0)

    const result = await validate(pkg)
    expect(result.format).toMatchObject({ family: 'odf', kind: 'odp' })
    expect(errorsOnly(result.issues)).toEqual([])
  })

  it('an added slide alone is spec-clean', async () => {
    const pkg = await buildCompletePackage()
    const { pkg: withNewSlide } = addSlide(pkg, 0)
    const result = await validate(withNewSlide)
    expect(errorsOnly(result.issues)).toEqual([])
  })

  it('mimetype stays first and stored through a save with edits', async () => {
    let pkg = await buildCompletePackage()
    pkg = setShapeText(pkg, 0, '0', 'Edited')
    const result = await validate(pkg)
    expect(result.issues.filter((i) => i.code.startsWith('odf-mimetype'))).toEqual([])
  })
})
