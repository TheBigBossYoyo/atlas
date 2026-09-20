// Atlas -- tests for the attribute-datatype checks
// `scripts/lib/officeValidator.mjs` added on top of its existing structural
// checks (see that file's "Attribute datatype checks" section). Each pair
// of fixtures under `./fixtures/` isolates exactly one ST_* simple type:
// a file whose value violates it, and a file whose value satisfies it (see
// `./fixtures/generate.mjs` for how they were built and to regenerate).
//
// Run with:
//   npx vitest run --config scripts/lib/__tests__/vitest.config.mjs --maxWorkers=1 scripts/lib/__tests__/officeValidatorDatatypes.test.mjs
import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { validateOfficeFile } from '../officeValidator.mjs'

/** @typedef {import('../officeValidator.d.mts').OfficeValidationIssue} Issue */

const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

/**
 * Builds a minimal one-part .docx around a `<w:body>` fragment, in memory (no fixture file needed).
 * @param {string} bodyXml
 * @returns {Promise<Buffer>}
 */
async function buildDocx(bodyXml) {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${CT_NS}">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `</Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
      `</Relationships>`,
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W_NS}"><w:body>${bodyXml}<w:sectPr/></w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const FIXTURES_DIR = path.join(__dirname, 'fixtures')

/** @param {string} filename */
async function loadFixture(filename) {
  const buffer = await readFile(path.join(FIXTURES_DIR, filename))
  return validateOfficeFile(buffer)
}

/** @param {ReadonlyArray<Issue>} issues */
function errorsOnly(issues) {
  return issues.filter((i) => i.severity === 'error')
}

describe('officeValidator -- ST_HexColor', () => {
  it('flags a leading "#" (the D-VALID sweep bug: <w:color w:val="#ff0000"/>)', async () => {
    const result = await loadFixture('hex-color-invalid.docx')
    const bad = errorsOnly(result.issues).filter((i) => i.code === 'invalid-attribute-value')
    expect(bad.length).toBeGreaterThan(0)
    for (const i of bad) {
      expect(i.message).toContain('w:color')
      expect(i.message).toContain('w:val')
      expect(i.message).toContain('#ff0000')
      expect(i.message).toContain('ST_HexColor')
    }
  })

  it('accepts "auto" and six bare hex digits', async () => {
    const result = await loadFixture('hex-color-valid.docx')
    expect(errorsOnly(result.issues)).toEqual([])
  })
})

describe('officeValidator -- ST_DecimalNumber', () => {
  it('flags a non-integer numbering value (the D-VALID sweep bug shape: a label instead of an id)', async () => {
    const result = await loadFixture('decimal-number-invalid.docx')
    const bad = errorsOnly(result.issues).filter((i) => i.code === 'invalid-attribute-value')
    expect(bad.length).toBe(1)
    expect(bad[0].message).toContain('w:ilvl')
    expect(bad[0].message).toContain('w:val')
    expect(bad[0].message).toContain('not-a-number')
    expect(bad[0].message).toContain('ST_DecimalNumber')
  })

  it('accepts a bare integer', async () => {
    const result = await loadFixture('decimal-number-valid.docx')
    expect(errorsOnly(result.issues)).toEqual([])
  })
})

describe('officeValidator -- ST_OnOff', () => {
  it('flags a value outside true/false/1/0/on/off', async () => {
    const result = await loadFixture('onoff-invalid.docx')
    const bad = errorsOnly(result.issues).filter((i) => i.code === 'invalid-attribute-value')
    expect(bad.length).toBe(1)
    expect(bad[0].message).toContain('w:b')
    expect(bad[0].message).toContain('w:val')
    expect(bad[0].message).toContain('yes')
    expect(bad[0].message).toContain('ST_OnOff')
  })

  it('accepts every spelling in the legal set, not just true/false', async () => {
    const result = await loadFixture('onoff-valid.docx')
    expect(errorsOnly(result.issues)).toEqual([])
  })
})

describe('officeValidator -- ST_TwipsMeasure (unsigned)', () => {
  it('flags a negative value on an attribute that must never be signed', async () => {
    const result = await loadFixture('twips-measure-invalid.docx')
    const bad = errorsOnly(result.issues).filter((i) => i.code === 'invalid-attribute-value')
    expect(bad.length).toBe(1)
    expect(bad[0].message).toContain('w:spacing')
    expect(bad[0].message).toContain('w:before')
    expect(bad[0].message).toContain('-240')
    expect(bad[0].message).toContain('ST_TwipsMeasure')
  })

  it('accepts a bare non-negative integer and a positive unit-suffixed measurement', async () => {
    const result = await loadFixture('twips-measure-valid.docx')
    expect(errorsOnly(result.issues)).toEqual([])
  })
})

describe('officeValidator -- ST_SignedTwipsMeasure', () => {
  it('flags a value that is neither an integer nor a unit-suffixed measurement', async () => {
    const result = await loadFixture('signed-twips-measure-invalid.docx')
    const bad = errorsOnly(result.issues).filter((i) => i.code === 'invalid-attribute-value')
    expect(bad.length).toBe(1)
    expect(bad[0].message).toContain('w:ind')
    expect(bad[0].message).toContain('w:left')
    expect(bad[0].message).toContain('wide')
    expect(bad[0].message).toContain('ST_SignedTwipsMeasure')
  })

  it('does NOT flag a negative value -- unlike ST_TwipsMeasure, negative is legal here', async () => {
    const result = await loadFixture('signed-twips-measure-valid.docx')
    expect(errorsOnly(result.issues)).toEqual([])
  })
})

describe('officeValidator -- attribute datatype checks stay silent when the attribute is absent', () => {
  it('does not flag a CT_OnOff element with no w:val at all (absence means "on")', async () => {
    const buffer = await buildDocx('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>x</w:t></w:r></w:p>')
    const result = validateOfficeFile(buffer)
    expect(errorsOnly(result.issues).filter((i) => i.code === 'invalid-attribute-value')).toEqual([])
  })

  it('does not flag an ST_SignedTwipsMeasure attribute that is simply not present', async () => {
    const buffer = await buildDocx('<w:p><w:pPr><w:ind w:firstLine="120"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>')
    const result = validateOfficeFile(buffer)
    expect(errorsOnly(result.issues).filter((i) => i.code === 'invalid-attribute-value')).toEqual([])
  })
})
