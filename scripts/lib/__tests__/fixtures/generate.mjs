#!/usr/bin/env node
// Atlas -- one-off generator for `scripts/lib/__tests__/fixtures/*.docx`,
// the fixture corpus `officeValidatorDatatypes.test.mjs` reads from disk.
//
// These are minimal, hand-built .docx packages (a small handful of parts,
// nothing Atlas's own writers touch) -- each pair isolates exactly one
// ST_* attribute-datatype rule from `scripts/lib/officeValidator.mjs`'s
// `ELEMENT_ATTRIBUTE_TYPES` table: one file with a value that violates the
// rule, one with a value that satisfies it. Regenerate after editing this
// file with:
//   node scripts/lib/__tests__/fixtures/generate.mjs
//
// Not run as part of the test suite (it writes files) and not part of CI --
// the committed .docx fixtures are what the tests actually read.
import { writeFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

import JSZip from 'jszip'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

const CONTENT_TYPES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${CT_NS}">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
  `</Types>`

const PACKAGE_RELS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
  `</Relationships>`

/**
 * Wraps a `<w:body>` fragment in a minimal-but-complete document.xml, and
 * that in a minimal-but-complete .docx package.
 * @param {string} bodyXml
 * @returns {Promise<Buffer>}
 */
async function buildDocx(bodyXml) {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="${W_NS}"><w:body>${bodyXml}<w:sectPr/></w:body></w:document>`

  const zip = new JSZip()
  zip.file('[Content_Types].xml', CONTENT_TYPES_XML)
  zip.file('_rels/.rels', PACKAGE_RELS_XML)
  zip.file('word/document.xml', documentXml)
  return zip.generateAsync({ type: 'nodebuffer' })
}

const FIXTURES = {
  // ST_HexColor -- the exact shape from the D-VALID sweep: a leading "#" is
  // never legal (ST_HexColorRGB is xsd:hexBinary, not a CSS-style color).
  'hex-color-invalid.docx': `<w:p><w:pPr><w:rPr><w:color w:val="#ff0000"/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val="#ff0000"/></w:rPr><w:t>x</w:t></w:r></w:p>`,
  'hex-color-valid.docx': `<w:p><w:pPr><w:rPr><w:color w:val="FF0000"/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val="auto"/></w:rPr><w:t>x</w:t></w:r></w:p>`,

  // ST_DecimalNumber -- the other D-VALID shape: a label instead of an
  // integer. Exercised on w:ilvl/w:numId (val-elements inside w:numPr),
  // the pair actually reachable from a paragraph body.
  'decimal-number-invalid.docx': `<w:p><w:pPr><w:numPr><w:ilvl w:val="not-a-number"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
  'decimal-number-valid.docx': `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,

  // ST_OnOff -- an illegal token on a CT_OnOff run toggle, and the same
  // toggle using two different legal spellings (proving the checker accepts
  // the whole legal set, not just "true"/"false").
  'onoff-invalid.docx': `<w:p><w:r><w:rPr><w:b w:val="yes"/></w:rPr><w:t>x</w:t></w:r></w:p>`,
  'onoff-valid.docx': `<w:p><w:r><w:rPr><w:b w:val="1"/><w:i w:val="on"/></w:rPr><w:t>x</w:t></w:r></w:p>`,

  // ST_TwipsMeasure (unsigned) -- w:spacing's before/after never go
  // negative; a unit-suffixed positive value is legal.
  'twips-measure-invalid.docx': `<w:p><w:pPr><w:spacing w:before="-240" w:after="240"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
  'twips-measure-valid.docx': `<w:p><w:pPr><w:spacing w:before="240" w:after="0.5in"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,

  // ST_SignedTwipsMeasure -- the counterpart proving negative IS legal here
  // (w:ind's left/right), so a validator that reused the unsigned rule
  // everywhere would false-fail this file; only a non-numeric, non-measure
  // value is actually invalid.
  'signed-twips-measure-invalid.docx': `<w:p><w:pPr><w:ind w:left="wide" w:right="-360"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
  'signed-twips-measure-valid.docx': `<w:p><w:pPr><w:ind w:left="-360" w:right="360"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
}

for (const [filename, bodyXml] of Object.entries(FIXTURES)) {
  const buffer = await buildDocx(bodyXml)
  const outPath = path.join(__dirname, filename)
  await writeFile(outPath, buffer)
  console.log(`wrote ${outPath} (${buffer.length} bytes)`)
}
