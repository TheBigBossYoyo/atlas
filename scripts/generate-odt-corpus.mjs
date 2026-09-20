#!/usr/bin/env node
// Deterministic generator for the ODT fixture corpus (T8/DAT-19 — see
// .sisyphus/plans/atlas-phase3-improvement.md and
// .sisyphus/plans/atlas-phase3-findings-register.md).
//
// odf-kit is a pre-1.0 dependency with no local fixtures pinned against it
// (DAT-19: "ODT fidelity rests on a pre-1.0 dependency with no local
// fixtures"). This script produces a small, focused `.odt` per real-world
// ODF feature into src/viewers/__fixtures__/odt-corpus/, consumed by
// src/viewers/__tests__/OdtViewer.corpus.test.ts, which snapshots each
// fixture's rendered HTML — so a future odf-kit minor bump that silently
// changes output has something concrete to diff against instead of going
// unnoticed.
//
// Run with: node scripts/generate-odt-corpus.mjs
import { mkdir, writeFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

import JSZip from 'jszip'

import { createSolidPng } from './lib/corpusPng.mjs'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const projectRoot = path.join(__dirname, '..')
const outputDir = path.join(projectRoot, 'src', 'viewers', '__fixtures__', 'odt-corpus')

const MANIFEST_XML = `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
`

// US Letter (matches PageEstimate's DEFAULT_PAGE_HEIGHT_PX fallback), so a
// fixture-corpus assertion and the no-page-layout fallback agree by
// construction rather than by coincidence.
const STYLES_XML = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  office:version="1.2">
  <office:styles/>
  <office:automatic-styles>
    <style:page-layout style:name="pm1">
      <style:page-layout-properties fo:page-width="21.59cm" fo:page-height="27.94cm" fo:margin-top="2.54cm" fo:margin-bottom="2.54cm" fo:margin-left="2.54cm" fo:margin-right="2.54cm"/>
    </style:page-layout>
  </office:automatic-styles>
  <office:master-styles>
    <style:master-page style:name="Standard" style:page-layout-name="pm1"/>
  </office:master-styles>
</office:document-styles>
`

const CONTENT_XML_NS = `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"
  office:version="1.2"`

/**
 * @param {string} automaticStyles
 * @param {string} bodyXml
 * @returns {string}
 */
function contentXml(automaticStyles, bodyXml) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${CONTENT_XML_NS}>
  <office:automatic-styles>
${automaticStyles}
  </office:automatic-styles>
  <office:body>
    <office:text>
${bodyXml}
    </office:text>
  </office:body>
</office:document-content>
`
}

/**
 * @param {string} name
 * @param {{ automaticStyles?: string, body: string }} fixture
 * @returns {Promise<string>}
 */
async function writeOdtFixture(name, { automaticStyles = '', body }) {
  const zip = new JSZip()
  zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { compression: 'STORE' })
  zip.file('META-INF/manifest.xml', MANIFEST_XML)
  zip.file('styles.xml', STYLES_XML)
  zip.file('content.xml', contentXml(automaticStyles, body))
  const buffer = await zip.generateAsync({ type: 'nodebuffer' })
  const targetPath = path.join(outputDir, name)
  await writeFile(targetPath, buffer)
  return targetPath
}

/** A table (header row + two data rows) and one bulleted + one numbered list. */
function tablesAndListsFixture() {
  const automaticStyles = `
    <text:list-style style:name="L1">
      <text:list-level-style-bullet text:level="1" text:bullet-char="&#8226;"/>
    </text:list-style>
    <text:list-style style:name="L2">
      <text:list-level-style-number text:level="1" style:num-format="1"/>
    </text:list-style>
  `
  const body = `
      <text:p>A table with a header row and two data rows:</text:p>
      <table:table table:name="Table1">
        <table:table-column table:number-columns-repeated="2"/>
        <table:table-row>
          <table:table-cell office:value-type="string"><text:p>Format</text:p></table:table-cell>
          <table:table-cell office:value-type="string"><text:p>Status</text:p></table:table-cell>
        </table:table-row>
        <table:table-row>
          <table:table-cell office:value-type="string"><text:p>ODT</text:p></table:table-cell>
          <table:table-cell office:value-type="string"><text:p>Supported</text:p></table:table-cell>
        </table:table-row>
        <table:table-row>
          <table:table-cell office:value-type="string"><text:p>RTF</text:p></table:table-cell>
          <table:table-cell office:value-type="string"><text:p>Supported</text:p></table:table-cell>
        </table:table-row>
      </table:table>
      <text:p>An unordered list:</text:p>
      <text:list text:style-name="L1">
        <text:list-item><text:p>First bullet</text:p></text:list-item>
        <text:list-item><text:p>Second bullet</text:p></text:list-item>
      </text:list>
      <text:p>An ordered list:</text:p>
      <text:list text:style-name="L2">
        <text:list-item><text:p>First step</text:p></text:list-item>
        <text:list-item><text:p>Second step</text:p></text:list-item>
      </text:list>
  `
  return { automaticStyles, body }
}

/**
 * An inline base64-embedded image (ODF's `office:binary-data` form — no
 * separate Pictures/ entry needed). `loext:mime-type` (a LibreOffice
 * extension attribute genuinely present on real LibreOffice-authored
 * inline images) is what odf-kit's reader falls back to for the MIME type
 * when there's no `xlink:href` into a manifest-registered part — without
 * it, the parsed ImageNode has no `mediaType` and the HTML renderer omits
 * `src` entirely (see html-renderer.js's `renderImage`).
 * @param {string} pngBase64
 * @returns {{ automaticStyles: string, body: string }}
 */
function embeddedImageFixture(pngBase64) {
  const body = `
      <text:p>An embedded image follows:</text:p>
      <text:p>
        <draw:frame draw:name="Image1" svg:width="2cm" svg:height="2cm">
          <draw:image loext:mime-type="image/png">
            <office:binary-data>${pngBase64}</office:binary-data>
          </draw:image>
        </draw:frame>
      </text:p>
      <text:p>Caption text after the image.</text:p>
  `
  return { automaticStyles: '', body }
}

/**
 * A block-level inserted paragraph and a block-level deleted paragraph,
 * structured per ODF §5.5 (text:tracked-changes registry + text:change-start
 * / text:change-end / text:change body markers) — see the reader's
 * parser.js `parseBodyNodes` for exactly how these are consumed. Author/date
 * are direct children of text:insertion/text:deletion (not wrapped in
 * office:change-info) to match what odf-kit 0.13.4's `parseChangedRegions`
 * actually reads today; if a future version adds office:change-info
 * support, this fixture's snapshot will need updating alongside it.
 */
function trackedChangesFixture() {
  const body = `
      <text:tracked-changes>
        <text:changed-region text:id="ct1">
          <text:insertion>
            <dc:creator>Atlas QA</dc:creator>
            <dc:date>2026-01-01T00:00:00</dc:date>
          </text:insertion>
        </text:changed-region>
        <text:changed-region text:id="ct2">
          <text:deletion>
            <dc:creator>Atlas QA</dc:creator>
            <dc:date>2026-01-01T00:00:00</dc:date>
            <text:p>This paragraph was deleted.</text:p>
          </text:deletion>
        </text:changed-region>
      </text:tracked-changes>
      <text:p>An unchanged intro paragraph.</text:p>
      <text:change-start text:change-id="ct1"/>
      <text:p>This paragraph was inserted by a reviewer.</text:p>
      <text:change-end text:change-id="ct1"/>
      <text:change text:change-id="ct2"/>
      <text:p>An unchanged closing paragraph.</text:p>
  `
  return { automaticStyles: '', body }
}

export async function generateOdtCorpus() {
  await mkdir(outputDir, { recursive: true })
  const pngBase64 = createSolidPng(4, 4, [37, 99, 235]).toString('base64')

  const writes = [
    writeOdtFixture('tables-and-lists.odt', tablesAndListsFixture()),
    writeOdtFixture('embedded-image.odt', embeddedImageFixture(pngBase64)),
    writeOdtFixture('tracked-changes.odt', trackedChangesFixture()),
  ]
  return Promise.all(writes)
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const files = await generateOdtCorpus()
  for (const file of files) {
    console.log(path.relative(projectRoot, file).replace(/\\/g, '/'))
  }
}
