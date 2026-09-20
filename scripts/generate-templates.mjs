#!/usr/bin/env node
// NEW-01 — regenerates Atlas's blank "New Document" templates, one per
// editable+saveable format (docx/xlsx/ods/pptx/odp — see
// `electron/lib/newDocumentTemplates.cjs`'s header for why markdown/csv/tsv
// need no template file at all). Run `node scripts/generate-templates.mjs`
// (or `npm run generate:templates`) after changing anything below, then
// commit the regenerated files under `electron/templates/` — unlike
// `generate-extension-manifest.mjs` this is NOT wired into `prebuild`,
// because `saveDocx` stamps a fresh `docProps/core.xml` modified date on
// every run (see `buildBlankDocx` below), which would otherwise make every
// build touch a tracked binary for no real reason.
//
// Each generated file is meant to be:
//   1. Indistinguishable from a document a real Office/LibreOffice install
//      would produce for "File > New" — closely following the OOXML/ODF
//      packaging rules (correct `[Content_Types].xml`/`_rels` wiring, ODF's
//      `mimetype` stored first and uncompressed) so Word/Excel/PowerPoint/
//      LibreOffice open it with no repair prompt.
//   2. A faithful round trip through ATLAS'S OWN parser/serializer for the
//      formats where Atlas has one:
//        - docx: built from a minimal hand-authored bootstrap package, then
//          parsed with `loadDocx` and immediately re-serialized with
//          `saveDocx` — so the bytes committed here are exactly what
//          Atlas's own DOCX writer produces, not a third-party library's
//          idea of a blank document.
//        - xlsx/ods: built directly with SheetJS (`xlsx` package), the same
//          library `src/viewers/spreadsheet/spreadsheetWrite.ts` writes
//          with, matching what a real Save from an empty spreadsheet
//          document would already produce.
//        - pptx/odp: Atlas's own slide parsers (`src/viewers/slides/**`)
//          need a live DOM (`DOMParser`), which a plain Node script doesn't
//          have — these two are hand-built OOXML/ODF packages instead,
//          modeled directly on the parsers' own test fixtures
//          (`pptxFixture.ts`/`odpFixture.ts`), and round-trip-verified
//          separately by `src/__tests__/templates.test.ts` (which runs
//          under Vitest's jsdom environment).
import { writeFile, mkdir } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

import { createServer } from 'vite'
import JSZip from 'jszip'
import * as XLSX from 'xlsx'

const here = path.dirname(fileURLToPath(import.meta.url))
const repoRoot = path.resolve(here, '..')
const outDir = path.join(repoRoot, 'electron', 'templates')

// `src/docx/index.ts` (and everything it imports) uses TypeScript's
// "bundler" module resolution — extensionless, directory-style imports like
// `from './parser'` resolving to `parser/index.ts` — which plain Node's ESM
// loader doesn't do (it requires an explicit file extension and has no
// directory-index fallback for ESM). Rather than duplicate Atlas's own
// DOCX parser/serializer here (defeating the point of reusing "the
// project's own writers"), this loads that one module through Vite's own
// SSR module graph — the exact same resolver `vite build`/Vitest already
// use for the rest of the codebase — instead of Node's bare loader.
/**
 * @typedef {{
 *   loadDocx: (buffer: ArrayBuffer) => Promise<unknown>,
 *   saveDocx: (bundle: unknown) => Promise<Uint8Array>,
 * }} DocxRoundTripModule
 */

/**
 * @returns {Promise<DocxRoundTripModule>}
 */
async function loadDocxModule() {
  const server = await createServer({
    root: repoRoot,
    configFile: path.join(repoRoot, 'vite.config.ts'),
    server: { middlewareMode: true },
    appType: 'custom',
    logLevel: 'warn',
  })
  try {
    // Vite's SSR module graph types this as `Record<string, any>` since it
    // loads arbitrary modules; `src/docx/index.ts`'s actual exports are
    // `loadDocx`/`saveDocx`, asserted here rather than threading Vite's
    // generic module type through every caller below.
    return /** @type {DocxRoundTripModule} */ (await server.ssrLoadModule('/src/docx/index.ts'))
  } finally {
    await server.close()
  }
}

// ---------------------------------------------------------------------------
// DOCX — minimal bootstrap package, then round-tripped through Atlas's own
// parser + serializer so the committed bytes are Atlas's own writer output.
// ---------------------------------------------------------------------------

const DOCX_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`

const DOCX_PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`

const DOCX_DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p/>
<w:sectPr>
<w:pgSz w:w="12240" w:h="15840"/>
<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
</w:sectPr>
</w:body>
</w:document>`

const DOCX_DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`

const DOCX_STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault/>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
<w:name w:val="Normal"/>
<w:qFormat/>
</w:style>
</w:styles>`

const DOCX_CORE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>`

const DOCX_APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Atlas</Application></Properties>`

/**
 * @param {DocxRoundTripModule} docxModule
 * @returns {Promise<Uint8Array>}
 */
async function buildBlankDocx(docxModule) {
  const { loadDocx, saveDocx } = docxModule
  const bootstrap = new JSZip()
  bootstrap.file('[Content_Types].xml', DOCX_CONTENT_TYPES)
  bootstrap.file('_rels/.rels', DOCX_PACKAGE_RELS)
  bootstrap.file('word/document.xml', DOCX_DOCUMENT_XML)
  bootstrap.file('word/_rels/document.xml.rels', DOCX_DOCUMENT_RELS)
  bootstrap.file('word/styles.xml', DOCX_STYLES_XML)
  bootstrap.file('docProps/core.xml', DOCX_CORE_XML)
  bootstrap.file('docProps/app.xml', DOCX_APP_XML)

  const bootstrapBuffer = await bootstrap.generateAsync({ type: 'arraybuffer' })

  // Round-trip through Atlas's OWN parser + serializer (loadDocx/saveDocx —
  // the exact pair DocxViewer's open/save path uses) so the committed bytes
  // are Atlas's own writer output, guaranteed to parse back cleanly.
  const bundle = await loadDocx(bootstrapBuffer)
  return saveDocx(bundle)
}

// ---------------------------------------------------------------------------
// XLSX / ODS — a single empty "Sheet1", A1 as the only declared range,
// matching the shape of a real Excel/Calc "blank workbook" (no cell data at
// all yet) — built with the same `xlsx` (SheetJS) library
// `spreadsheetWrite.ts` uses to save one.
// ---------------------------------------------------------------------------

/**
 * @param {import('xlsx').BookType} bookType
 * @returns {Buffer}
 */
function buildBlankWorkbookBytes(bookType) {
  const wb = XLSX.utils.book_new()
  const ws = { '!ref': 'A1' }
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')
  return XLSX.write(wb, { type: 'buffer', bookType })
}

/**
 * SheetJS's `bookType: 'ods'` writer does not put `mimetype` first and
 * stored (verified directly: it emits `META-INF/manifest.xml, meta.xml,
 * mimetype, styles.xml, content.xml, manifest.rdf`), which the ODF Package
 * spec (OASIS ODF 1.2 part 3, §2.2) requires. Mirrors the fix in
 * `src/viewers/spreadsheet/spreadsheetWrite.ts`'s `fixOdsPackaging` (see
 * its header for the full explanation) so the committed template matches
 * what a real save now produces.
 *
 * @param {Buffer} bytes
 * @returns {Promise<Buffer | Uint8Array>}
 */
async function fixOdsPackaging(bytes) {
  const original = await JSZip.loadAsync(bytes)
  const mimetype = await original.file('mimetype')?.async('uint8array')
  if (mimetype === undefined) return bytes

  const repacked = new JSZip()
  repacked.file('mimetype', mimetype, { compression: 'STORE' })
  for (const [path, entry] of Object.entries(original.files)) {
    if (path === 'mimetype' || entry.dir) continue
    repacked.file(path, await entry.async('uint8array'))
  }
  return repacked.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

// ---------------------------------------------------------------------------
// PPTX — one 16:9 slide with an empty title placeholder. Modeled on
// `src/viewers/slides/pptx/__tests__/pptxFixture.ts`'s minimal package shape,
// extended with the OPC-required `[Content_Types].xml` + `_rels/.rels` +
// `docProps/*` that fixture skips (it hands its zip straight to the parser,
// bypassing the real "is this a valid package" concerns a file on disk has).
// ---------------------------------------------------------------------------

const PPTX_SLIDE_CX = 12192000 // 16:9 at 96dpi/9525 EMU-per-px == 1280px
const PPTX_SLIDE_CY = 6858000 // == 720px

const PPTX_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`

const PPTX_PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`

const PPTX_PRESENTATION_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster"/></p:sldMasterIdLst>
<p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst>
<p:sldSz cx="${PPTX_SLIDE_CX}" cy="${PPTX_SLIDE_CY}" type="screen16x9"/>
</p:presentation>`

const PPTX_PRESENTATION_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
<Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`

const PPTX_THEME_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Atlas Blank">
<a:themeElements>
<a:clrScheme name="Atlas">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="1F497D"/></a:dk2>
<a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1>
<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
<a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
<a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink>
<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
</a:themeElements>
</a:theme>`

const PPTX_SLIDE_MASTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld>
<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg>
<p:spTree>
<p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="${PPTX_SLIDE_CX - 914400}" cy="1143000"/></a:xfrm></p:spPr><p:txBody><a:p/></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="${PPTX_SLIDE_CX - 914400}" cy="${PPTX_SLIDE_CY - 1975200}"/></a:xfrm></p:spPr><p:txBody><a:p/></p:txBody></p:sp>
</p:spTree>
</p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:txStyles>
<p:titleStyle><a:lvl1pPr><a:defRPr sz="4400" b="1"/></a:lvl1pPr></p:titleStyle>
<p:bodyStyle><a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr></p:bodyStyle>
<p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle>
</p:txStyles>
</p:sldMaster>`

const PPTX_SLIDE_MASTER_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`

const PPTX_SLIDE_LAYOUT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title">
<p:cSld name="Title Slide">
<p:spTree>
<p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:p/></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:p/></p:txBody></p:sp>
</p:spTree>
</p:cSld>
</p:sldLayout>`

const PPTX_SLIDE_LAYOUT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`

// One blank slide with just an (empty) title placeholder — S17's nav label
// falls back to "Slide 1" with no title text, exactly like a fresh Office
// "New Presentation" would show "Click to add title" (a layout-level prompt,
// never actual saved content).
const PPTX_SLIDE1_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld>
<p:spTree>
<p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:p/></p:txBody></p:sp>
</p:spTree>
</p:cSld>
</p:sld>`

const PPTX_SLIDE1_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`

async function buildBlankPptx() {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', PPTX_CONTENT_TYPES)
  zip.file('_rels/.rels', PPTX_PACKAGE_RELS)
  zip.file('docProps/core.xml', DOCX_CORE_XML)
  zip.file('docProps/app.xml', DOCX_APP_XML)
  zip.file('ppt/presentation.xml', PPTX_PRESENTATION_XML)
  zip.file('ppt/_rels/presentation.xml.rels', PPTX_PRESENTATION_RELS)
  zip.file('ppt/theme/theme1.xml', PPTX_THEME_XML)
  zip.file('ppt/slideMasters/slideMaster1.xml', PPTX_SLIDE_MASTER_XML)
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', PPTX_SLIDE_MASTER_RELS)
  zip.file('ppt/slideLayouts/slideLayout1.xml', PPTX_SLIDE_LAYOUT_XML)
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', PPTX_SLIDE_LAYOUT_RELS)
  zip.file('ppt/slides/slide1.xml', PPTX_SLIDE1_XML)
  zip.file('ppt/slides/_rels/slide1.xml.rels', PPTX_SLIDE1_RELS)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } })
}

// ---------------------------------------------------------------------------
// ODP — one 16:9 slide with an empty title frame. Modeled on
// `src/viewers/slides/odp/__tests__/odpFixture.ts`'s minimal package,
// extended with `META-INF/manifest.xml` (ODF's own part manifest — the
// `[Content_Types].xml` equivalent) so a real ODF consumer recognizes it.
// ---------------------------------------------------------------------------

const ODP_MANIFEST_XML = `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.presentation"/>
<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
</manifest:manifest>`

const ODP_STYLES_XML = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  office:version="1.2">
<office:styles>
<style:style style:name="TitleText" style:family="text">
<style:text-properties fo:font-size="32pt" fo:font-weight="bold"/>
</style:style>
</office:styles>
<office:automatic-styles>
<style:page-layout style:name="PM1">
<style:page-layout-properties fo:page-width="1280px" fo:page-height="720px"/>
</style:page-layout>
<style:style style:name="DP1" style:family="drawing-page">
<style:drawing-page-properties draw:fill="solid" draw:fill-color="#FFFFFF"/>
</style:style>
</office:automatic-styles>
<office:master-styles>
<style:master-page style:name="Main" style:page-layout-name="PM1" draw:style-name="DP1"/>
</office:master-styles>
</office:document-styles>`

// A single blank title frame, mirroring what a fresh LibreOffice Impress
// "Title Slide" layout shows before any text is typed (the "Click to add
// title" prompt is layout chrome, never saved content — so the frame's own
// text box is empty here too).
const ODP_CONTENT_XML = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
  office:version="1.2">
<office:body>
<office:presentation>
<draw:page draw:name="Slide 1" draw:master-page-name="Main">
<draw:frame presentation:class="title" svg:x="40px" svg:y="20px" svg:width="1200px" svg:height="100px">
<draw:text-box><text:p/></draw:text-box>
</draw:frame>
</draw:page>
</office:presentation>
</office:body>
</office:document-content>`

async function buildBlankOdp() {
  const zip = new JSZip()
  // ODF: `mimetype` must be the first entry and stored uncompressed, or the
  // file is not recognized as ODF at all — mirrors
  // `office/officePackage.ts`'s `writeOfficePackage` doing the same on save.
  zip.file('mimetype', 'application/vnd.oasis.opendocument.presentation', { compression: 'STORE' })
  zip.file('META-INF/manifest.xml', ODP_MANIFEST_XML)
  zip.file('styles.xml', ODP_STYLES_XML)
  zip.file('content.xml', ODP_CONTENT_XML)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } })
}

// ---------------------------------------------------------------------------

async function main() {
  await mkdir(outDir, { recursive: true })

  const docxModule = await loadDocxModule()
  const [docxBytes, pptxBytes, odpBytes] = await Promise.all([
    buildBlankDocx(docxModule),
    buildBlankPptx(),
    buildBlankOdp(),
  ])

  await writeFile(path.join(outDir, 'blank.docx'), docxBytes)
  await writeFile(path.join(outDir, 'blank.xlsx'), buildBlankWorkbookBytes('xlsx'))
  await writeFile(path.join(outDir, 'blank.ods'), await fixOdsPackaging(buildBlankWorkbookBytes('ods')))
  await writeFile(path.join(outDir, 'blank.pptx'), pptxBytes)
  await writeFile(path.join(outDir, 'blank.odp'), odpBytes)

  console.log(`generate-templates: wrote 5 blank document templates to ${path.relative(repoRoot, outDir)}/`)
}

main().catch((err) => {
  console.error('generate-templates failed:', err)
  process.exitCode = 1
})
