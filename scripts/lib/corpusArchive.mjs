// Post-processing helpers for the DOCX corpus generator
// (scripts/generate-docx-corpus.mjs). The `docx` npm package covers most of
// the OOXML surface; these helpers patch in the handful of constructs it
// cannot express (content controls, mc:AlternateContent, comment replies, a
// custom banded table style) and normalize the package so re-running the
// generator produces byte-identical output.
import JSZip from 'jszip'

// Fixed instant used everywhere a timestamp would otherwise be
// `new Date()` (docProps/core.xml, zip entry mtimes) so the generator is
// deterministic across runs.
export const FIXED_DATE = new Date(Date.UTC(2020, 0, 1, 0, 0, 0))
export const FIXED_DATE_ISO = FIXED_DATE.toISOString().replace(/\.\d{3}Z$/, 'Z')

const TEXT_EXTENSIONS = new Set(['.xml', '.rels'])

function isTextPath(path) {
  return TEXT_EXTENSIONS.has(path.slice(path.lastIndexOf('.')))
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

/**
 * Loads a `.docx` (or any OPC) buffer into a map of path -> text/Buffer,
 * decoding every `.xml`/`.rels` part as UTF-8 text and leaving binary parts
 * (media, embeddings) as Buffers.
 *
 * @param {Buffer} buffer
 * @returns {Promise<Map<string, string | Buffer>>}
 */
export async function unzipToFiles(buffer) {
  const zip = await JSZip.loadAsync(buffer)
  const files = new Map()
  const entries = Object.values(zip.files).filter((entry) => !entry.dir)
  for (const entry of entries) {
    if (isTextPath(entry.name)) {
      files.set(entry.name, await entry.async('string'))
    } else {
      files.set(entry.name, await entry.async('nodebuffer'))
    }
  }
  return files
}

/**
 * Re-zips a path -> text/Buffer map into a deterministic `.docx` buffer:
 * sorted entry order, a fixed mtime on every entry, and stable DEFLATE
 * compression settings.
 *
 * @param {Map<string, string | Buffer>} files
 * @returns {Promise<Buffer>}
 */
export async function zipFromFiles(files) {
  const zip = new JSZip()
  for (const path of [...files.keys()].sort()) {
    // `createFolders: false` matters for determinism: JSZip otherwise
    // auto-vivifies parent directory entries stamped with `new Date()`,
    // which real `.docx` zips don't have anyway.
    zip.file(path, files.get(path), { date: FIXED_DATE, createFolders: false })
  }
  return zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
    platform: 'UNIX',
  })
}

/**
 * Overwrites docProps/core.xml's `dcterms:created`/`dcterms:modified` with a
 * fixed instant so the generator's output does not depend on wall-clock
 * time.
 *
 * @param {Map<string, string | Buffer>} files
 */
export function stabilizeCoreProperties(files) {
  const path = 'docProps/core.xml'
  const xml = files.get(path)
  if (typeof xml !== 'string') return
  const stabilized = xml
    .replace(/(<dcterms:created[^>]*>)[^<]*(<\/dcterms:created>)/, `$1${FIXED_DATE_ISO}$2`)
    .replace(/(<dcterms:modified[^>]*>)[^<]*(<\/dcterms:modified>)/, `$1${FIXED_DATE_ISO}$2`)
  files.set(path, stabilized)
}

/**
 * The `docx` package assigns external-hyperlink relationship ids via a
 * random nanoid (`rId<random>`) rather than a sequential counter, which
 * makes any fixture containing an `ExternalHyperlink` non-deterministic
 * across generator runs. This finds every External hyperlink relationship
 * in every `.rels` part, renames it to a stable `rIdHyperlinkN` id (in file
 * then document order), and rewrites every reference to that id in the
 * `.rels` file and the part it belongs to.
 *
 * @param {Map<string, string | Buffer>} files
 */
export function stabilizeHyperlinkRelationshipIds(files) {
  const relsPaths = [...files.keys()].filter((path) => path.endsWith('.rels')).sort()
  let counter = 0

  for (const relsPath of relsPaths) {
    const relsXml = files.get(relsPath)
    if (typeof relsXml !== 'string') continue

    const hyperlinkIdPattern =
      /<Relationship\s+Id="([^"]+)"[^>]*Type="[^"]*\/hyperlink"[^>]*TargetMode="External"[^>]*\/>/g
    const originalIds = [...relsXml.matchAll(hyperlinkIdPattern)].map((match) => match[1])
    if (originalIds.length === 0) continue

    // `.rels` files live in a `_rels` subdirectory next to the part they
    // describe, e.g. `word/_rels/document.xml.rels` -> `word/document.xml`.
    const partPath = relsPath.replace(/\/_rels\/([^/]+)\.rels$/, '/$1')
    const partXml = files.get(partPath)

    let updatedRels = relsXml
    let updatedPart = typeof partXml === 'string' ? partXml : undefined

    for (const originalId of originalIds) {
      counter += 1
      const stableId = `rIdHyperlink${counter}`
      const idPattern = new RegExp(`(["'])${escapeRegExp(originalId)}(["'])`, 'g')
      updatedRels = updatedRels.replace(idPattern, `$1${stableId}$2`)
      if (updatedPart !== undefined) {
        updatedPart = updatedPart.replace(idPattern, `$1${stableId}$2`)
      }
    }

    files.set(relsPath, updatedRels)
    if (updatedPart !== undefined) {
      files.set(partPath, updatedPart)
    }
  }
}

/**
 * Runs every determinism-normalizing pass over a freshly-packed archive.
 *
 * @param {Map<string, string | Buffer>} files
 */
export function stabilizeArchive(files) {
  stabilizeCoreProperties(files)
  stabilizeHyperlinkRelationshipIds(files)
}

const AUTO_GENERATED_FOOTNOTE_ENDNOTE_PARTS = [
  'word/footnotes.xml',
  'word/endnotes.xml',
  'word/_rels/footnotes.xml.rels',
  'word/_rels/endnotes.xml.rels',
]

/**
 * The `docx` package (matching real Word) always emits `word/footnotes.xml`
 * / `word/endnotes.xml` containing only the boilerplate `separator` /
 * `continuationSeparator` entries every document needs, even when the
 * document itself never references a footnote or endnote. Atlas's
 * footnote/endnote parser currently represents every note's body as an
 * opaque `UnknownNode` ("Option B", see src/docx/parser/footnotes.ts), and
 * its serializer refuses to write a standalone part containing anything but
 * paragraph blocks — so simply loading and re-saving *any* real-world
 * `.docx` throws today (DXS-01), independent of whether that document
 * visibly uses footnotes.
 *
 * The corpus fixture built to exercise that failure on purpose
 * (`footnotes-endnotes.docx`) keeps these parts. Every other fixture strips
 * them so its round-trip failure surface stays about the feature it's
 * actually testing, rather than being masked by this one already-tracked,
 * already-dedicated-fixture bug.
 *
 * @param {Map<string, string | Buffer>} files
 */
export function stripAutoGeneratedFootnotesEndnotes(files) {
  for (const part of AUTO_GENERATED_FOOTNOTE_ENDNOTE_PARTS) {
    files.delete(part)
  }

  const contentTypesPath = '[Content_Types].xml'
  const contentTypesXml = files.get(contentTypesPath)
  if (typeof contentTypesXml === 'string') {
    files.set(
      contentTypesPath,
      contentTypesXml
        .replace(/<Override[^>]*PartName="\/word\/footnotes\.xml"[^>]*\/>/, '')
        .replace(/<Override[^>]*PartName="\/word\/endnotes\.xml"[^>]*\/>/, ''),
    )
  }

  const documentRelsPath = 'word/_rels/document.xml.rels'
  const documentRelsXml = files.get(documentRelsPath)
  if (typeof documentRelsXml === 'string') {
    files.set(
      documentRelsPath,
      documentRelsXml
        .replace(/<Relationship[^>]*Type="[^"]*\/footnotes"[^>]*\/>/, '')
        .replace(/<Relationship[^>]*Type="[^"]*\/endnotes"[^>]*\/>/, ''),
    )
  }
}

/**
 * Replaces the first `<w:p>...</w:p>` element containing `marker` with the
 * result of `wrap(paragraphXml)`. Used to splice hand-written OOXML (a
 * content-control wrapper, for instance) around a paragraph that `docx`
 * generated normally.
 *
 * @param {string} xml
 * @param {string} marker
 * @param {(paragraphXml: string) => string} wrap
 * @returns {string}
 */
export function wrapParagraphContaining(xml, marker, wrap) {
  const paragraphPattern = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g
  let replaced = false
  const result = xml.replace(paragraphPattern, (match) => {
    if (!replaced && match.includes(marker)) {
      replaced = true
      return wrap(match)
    }
    return match
  })
  if (!replaced) {
    throw new Error(`wrapParagraphContaining: marker not found: ${marker}`)
  }
  return result
}

/**
 * Replaces the first `<w:r>...</w:r>` element containing `marker` with raw
 * replacement XML. Used to splice a hand-written run (an
 * `mc:AlternateContent` shape, for instance) in place of a placeholder run
 * that `docx` generated normally.
 *
 * @param {string} xml
 * @param {string} marker
 * @param {string} replacementXml
 * @returns {string}
 */
export function replaceRunContaining(xml, marker, replacementXml) {
  const runPattern = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g
  let replaced = false
  const result = xml.replace(runPattern, (match) => {
    if (!replaced && match.includes(marker)) {
      replaced = true
      return replacementXml
    }
    return match
  })
  if (!replaced) {
    throw new Error(`replaceRunContaining: marker not found: ${marker}`)
  }
  return result
}

// ---------------------------------------------------------------------------
// OPC package surgery: relationships + content types
//
// The handful of real-world OOXML features `docx` has no API for at all
// (custom XML parts, VBA/macro projects, embedded-font parts with an actual
// font binary, EMF/WMF media) need a brand-new part plus its own
// relationship and content-type entry, not just a text patch inside an
// existing part. These three helpers do that generically so each corpus
// fixture only has to say *what* part/relationship/content-type it wants,
// not *how* to splice OPC XML.
// ---------------------------------------------------------------------------

/**
 * Adds a `<Relationship>` entry to an already-existing `.rels` part (e.g.
 * `word/_rels/document.xml.rels`). `id` must be unique within that part
 * (relationship ids are opaque strings per the OPC schema — they need not
 * match the conventional `rIdN` shape `docx` itself generates).
 *
 * @param {Map<string, string | Buffer>} files
 * @param {string} relsPath
 * @param {string} id
 * @param {string} type
 * @param {string} target
 * @param {string} [targetMode]
 */
export function addRelationshipEntry(files, relsPath, id, type, target, targetMode) {
  const xml = files.get(relsPath)
  if (typeof xml !== 'string' || !xml.includes('</Relationships>')) {
    throw new Error(`addRelationshipEntry: expected ${relsPath} to contain a closing </Relationships> tag`)
  }
  const modeAttr = targetMode ? ` TargetMode="${targetMode}"` : ''
  const entry = `<Relationship Id="${id}" Type="${type}" Target="${target}"${modeAttr}/>`
  files.set(relsPath, xml.replace('</Relationships>', `${entry}</Relationships>`))
}

/**
 * Adds a `<Default>` extension mapping to `[Content_Types].xml`, unless that
 * extension already has one (real Word packages never declare the same
 * extension twice — `docx` itself already registers `png`/`jpeg`/`xml`/etc.).
 *
 * @param {Map<string, string | Buffer>} files
 * @param {string} extension
 * @param {string} contentType
 */
export function addContentTypeDefault(files, extension, contentType) {
  const path = '[Content_Types].xml'
  const xml = files.get(path)
  if (typeof xml !== 'string' || !xml.includes('</Types>')) {
    throw new Error(`addContentTypeDefault: expected ${path} to contain a closing </Types> tag`)
  }
  if (new RegExp(`<Default\\b[^>]*Extension="${extension}"`, 'i').test(xml)) {
    return
  }
  const entry = `<Default Extension="${extension}" ContentType="${contentType}"/>`
  files.set(path, xml.replace('</Types>', `${entry}</Types>`))
}

/**
 * Adds an `<Override>` part-specific content-type entry to
 * `[Content_Types].xml` (used for a part whose content type can't be
 * inferred from its extension alone, e.g. `customXml/itemProps1.xml`).
 *
 * @param {Map<string, string | Buffer>} files
 * @param {string} partName - Package-absolute part name, e.g. `/customXml/itemProps1.xml`.
 * @param {string} contentType
 */
export function addContentTypeOverride(files, partName, contentType) {
  const path = '[Content_Types].xml'
  const xml = files.get(path)
  if (typeof xml !== 'string' || !xml.includes('</Types>')) {
    throw new Error(`addContentTypeOverride: expected ${path} to contain a closing </Types> tag`)
  }
  const entry = `<Override PartName="${partName}" ContentType="${contentType}"/>`
  files.set(path, xml.replace('</Types>', `${entry}</Types>`))
}
