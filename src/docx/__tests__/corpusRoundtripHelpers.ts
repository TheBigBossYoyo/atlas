// Shared helpers for src/docx/__tests__/roundtrip.corpus.test.ts (P0.5 /
// QA-09 / DXS-17). Not itself a test file (Vitest only collects
// `*.test.ts(x)`), so it can hold plain utility functions the corpus suite
// needs: OPC-level structural inspection of a `.docx` buffer, independent of
// Atlas's own DOCX model.
import { readdir, readFile } from 'node:fs/promises'
import path from 'node:path'

import { XMLParser, XMLValidator } from 'fast-xml-parser'
import JSZip from 'jszip'

export const CORPUS_DIR = path.resolve(process.cwd(), 'src/docx/__fixtures__/corpus')

/** Lists the corpus fixture ids (filenames without the `.docx` extension), sorted. */
export async function listCorpusFixtureIds(): Promise<ReadonlyArray<string>> {
  const entries = await readdir(CORPUS_DIR)
  return entries
    .filter((name) => name.endsWith('.docx'))
    .map((name) => name.slice(0, -'.docx'.length))
    .sort()
}

/**
 * Reads a corpus fixture and returns it as an `ArrayBuffer` backed by
 * *this* JS realm's `Uint8Array`/`ArrayBuffer` constructors.
 *
 * Vitest's `jsdom` environment gives the test module its own `ArrayBuffer`
 * global, distinct from the one backing a `Buffer` returned by Node's
 * `fs.readFile` — so `nodeBuffer.buffer instanceof ArrayBuffer` is `false`
 * here, and JSZip (used by both this test's helpers and `loadDocx` itself)
 * refuses to load such a buffer with an opaque "Can't read the data of the
 * loaded zip file" error. Copying through `Uint8Array.from` rebuilds the
 * bytes with the realm's own constructors, which JSZip accepts.
 */
export async function readCorpusFixture(fixtureId: string): Promise<ArrayBuffer> {
  const bytes = await readFile(path.join(CORPUS_DIR, `${fixtureId}.docx`))
  return Uint8Array.from(bytes).buffer
}

// ---------------------------------------------------------------------------
// Raw OPC package inspection (independent of Atlas's own DOCX parser)
// ---------------------------------------------------------------------------

export type RawPackage = ReadonlyMap<string, Buffer>

const TEXT_PART_PATTERN = /\.(xml|rels)$/i

export async function loadRawPackage(buffer: ArrayBuffer | Buffer | Uint8Array): Promise<RawPackage> {
  const zip = await JSZip.loadAsync(buffer)
  const files = new Map<string, Buffer>()
  const entries = Object.values(zip.files).filter((entry) => !entry.dir)
  for (const entry of entries) {
    files.set(entry.name, await entry.async('nodebuffer'))
  }
  return files
}

export function textOf(pkg: RawPackage, partPath: string): string | undefined {
  const bytes = pkg.get(partPath)
  return bytes === undefined ? undefined : bytes.toString('utf-8')
}

/**
 * Every `.xml`/`.rels` part in the package must parse as well-formed XML.
 * Returns the list of `path: message` failures (empty when everything is
 * well-formed).
 */
export function findMalformedXmlParts(pkg: RawPackage): ReadonlyArray<string> {
  const problems: string[] = []
  for (const [partPath, bytes] of pkg.entries()) {
    if (!TEXT_PART_PATTERN.test(partPath)) continue
    const result = XMLValidator.validate(bytes.toString('utf-8'))
    if (result !== true) {
      problems.push(`${partPath}: ${result.err.msg}`)
    }
  }
  return problems
}

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

function countTag(node: unknown, tagName: string): number {
  if (node === null || typeof node !== 'object') return 0
  if (Array.isArray(node)) {
    return node.reduce((sum: number, item) => sum + countTag(item, tagName), 0)
  }
  let total = 0
  for (const [key, value] of Object.entries(node as Record<string, unknown>)) {
    if (key === tagName) {
      total += Array.isArray(value) ? value.length : 1
    }
    total += countTag(value, tagName)
  }
  return total
}

export interface DocumentStructuralCounts {
  readonly paragraphs: number
  readonly tables: number
  readonly drawings: number
  readonly sectPr: number
  readonly tblGrid: number
}

/**
 * A small structural fingerprint of `word/document.xml`: element counts that
 * should be unchanged by a no-op save or by inserting text into the first
 * paragraph (neither adds/removes paragraphs, tables, drawings, sections, or
 * table grids). Computed from the raw XML text, independent of Atlas's own
 * parser, so it also catches a part being silently dropped or corrupted on
 * the way through Atlas's model (e.g. DXS-02's `w:tblGrid` loss).
 */
export function documentStructuralCounts(documentXml: string): DocumentStructuralCounts {
  const tree = xmlParser.parse(documentXml) as unknown
  return {
    paragraphs: countTag(tree, 'w:p'),
    tables: countTag(tree, 'w:tbl'),
    drawings: countTag(tree, 'w:drawing'),
    sectPr: countTag(tree, 'w:sectPr'),
    tblGrid: countTag(tree, 'w:tblGrid'),
  }
}

export function documentXmlOf(pkg: RawPackage): string {
  const xml = textOf(pkg, 'word/document.xml')
  if (xml === undefined) {
    throw new Error('Package is missing word/document.xml')
  }
  return xml
}

// ---------------------------------------------------------------------------
// Relationship + content-type validation
// ---------------------------------------------------------------------------

interface ParsedRelationship {
  readonly id: string
  readonly type: string
  readonly target: string
  readonly targetMode?: string
}

function parseRelationshipsXml(xml: string): ReadonlyArray<ParsedRelationship> {
  const tree = xmlParser.parse(xml) as {
    Relationships?: { Relationship?: unknown }
  }
  const raw = tree.Relationships?.Relationship
  if (raw === undefined) return []
  const list = Array.isArray(raw) ? raw : [raw]
  return list.map((entry) => {
    const rel = entry as Record<string, unknown>
    return {
      id: String(rel['@_Id']),
      type: String(rel['@_Type']),
      target: String(rel['@_Target']),
      targetMode: rel['@_TargetMode'] === undefined ? undefined : String(rel['@_TargetMode']),
    }
  })
}

interface ContentTypesInfo {
  readonly defaults: ReadonlySet<string>
  readonly overrides: ReadonlySet<string>
}

function parseContentTypesXml(xml: string): ContentTypesInfo {
  const tree = xmlParser.parse(xml) as {
    Types?: { Default?: unknown; Override?: unknown }
  }
  const asList = (value: unknown): Record<string, unknown>[] =>
    value === undefined ? [] : ((Array.isArray(value) ? value : [value]) as Record<string, unknown>[])

  const defaults = new Set(
    asList(tree.Types?.Default).map((entry) => String(entry['@_Extension']).toLowerCase()),
  )
  const overrides = new Set(asList(tree.Types?.Override).map((entry) => String(entry['@_PartName'])))
  return { defaults, overrides }
}

function resolveRelationshipTarget(relsPartPath: string, target: string): string {
  // `word/_rels/document.xml.rels` describes parts relative to `word/`.
  const baseDir = path.posix.dirname(path.posix.dirname(relsPartPath))
  const resolved = target.startsWith('/') ? target.slice(1) : path.posix.join(baseDir, target)
  return path.posix.normalize(resolved)
}

/**
 * For every internal (non-External) relationship in every `.rels` part:
 * the target part must exist in the package, and must have a resolvable
 * content type (an `Override` for its exact part name, or a `Default` for
 * its extension). Returns the list of problems found (empty when the
 * package is OPC-valid on this dimension).
 */
export function findRelationshipProblems(pkg: RawPackage): ReadonlyArray<string> {
  const contentTypesXml = textOf(pkg, '[Content_Types].xml')
  if (contentTypesXml === undefined) {
    return ['[Content_Types].xml is missing']
  }
  const contentTypes = parseContentTypesXml(contentTypesXml)
  const problems: string[] = []

  for (const [partPath, bytes] of pkg.entries()) {
    if (!partPath.endsWith('.rels')) continue
    const relationships = parseRelationshipsXml(bytes.toString('utf-8'))

    for (const rel of relationships) {
      if (rel.targetMode === 'External') continue

      const resolvedTarget = resolveRelationshipTarget(partPath, rel.target)
      if (!pkg.has(resolvedTarget)) {
        problems.push(`${partPath}: relationship ${rel.id} target "${resolvedTarget}" does not exist`)
        continue
      }

      const extension = resolvedTarget.slice(resolvedTarget.lastIndexOf('.') + 1).toLowerCase()
      const hasOverride = contentTypes.overrides.has(`/${resolvedTarget}`)
      const hasDefault = contentTypes.defaults.has(extension)
      if (!hasOverride && !hasDefault) {
        problems.push(`${partPath}: relationship ${rel.id} target "${resolvedTarget}" has no content type`)
      }
    }
  }

  return problems
}

// ---------------------------------------------------------------------------
// Passthrough-part fidelity
// ---------------------------------------------------------------------------

const ALWAYS_OWNED_PARTS: ReadonlySet<string> = new Set([
  'word/document.xml',
  'word/styles.xml',
  'word/numbering.xml',
  'word/comments.xml',
  'word/footnotes.xml',
  'word/endnotes.xml',
  'word/_rels/document.xml.rels',
  '_rels/.rels',
  '[Content_Types].xml',
  // D19 / DXS-13: saveDocx now regenerates dcterms:modified/cp:lastModifiedBy
  // on every save, so this part legitimately changes even for a "no-op"
  // round trip.
  'docProps/core.xml',
])

/**
 * Parts Atlas's serializer intentionally rewrites (see src/docx/index.ts's
 * `saveDocx`): the fixed set plus whichever header/footer parts the
 * ORIGINAL package's own relationships point at. Everything else is
 * passthrough and must survive a save byte-for-byte.
 */
export function ownedPartPaths(originalPkg: RawPackage): ReadonlySet<string> {
  const owned = new Set(ALWAYS_OWNED_PARTS)
  const relsXml = textOf(originalPkg, 'word/_rels/document.xml.rels')
  if (relsXml === undefined) return owned

  for (const rel of parseRelationshipsXml(relsXml)) {
    if (rel.type.endsWith('/header') || rel.type.endsWith('/footer')) {
      owned.add(resolveRelationshipTarget('word/_rels/document.xml.rels', rel.target))
    }
  }
  return owned
}

/**
 * Every part in `originalPkg` that isn't in `owned` must exist in
 * `resultPkg` with byte-identical content. Returns the list of paths that
 * differ or are missing (empty when passthrough fidelity holds).
 */
export function findPassthroughMismatches(
  originalPkg: RawPackage,
  resultPkg: RawPackage,
  owned: ReadonlySet<string>,
): ReadonlyArray<string> {
  const mismatches: string[] = []
  for (const [partPath, originalBytes] of originalPkg.entries()) {
    if (owned.has(partPath)) continue
    const resultBytes = resultPkg.get(partPath)
    if (resultBytes === undefined) {
      mismatches.push(`${partPath}: missing from the saved package`)
    } else if (!resultBytes.equals(originalBytes)) {
      mismatches.push(`${partPath}: content changed by a save that should have left it untouched`)
    }
  }
  return mismatches
}
