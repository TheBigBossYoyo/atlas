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

/**
 * Collects the value of `@_<attrName>` off every `<tagName .../>` element in
 * the tree, in document order (siblings of the same tag name come back as an
 * array from fast-xml-parser, so those are visited in source order; distinct
 * sibling tag names are visited in the object's own key order, which
 * fast-xml-parser assigns in the order it encounters them in the source —
 * i.e. also document order). An element present with no `attrName` attribute
 * at all (e.g. a bare `<w:vMerge/>`, whose OOXML-spec default is
 * `"continue"`) records `attrDefault` instead of being silently skipped, so
 * dropping the attribute doesn't just quietly shrink the array — it changes
 * a recorded value, which `toEqual` still catches.
 */
function collectAttrValues(
  node: unknown,
  tagName: string,
  attrName: string,
  attrDefault: string,
): ReadonlyArray<string> {
  const attrKey = `@_${attrName}`
  const values: string[] = []

  function walk(current: unknown): void {
    if (current === null || typeof current !== 'object') return
    if (Array.isArray(current)) {
      for (const item of current) walk(item)
      return
    }
    for (const [key, value] of Object.entries(current as Record<string, unknown>)) {
      if (key === tagName) {
        for (const element of Array.isArray(value) ? value : [value]) {
          if (element !== null && typeof element === 'object') {
            const attrs = element as Record<string, unknown>
            values.push(attrKey in attrs ? String(attrs[attrKey]) : attrDefault)
          } else {
            // A self-closing element with no attributes at all parses to a
            // primitive (empty string), not an object with no matching key.
            values.push(attrDefault)
          }
        }
      }
      walk(value)
    }
  }

  walk(node)
  return values
}

export interface DocumentStructuralCounts {
  readonly paragraphs: number
  readonly tables: number
  readonly drawings: number
  readonly sectPr: number
  readonly tblGrid: number
  /**
   * Element occurrence counts — TEST-2 (phase4/b4-corpus-tests): before this,
   * `w:ins`/`w:del` weren't inspected at all, so a save that silently
   * dropped all tracked-change wrapping left the five counts above
   * unchanged and passed silently.
   */
  readonly insCount: number
  readonly delCount: number
  /**
   * Attribute-level values, in document order — TEST-2: the counts above
   * can't distinguish "list item flattened from level 2 to level 0" (same
   * number of `w:ilvl` elements, different values) from an unchanged
   * document, nor "vertical merge silently turned into a plain cell"
   * (`w:gridSpan`/`w:vMerge` values changed, or degraded to the default,
   * without the element count changing). `runSzValues`/`runColorValues`
   * cover the same gap for direct run formatting (`w:sz`/`w:color` inside a
   * `w:rPr`) — e.g. a save that silently coerced every run's font size or
   * color to some default.
   */
  readonly ilvlValues: ReadonlyArray<string>
  readonly numIdValues: ReadonlyArray<string>
  readonly gridSpanValues: ReadonlyArray<string>
  readonly vMergeValues: ReadonlyArray<string>
  readonly runSzValues: ReadonlyArray<string>
  readonly runColorValues: ReadonlyArray<string>
}

/**
 * A structural fingerprint of `word/document.xml`: element counts plus
 * attribute-level values that should be unchanged by a no-op save or by the
 * corpus's per-fixture trivial edit (see roundtrip.corpus.test.ts's
 * `EDIT_POSITIONS`) — none of those insert/remove a paragraph, table,
 * drawing, section, table grid, tracked change, list item, merged cell, or
 * run-formatting property. Computed from the raw XML text, independent of
 * Atlas's own parser, so it also catches a part being silently dropped or
 * corrupted on the way through Atlas's model (e.g. DXS-02's `w:tblGrid`
 * loss, or a regression that flattened every `w:ilvl` to 0).
 */
export function documentStructuralCounts(documentXml: string): DocumentStructuralCounts {
  const tree = xmlParser.parse(documentXml) as unknown
  return {
    paragraphs: countTag(tree, 'w:p'),
    tables: countTag(tree, 'w:tbl'),
    drawings: countTag(tree, 'w:drawing'),
    sectPr: countTag(tree, 'w:sectPr'),
    tblGrid: countTag(tree, 'w:tblGrid'),
    insCount: countTag(tree, 'w:ins'),
    delCount: countTag(tree, 'w:del'),
    ilvlValues: collectAttrValues(tree, 'w:ilvl', 'w:val', '0'),
    numIdValues: collectAttrValues(tree, 'w:numId', 'w:val', ''),
    gridSpanValues: collectAttrValues(tree, 'w:gridSpan', 'w:val', '1'),
    // OOXML: a `w:vMerge` with no `w:val` means "continue" the merge above.
    vMergeValues: collectAttrValues(tree, 'w:vMerge', 'w:val', 'continue'),
    runSzValues: collectAttrValues(tree, 'w:sz', 'w:val', ''),
    runColorValues: collectAttrValues(tree, 'w:color', 'w:val', ''),
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
