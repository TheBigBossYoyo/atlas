// Atlas — structural validation harness for generated Office/ODF packages.
//
// Checks a .docx/.xlsx/.pptx/.odt/.odp/.ods file against the actual OPC/
// OOXML/ODF package specs, independent of Atlas's own readers/writers:
//
//  - the zip container itself (entry names, no duplicates, ODF's mimetype
//    first-and-stored rule) via `zipReader.mjs`'s own from-scratch PKZIP
//    parser, not JSZip;
//  - every XML/rels part is well-formed and free of characters XML 1.0
//    forbids;
//  - namespace prefixes used in a part are actually declared somewhere in
//    scope;
//  - OPC: `[Content_Types].xml` covers every part (Default or Override, no
//    duplicates), every `.rels` target exists and has a content type, every
//    `r:id`-shaped attribute a part uses resolves to a relationship that
//    part's own `.rels` file actually declares, required parts are present;
//  - ODF: `META-INF/manifest.xml` lists every file (and nothing that isn't
//    one), required parts are present;
//  - schema-mandated element order in the three spots most prone to silent
//    breakage: `w:pPr` first in `w:p`, `w:tblPr`/`w:tblGrid`/`w:tr` order in
//    `w:tbl`, `a:bodyPr`/`a:lstStyle`/`a:p` order in `p:txBody`, and the
//    OOXML `CT_Worksheet` child sequence.
//
// Used both by `scripts/validate-office-file.mjs` (CLI) and by unit tests
// (`src/**/__tests__/*.test.ts`, which import this module directly — see
// `officeValidator.d.mts` for its type surface there).
import { XMLParser, XMLValidator } from 'fast-xml-parser'

import { readZipArchive } from './zipReader.mjs'

// ---------------------------------------------------------------------------
// Issue collection
// ---------------------------------------------------------------------------

/** @typedef {{ severity: 'error' | 'warning', code: string, part: string | null, message: string }} Issue */

function issue(severity, code, part, message) {
  return { severity, code, part, message }
}

// ---------------------------------------------------------------------------
// XML 1.0 character legality + well-formedness
// ---------------------------------------------------------------------------

// XML 1.0 Char production: #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD]
// | [#x10000-#x10FFFF]. Everything else -- including the C0 controls Atlas's
// own `xmlSafeText` strips at write time -- is illegal in a well-formed XML
// document (a permissive parser may accept it; Word/Excel will not).
//
// Expressed as explicit numeric code-unit ranges (checked in a plain loop
// below), not a regex-with-escapes literal: several of these code points
// (the C0 controls, the two noncharacters at the top of the BMP) are
// themselves invisible or easily mis-typed, and this file has already once
// ended up with literal raw control bytes sitting silently inside a regex
// character class here (still ran correctly, but unreadable and fragile to
// the next edit) after a copy/paste through a layer that decodes escape
// sequences before the text reaches disk. Numeric comparisons avoid the
// whole class of problem.
const ILLEGAL_CODE_UNIT_RANGES = [
  [0x00, 0x08], // C0 controls before TAB
  [0x0b, 0x0c], // vertical tab, form feed
  [0x0e, 0x1f], // C0 controls after CR, before space
  [0xfffe, 0xffff], // noncharacters
]

function isIllegalCodeUnit(code) {
  return ILLEGAL_CODE_UNIT_RANGES.some(([lo, hi]) => code >= lo && code <= hi)
}

function findIllegalXmlChars(text) {
  const found = []
  for (let i = 0; i < text.length && found.length < 5; i++) {
    const code = text.charCodeAt(i)
    if (isIllegalCodeUnit(code)) {
      found.push({ index: i, codePoint: code })
      continue
    }
    // Unpaired surrogates are also illegal (not valid Unicode scalar values).
    const isHighSurrogate = code >= 0xd800 && code <= 0xdbff
    const isLowSurrogate = code >= 0xdc00 && code <= 0xdfff
    if (isHighSurrogate) {
      const next = text.charCodeAt(i + 1)
      if (!(next >= 0xdc00 && next <= 0xdfff)) found.push({ index: i, codePoint: code })
    } else if (isLowSurrogate) {
      const prev = text.charCodeAt(i - 1)
      if (!(prev >= 0xd800 && prev <= 0xdbff)) found.push({ index: i, codePoint: code })
    }
  }
  return found
}

const orderedXmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  trimValues: false,
})

/** Returns `{ ok: true, tree }` or `{ ok: false, message }`. Never throws. */
function parseOrdered(xmlText) {
  const validation = XMLValidator.validate(xmlText, { allowBooleanAttributes: true })
  if (validation !== true) {
    return { ok: false, message: `${validation.err.msg} (line ${validation.err.line}, col ${validation.err.col})` }
  }
  try {
    return { ok: true, tree: orderedXmlParser.parse(xmlText) }
  } catch (cause) {
    return { ok: false, message: cause instanceof Error ? cause.message : String(cause) }
  }
}

function nodeName(node) {
  for (const key of Object.keys(node)) {
    if (key !== ':@') return key
  }
  return undefined
}

function isElementNode(node) {
  const name = nodeName(node)
  return name !== undefined && name !== '#text' && !name.startsWith('?')
}

function elementChildren(node) {
  const name = nodeName(node)
  const children = name === undefined ? undefined : node[name]
  return Array.isArray(children) ? children.filter(isElementNode) : []
}

function attributesOf(node) {
  return node[':@'] ?? {}
}

/** Recursively collects every element node anywhere in the tree whose tag equals `tag`. */
function collectByTag(nodes, tag, out = []) {
  for (const node of nodes) {
    if (!isElementNode(node)) continue
    if (nodeName(node) === tag) out.push(node)
    const name = nodeName(node)
    const children = node[name]
    if (Array.isArray(children)) collectByTag(children, tag, out)
  }
  return out
}

// ---------------------------------------------------------------------------
// Namespace-prefix-declared check
// ---------------------------------------------------------------------------

const RESERVED_PREFIXES = new Set(['xml', 'xmlns'])

function declaredPrefixesFrom(attrs) {
  const declared = []
  for (const key of Object.keys(attrs)) {
    // attributeNamePrefix is '@_'; xmlns:foo="..." -> '@_xmlns:foo'.
    const bare = key.slice(2)
    if (bare === 'xmlns') declared.push('') // default namespace
    else if (bare.startsWith('xmlns:')) declared.push(bare.slice('xmlns:'.length))
  }
  return declared
}

function prefixOf(qualifiedName) {
  const colon = qualifiedName.indexOf(':')
  return colon === -1 ? '' : qualifiedName.slice(0, colon)
}

function checkNamespacesDeclared(nodes, scope, partPath, issues) {
  for (const node of nodes) {
    if (!isElementNode(node)) continue
    const name = nodeName(node)
    const attrs = attributesOf(node)
    const localScope = new Set([...scope, ...declaredPrefixesFrom(attrs)])

    const elementPrefix = prefixOf(name)
    if (elementPrefix !== '' && !RESERVED_PREFIXES.has(elementPrefix) && !localScope.has(elementPrefix)) {
      issues.push(
        issue('error', 'namespace-not-declared', partPath, `Element <${name}> uses undeclared namespace prefix "${elementPrefix}:".`),
      )
    }
    for (const key of Object.keys(attrs)) {
      const attrName = key.slice(2) // strip '@_'
      if (attrName === 'xmlns' || attrName.startsWith('xmlns:')) continue
      const attrPrefix = prefixOf(attrName)
      if (attrPrefix !== '' && !RESERVED_PREFIXES.has(attrPrefix) && !localScope.has(attrPrefix)) {
        issues.push(
          issue(
            'error',
            'namespace-not-declared',
            partPath,
            `Attribute "${attrName}" on <${name}> uses undeclared namespace prefix "${attrPrefix}:".`,
          ),
        )
      }
    }

    const children = node[name]
    if (Array.isArray(children)) checkNamespacesDeclared(children, localScope, partPath, issues)
  }
}

// ---------------------------------------------------------------------------
// Element-order checks
// ---------------------------------------------------------------------------

/** `tag`, if it occurs among `node`'s element children at all, must be the very first one. */
function assertFirstIfPresent(node, tag, partPath, issues, schemaNote) {
  const children = elementChildren(node)
  const index = children.findIndex((child) => nodeName(child) === tag)
  if (index > 0) {
    issues.push(
      issue(
        'error',
        'element-order',
        partPath,
        `<${nodeName(node)}> has <${tag}> at child position ${index}, but the schema requires it first${schemaNote ? ` (${schemaNote})` : ''}.`,
      ),
    )
  }
}

/**
 * Every element child of `node` whose tag is in `canonicalOrder` must appear
 * in that relative order (elements not in the list are ignored, so this
 * only asserts ordering among the recognized set — the right level of
 * strictness for a large closed content model like CT_Worksheet where an
 * unrecognized/future element shouldn't fail the check on its own).
 */
function assertCanonicalOrder(node, canonicalOrder, partPath, issues, schemaLabel) {
  const rank = new Map(canonicalOrder.map((tag, i) => [tag, i]))
  const children = elementChildren(node)
  let lastRank = -1
  let lastTag = null
  for (const child of children) {
    const tag = nodeName(child)
    const r = rank.get(tag)
    if (r === undefined) continue
    if (r < lastRank) {
      issues.push(
        issue(
          'error',
          'element-order',
          partPath,
          `<${nodeName(node)}> has <${tag}> after <${lastTag}>, violating the ${schemaLabel} child order.`,
        ),
      )
    }
    lastRank = Math.max(lastRank, r)
    lastTag = tag
  }
}

const CT_WORKSHEET_ORDER = [
  'sheetPr',
  'dimension',
  'sheetViews',
  'sheetFormatPr',
  'cols',
  'sheetData',
  'sheetCalcPr',
  'sheetProtection',
  'protectedRanges',
  'scenarios',
  'autoFilter',
  'sortState',
  'dataConsolidate',
  'customSheetViews',
  'mergeCells',
  'phoneticPr',
  'conditionalFormatting',
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
]

const TX_BODY_ORDER = ['a:bodyPr', 'a:lstStyle', 'a:p']
const TBL_ORDER = ['w:tblPr', 'w:tblGrid', 'w:tr']

function checkElementOrder(rootNodes, partPath, issues) {
  for (const p of collectByTag(rootNodes, 'w:p')) {
    assertFirstIfPresent(p, 'w:pPr', partPath, issues, 'w:pPr must precede run content in w:p')
  }
  for (const tbl of collectByTag(rootNodes, 'w:tbl')) {
    const hasGrid = elementChildren(tbl).some((child) => nodeName(child) === 'w:tblGrid')
    if (!hasGrid) {
      issues.push(issue('error', 'missing-tblgrid', partPath, '<w:tbl> has no required <w:tblGrid> child.'))
    }
    assertCanonicalOrder(tbl, TBL_ORDER, partPath, issues, 'CT_Tbl')
  }
  for (const txBody of collectByTag(rootNodes, 'p:txBody')) {
    assertCanonicalOrder(txBody, TX_BODY_ORDER, partPath, issues, 'CT_TextBody')
  }
  for (const worksheet of collectByTag(rootNodes, 'worksheet')) {
    assertCanonicalOrder(worksheet, CT_WORKSHEET_ORDER, partPath, issues, 'CT_Worksheet')
  }
}

// ---------------------------------------------------------------------------
// Zip container checks
// ---------------------------------------------------------------------------

function checkZipEntryNames(entries, issues) {
  const seen = new Map()
  for (const entry of entries) {
    if (entry.name.startsWith('/')) {
      issues.push(issue('error', 'zip-leading-slash', entry.name, 'Zip entry name starts with "/".'))
    }
    if (entry.name.includes('\\')) {
      issues.push(issue('error', 'zip-backslash', entry.name, 'Zip entry name contains a backslash; OPC/ODF require forward slashes.'))
    }
    if (entry.name.split('/').includes('..')) {
      issues.push(issue('error', 'zip-path-traversal', entry.name, 'Zip entry name contains a ".." path segment.'))
    }
    if (!entry.isDirectory) {
      seen.set(entry.name, (seen.get(entry.name) ?? 0) + 1)
    }
  }
  for (const [name, count] of seen) {
    if (count > 1) issues.push(issue('error', 'zip-duplicate-entry', name, `Zip contains ${count} entries with the same name.`))
  }
}

function checkOdfMimetypeEntry(entries, issues) {
  const mimetype = entries.find((e) => e.name === 'mimetype')
  if (!mimetype) return // caller only invokes this once ODF-ness is already established
  const first = entries.find((e) => !e.isDirectory)
  if (first?.name !== 'mimetype') {
    issues.push(issue('error', 'odf-mimetype-not-first', 'mimetype', `The "mimetype" entry must be the first file in the zip; found "${first?.name}" first.`))
  }
  if (mimetype.method !== 0) {
    issues.push(issue('error', 'odf-mimetype-not-stored', 'mimetype', `The "mimetype" entry must be stored (uncompressed); found compression method ${mimetype.method}.`))
  }
}

// ---------------------------------------------------------------------------
// OPC content-types / relationships
// ---------------------------------------------------------------------------

function parseSimpleXml(xmlText) {
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })
  return parser.parse(xmlText)
}

function asList(value) {
  if (value === undefined) return []
  return Array.isArray(value) ? value : [value]
}

function checkContentTypes(files, issues) {
  const xml = files.get('[Content_Types].xml')
  if (xml === undefined) {
    issues.push(issue('error', 'missing-content-types', '[Content_Types].xml', 'Package is missing [Content_Types].xml.'))
    return { defaults: new Set(), overrides: new Set() }
  }
  const parsed = parseOrdered(xml.toString('utf8'))
  if (!parsed.ok) {
    issues.push(issue('error', 'malformed-xml', '[Content_Types].xml', `Not well-formed: ${parsed.message}`))
    return { defaults: new Set(), overrides: new Set() }
  }

  const tree = parseSimpleXml(xml.toString('utf8'))
  const defaultsSeen = new Map()
  const overridesSeen = new Map()
  for (const d of asList(tree.Types?.Default)) {
    const ext = String(d['@_Extension'] ?? '').toLowerCase()
    defaultsSeen.set(ext, (defaultsSeen.get(ext) ?? 0) + 1)
  }
  for (const o of asList(tree.Types?.Override)) {
    const partName = String(o['@_PartName'] ?? '')
    overridesSeen.set(partName, (overridesSeen.get(partName) ?? 0) + 1)
  }
  for (const [ext, count] of defaultsSeen) {
    if (count > 1) issues.push(issue('error', 'duplicate-content-type-default', '[Content_Types].xml', `Duplicate <Default> entry for extension "${ext}".`))
  }
  for (const [partName, count] of overridesSeen) {
    if (count > 1) issues.push(issue('error', 'duplicate-content-type-override', '[Content_Types].xml', `Duplicate <Override> entry for part "${partName}".`))
  }

  return { defaults: new Set(defaultsSeen.keys()), overrides: new Set(overridesSeen.keys()) }
}

function hasContentType(contentTypes, partPath) {
  if (contentTypes.overrides.has(`/${partPath}`)) return true
  const dot = partPath.lastIndexOf('.')
  if (dot === -1) return false
  return contentTypes.defaults.has(partPath.slice(dot + 1).toLowerCase())
}

function checkEveryPartHasContentType(files, contentTypes, issues) {
  for (const partPath of files.keys()) {
    if (partPath === '[Content_Types].xml') continue
    if (!hasContentType(contentTypes, partPath)) {
      issues.push(issue('error', 'part-missing-content-type', partPath, `"${partPath}" has no Default or Override content-type entry.`))
    }
  }
}

const URL_SCHEME_PATTERN = /^[a-zA-Z][a-zA-Z\d+\-.]*:/

function resolveRelsBaseDir(relsPath) {
  const marker = '_rels/'
  const index = relsPath.lastIndexOf(marker)
  return index === -1 ? '' : relsPath.slice(0, index)
}

function normalizePartPath(p) {
  const segments = p.split('/')
  const resolved = []
  for (const segment of segments) {
    if (segment === '' || segment === '.') continue
    if (segment === '..') resolved.pop()
    else resolved.push(segment)
  }
  return resolved.join('/')
}

function checkRelationshipFiles(files, contentTypes, issues) {
  const relsByPart = new Map() // owning part path -> Set<relationship id>
  for (const [relsPath, bytes] of files) {
    if (!relsPath.endsWith('.rels')) continue
    const xmlText = bytes.toString('utf8')
    const parsed = parseOrdered(xmlText)
    if (!parsed.ok) {
      issues.push(issue('error', 'malformed-xml', relsPath, `Not well-formed: ${parsed.message}`))
      continue
    }
    const tree = parseSimpleXml(xmlText)
    const relationships = asList(tree.Relationships?.Relationship)
    const baseDir = resolveRelsBaseDir(relsPath)
    const ids = new Set()

    for (const rel of relationships) {
      const id = String(rel['@_Id'] ?? '')
      const target = String(rel['@_Target'] ?? '')
      const targetMode = rel['@_TargetMode']
      ids.add(id)
      if (targetMode === 'External' || URL_SCHEME_PATTERN.test(target)) continue

      const resolved = normalizePartPath(target.startsWith('/') ? target.slice(1) : `${baseDir}${target}`)
      if (!files.has(resolved)) {
        issues.push(issue('error', 'dangling-relationship-target', relsPath, `Relationship "${id}" targets "${resolved}", which is not in the package.`))
        continue
      }
      if (!hasContentType(contentTypes, resolved)) {
        issues.push(issue('error', 'relationship-target-missing-content-type', relsPath, `"${resolved}" is targeted by relationship "${id}" but has no content-type entry.`))
      }
    }

    // The owning part is `_rels/`'s sibling: `<baseDir><basename minus .rels>`.
    const basename = relsPath.slice(baseDir.length + '_rels/'.length, -'.rels'.length)
    relsByPart.set(`${baseDir}${basename}`, ids)
  }
  return relsByPart
}

const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const R_ID_ATTR_NAMES = ['id', 'embed', 'link', 'cs', 'dm', 'lo', 'qs', 'href', 'pict', 'topLeft']

/** Finds the namespace prefix a part binds to the `r:` relationships namespace (usually, but not necessarily, "r"). */
function relationshipsPrefixIn(xmlText) {
  const match = new RegExp(`xmlns:(\\w+)="${RELATIONSHIPS_NS.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`).exec(xmlText)
  return match ? match[1] : null
}

function checkDanglingRIdReferences(files, relsByPart, issues) {
  for (const [partPath, bytes] of files) {
    if (!partPath.endsWith('.xml') || partPath.endsWith('.rels') || partPath === '[Content_Types].xml') continue
    const xmlText = bytes.toString('utf8')
    const prefix = relationshipsPrefixIn(xmlText)
    if (!prefix) continue // part doesn't reference the relationships namespace at all

    const attrPattern = new RegExp(`\\b${prefix}:(${R_ID_ATTR_NAMES.join('|')})="([^"]*)"`, 'g')
    const referencedIds = new Set()
    let match
    while ((match = attrPattern.exec(xmlText)) !== null) {
      if (match[2] !== '') referencedIds.add(match[2])
    }
    if (referencedIds.size === 0) continue

    const declaredIds = relsByPart.get(partPath) ?? new Set()
    for (const id of referencedIds) {
      if (!declaredIds.has(id)) {
        issues.push(
          issue(
            'error',
            'dangling-rid',
            partPath,
            `References relationship "${id}" (via ${prefix}:*) that is not declared in its .rels file.`,
          ),
        )
      }
    }
  }
}

// ---------------------------------------------------------------------------
// ODF manifest
// ---------------------------------------------------------------------------

function checkOdfManifest(files, issues) {
  const manifestXml = files.get('META-INF/manifest.xml')
  if (manifestXml === undefined) {
    issues.push(issue('error', 'missing-manifest', 'META-INF/manifest.xml', 'ODF package is missing META-INF/manifest.xml.'))
    return
  }
  const xmlText = manifestXml.toString('utf8')
  const parsed = parseOrdered(xmlText)
  if (!parsed.ok) {
    issues.push(issue('error', 'malformed-xml', 'META-INF/manifest.xml', `Not well-formed: ${parsed.message}`))
    return
  }
  const tree = parseSimpleXml(xmlText)
  const entries = asList(tree['manifest:manifest']?.['manifest:file-entry'])
  const listedPaths = new Set()
  for (const entry of entries) {
    const fullPath = String(entry['@_manifest:full-path'] ?? '')
    if (fullPath !== '/') listedPaths.add(fullPath)
  }

  for (const [partPath] of files) {
    if (partPath === 'mimetype' || partPath === 'META-INF/manifest.xml') continue
    if (!listedPaths.has(partPath)) {
      issues.push(issue('error', 'odf-file-not-in-manifest', partPath, `"${partPath}" is in the package but not listed in META-INF/manifest.xml.`))
    }
  }
  for (const listed of listedPaths) {
    if (!files.has(listed)) {
      issues.push(issue('error', 'odf-manifest-dangling-entry', 'META-INF/manifest.xml', `manifest.xml lists "${listed}", which is not in the package.`))
    }
  }
  const hasRoot = entries.some((entry) => entry['@_manifest:full-path'] === '/')
  if (!hasRoot) {
    issues.push(issue('error', 'odf-manifest-missing-root', 'META-INF/manifest.xml', 'manifest.xml has no root file-entry (manifest:full-path="/").'))
  }
}

// ---------------------------------------------------------------------------
// Format detection
// ---------------------------------------------------------------------------

const ODF_MIME_TO_KIND = {
  'application/vnd.oasis.opendocument.text': 'odt',
  'application/vnd.oasis.opendocument.presentation': 'odp',
  'application/vnd.oasis.opendocument.spreadsheet': 'ods',
}

const OFFICE_DOCUMENT_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'

/**
 * The OPC "start part" for a package is whatever `_rels/.rels` names via an
 * `officeDocument` relationship (ECMA-376 Part 2, §10.2) — determining kind
 * this way, rather than by checking whether `word/document.xml` etc.
 * physically exists, means a package MISSING its main part still gets
 * classified correctly (so `REQUIRED_PARTS` below can flag it as missing,
 * instead of silently falling through to "unknown" and skipping that check
 * entirely).
 */
function startPartKindFromRels(files) {
  const relsXml = files.get('_rels/.rels')
  if (relsXml === undefined) return null
  let tree
  try {
    tree = parseSimpleXml(relsXml.toString('utf8'))
  } catch {
    return null
  }
  const relationships = asList(tree.Relationships?.Relationship)
  const officeDoc = relationships.find((rel) => rel['@_Type'] === OFFICE_DOCUMENT_REL_TYPE)
  const target = officeDoc ? String(officeDoc['@_Target'] ?? '') : ''
  // Exact-match the well-known main-part names rather than just prefix-
  // matching the folder: `xl/workbook.bin` (the binary XLSB variant) starts
  // with `xl/` same as `xl/workbook.xml` does, but requires a different
  // `REQUIRED_PARTS` entry (no `xl/workbook.xml` ever exists in an XLSB
  // package) — conflating the two produced a false `missing-required-part`
  // on every XLSB save.
  if (target === 'word/document.xml') return 'docx'
  if (target === 'ppt/presentation.xml') return 'pptx'
  if (target === 'xl/workbook.xml') return 'xlsx'
  if (target === 'xl/workbook.bin') return 'xlsb'
  return null
}

function detectFormat(files) {
  const mimetype = files.get('mimetype')
  if (mimetype !== undefined) {
    const content = mimetype.toString('utf8').trim()
    return { family: 'odf', kind: ODF_MIME_TO_KIND[content] ?? 'odf-unknown', mimetype: content }
  }
  if (files.has('[Content_Types].xml')) {
    const fromRels = startPartKindFromRels(files)
    if (fromRels) return { family: 'opc', kind: fromRels }
    if (files.has('word/document.xml')) return { family: 'opc', kind: 'docx' }
    if (files.has('ppt/presentation.xml')) return { family: 'opc', kind: 'pptx' }
    if (files.has('xl/workbook.xml')) return { family: 'opc', kind: 'xlsx' }
    return { family: 'opc', kind: 'opc-unknown' }
  }
  return { family: 'unknown', kind: 'unknown' }
}

const REQUIRED_PARTS = {
  docx: ['word/document.xml', '_rels/.rels'],
  pptx: ['ppt/presentation.xml', '_rels/.rels'],
  xlsx: ['xl/workbook.xml', '_rels/.rels'],
  xlsb: ['xl/workbook.bin', '_rels/.rels'],
  odt: ['content.xml', 'META-INF/manifest.xml', 'mimetype'],
  odp: ['content.xml', 'META-INF/manifest.xml', 'mimetype'],
  ods: ['content.xml', 'META-INF/manifest.xml', 'mimetype'],
}

// ---------------------------------------------------------------------------
// Top-level entry point
// ---------------------------------------------------------------------------

/**
 * @param {Buffer} buffer
 * @returns {{ format: { family: string, kind: string }, issues: Issue[] }}
 */
export function validateOfficeFile(buffer) {
  const issues = []
  let entries, files
  try {
    ;({ entries, files } = readZipArchive(buffer))
  } catch (cause) {
    issues.push(issue('error', 'not-a-zip', null, cause instanceof Error ? cause.message : String(cause)))
    return { format: { family: 'unknown', kind: 'unknown' }, issues }
  }

  checkZipEntryNames(entries, issues)

  const format = detectFormat(files)

  for (const requiredPart of REQUIRED_PARTS[format.kind] ?? []) {
    if (!files.has(requiredPart)) {
      issues.push(issue('error', 'missing-required-part', requiredPart, `Required part "${requiredPart}" is missing for a .${format.kind} package.`))
    }
  }

  // Every XML/rels part: well-formed, no illegal characters, namespaces declared.
  const orderedTrees = new Map()
  for (const [partPath, bytes] of files) {
    if (!/\.(xml|rels)$/i.test(partPath)) continue
    const xmlText = bytes.toString('utf8')

    const illegal = findIllegalXmlChars(xmlText)
    if (illegal.length > 0) {
      const sample = illegal.map((c) => `U+${c.codePoint.toString(16).toUpperCase().padStart(4, '0')}@${c.index}`).join(', ')
      issues.push(issue('error', 'illegal-xml-char', partPath, `Contains character(s) XML 1.0 forbids: ${sample}.`))
    }

    const parsed = parseOrdered(xmlText)
    if (!parsed.ok) {
      issues.push(issue('error', 'malformed-xml', partPath, `Not well-formed: ${parsed.message}`))
      continue
    }
    orderedTrees.set(partPath, parsed.tree)
    checkNamespacesDeclared(parsed.tree, new Set(), partPath, issues)
    checkElementOrder(parsed.tree, partPath, issues)
  }

  if (format.family === 'odf') {
    checkOdfMimetypeEntry(entries, issues)
    checkOdfManifest(files, issues)
  } else if (format.family === 'opc') {
    const contentTypes = checkContentTypes(files, issues)
    checkEveryPartHasContentType(files, contentTypes, issues)
    const relsByPart = checkRelationshipFiles(files, contentTypes, issues)
    checkDanglingRIdReferences(files, relsByPart, issues)
  }

  return { format, issues }
}

export function formatIssuesReport(filePath, result) {
  const lines = [`${filePath}  (${result.format.family}/${result.format.kind})`]
  if (result.issues.length === 0) {
    lines.push('  OK — no violations found.')
    return lines.join('\n')
  }
  for (const problem of result.issues) {
    lines.push(`  [${problem.severity}] ${problem.code} ${problem.part ? `(${problem.part})` : ''}: ${problem.message}`)
  }
  return lines.join('\n')
}
