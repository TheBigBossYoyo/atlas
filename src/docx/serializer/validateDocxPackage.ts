/**
 * Atlas — post-serialization DOCX package validation (D20 / DXS-19)
 *
 * `saveDocx` builds every part of a `.docx` package in memory before handing
 * it to `packDocx`/JSZip. Nothing previously checked that the XML it just
 * generated was actually well-formed, that every `w:tbl` carried its
 * schema-required `w:tblGrid` (the exact regression P1.5 fixed), or that
 * every relationship a part declares actually resolves to a part that both
 * exists in the package and has a content-type entry (the exact regression
 * P1.7 fixed for comments/images). A future change could silently
 * reintroduce any of those bugs; this module is the safety net that turns
 * that into a loud, specific save failure instead of a Word repair dialog.
 *
 * This is a self-check on Atlas's OWN freshly generated output, not a parse
 * of untrusted input — failures throw `DocxSaveError`, not `DocxParseError`.
 */

import { XMLParser } from 'fast-xml-parser'

import { parseContentTypes, type ContentTypes } from '../parser/contentTypes'
import { parseRelationships, type Relationship } from '../parser/relationships'

/** Thrown when a freshly generated DOCX package fails a post-save sanity check. */
export class DocxSaveError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'DocxSaveError'
    Object.setPrototypeOf(this, new.target.prototype)
  }
}

interface OrderedXmlNode {
  readonly [key: string]: unknown
}

const orderedXmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  trimValues: false,
})

/**
 * Validates every part about to be packed into a `.docx` archive.
 *
 * @param parts - The exact path → content map `saveDocx` is about to hand to
 *   `packDocx` (both freshly written parts and passed-through raw bytes).
 * @throws `DocxSaveError` naming the offending part and the specific
 *   invariant it violated.
 */
export function validateDocxPackage(parts: ReadonlyMap<string, string | Uint8Array>): void {
  const parsedXmlParts = new Map<string, ReadonlyArray<OrderedXmlNode>>()

  for (const [path, content] of parts) {
    if (typeof content !== 'string' || !path.toLowerCase().endsWith('.xml')) {
      continue
    }

    let parsed: OrderedXmlNode[]
    try {
      parsed = orderedXmlParser.parse(content) as OrderedXmlNode[]
    } catch (cause) {
      const msg = cause instanceof Error ? cause.message : String(cause)
      throw new DocxSaveError(`Generated "${path}" is not well-formed XML: ${msg}`)
    }

    parsedXmlParts.set(path, parsed)
    assertNoTableMissingGrid(parsed, path)
  }

  validateRelationshipTargets(parts)
}

// ---------------------------------------------------------------------------
// Invariant: every w:tbl carries a w:tblGrid (DXS-02)
// ---------------------------------------------------------------------------

function assertNoTableMissingGrid(nodes: ReadonlyArray<OrderedXmlNode>, partPath: string): void {
  for (const node of nodes) {
    const name = nodeName(node)
    if (name === undefined) {
      continue
    }

    const children = node[name]
    if (!Array.isArray(children)) {
      continue
    }

    if (name === 'w:tbl') {
      const hasGrid = (children as OrderedXmlNode[]).some((child) => nodeName(child) === 'w:tblGrid')
      if (!hasGrid) {
        throw new DocxSaveError(
          `Generated "${partPath}" contains a <w:tbl> with no required <w:tblGrid> child.`,
        )
      }
    }

    assertNoTableMissingGrid(children as OrderedXmlNode[], partPath)
  }
}

function nodeName(node: OrderedXmlNode): string | undefined {
  for (const key of Object.keys(node)) {
    if (key !== ':@') {
      return key
    }
  }
  return undefined
}

// ---------------------------------------------------------------------------
// Invariant: every relationship target exists and has a content type (DXS-06/07)
// ---------------------------------------------------------------------------

function validateRelationshipTargets(parts: ReadonlyMap<string, string | Uint8Array>): void {
  const contentTypesXml = parts.get('[Content_Types].xml')
  const contentTypes =
    typeof contentTypesXml === 'string' ? parseContentTypes(contentTypesXml) : undefined

  for (const [relsPath, content] of parts) {
    if (typeof content !== 'string' || !relsPath.endsWith('.rels')) {
      continue
    }

    const baseDir = resolveRelsBaseDir(relsPath)
    let relationships: ReadonlyArray<Relationship>
    try {
      relationships = parseRelationships(content)
    } catch (cause) {
      const msg = cause instanceof Error ? cause.message : String(cause)
      throw new DocxSaveError(`Generated "${relsPath}" is not well-formed XML: ${msg}`)
    }

    for (const rel of relationships) {
      // Real documents always set TargetMode="External" on external
      // relationships (hyperlinks, etc.), but be defensive about a
      // URL-shaped target that omits it rather than risk rejecting a
      // legitimate save over a part that was never meant to exist locally.
      if (rel.targetMode === 'External' || isExternalUrl(rel.target)) {
        continue
      }

      const resolvedPath = normalizePartPath(`${baseDir}${rel.target}`)
      if (!parts.has(resolvedPath)) {
        throw new DocxSaveError(
          `"${relsPath}" relationship "${rel.id}" targets "${resolvedPath}", but that part was not written to the package.`,
        )
      }

      if (contentTypes !== undefined && !hasContentType(contentTypes, resolvedPath)) {
        throw new DocxSaveError(
          `"${resolvedPath}" is referenced by relationship "${rel.id}" in "${relsPath}" but has no content-type entry in [Content_Types].xml.`,
        )
      }
    }
  }
}

const URL_SCHEME_PATTERN = /^[a-zA-Z][a-zA-Z\d+\-.]*:/

function isExternalUrl(target: string): boolean {
  return URL_SCHEME_PATTERN.test(target)
}

function resolveRelsBaseDir(relsPath: string): string {
  const marker = '_rels/'
  const index = relsPath.lastIndexOf(marker)
  return index === -1 ? '' : relsPath.slice(0, index)
}

function normalizePartPath(path: string): string {
  const segments = path.split('/')
  const resolved: string[] = []

  for (const segment of segments) {
    if (segment === '' || segment === '.') {
      continue
    }
    if (segment === '..') {
      resolved.pop()
      continue
    }
    resolved.push(segment)
  }

  return resolved.join('/')
}

function hasContentType(contentTypes: ContentTypes, partPath: string): boolean {
  const partName = `/${partPath}`
  if (contentTypes.overrides.some((override) => override.partName === partName)) {
    return true
  }

  const extension = partPath.includes('.') ? (partPath.split('.').pop() ?? '') : ''
  if (extension.length === 0) {
    return false
  }

  return contentTypes.defaults.some((entry) => entry.extension.toLowerCase() === extension.toLowerCase())
}
