/**
 * Atlas — DOCX `word/fontTable.xml` parser (DEFER-4 / DXP-13)
 *
 * `word/fontTable.xml` lists every font Word knows the document references,
 * and — when the author saved with "Embed fonts in the file" — up to four
 * `w:embedRegular`/`w:embedBold`/`w:embedItalic`/`w:embedBoldItalic`
 * children per `<w:font>`, each pointing (via `r:id`) at an obfuscated font
 * part in `word/fonts/*.fntdata` and carrying the `w:fontKey` GUID needed to
 * de-obfuscate it (see `./deobfuscate.ts`).
 *
 * This module only parses the font-table XML itself into a plain structure;
 * resolving `r:id` against `word/_rels/fontTable.xml.rels` and reading the
 * actual font part bytes is `./embedded.ts`'s job, so this stays a small,
 * independently testable unit like the rest of this codebase's per-part
 * parsers (`relationships.ts`, `headers.ts`, etc.).
 */

import { XMLParser } from 'fast-xml-parser'

import { DocxParseError } from '../parser/unzip'
import { assertXmlPartSizeWithinLimit } from '../parser/xmlSizeGuard'

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/** One `w:embedRegular`/`w:embedBold`/`w:embedItalic`/`w:embedBoldItalic` reference. */
export interface EmbeddedFontRef {
  /** Relationship id (`r:id`) resolving to the obfuscated font part. */
  readonly relId: string
  /**
   * The GUID (`w:fontKey`) used to de-obfuscate the font part's first 32
   * bytes. Absent only for a malformed document — Word always writes this
   * alongside an embed element for a font it obfuscated.
   */
  readonly fontKey?: string
  /** `w:subsetted="1"` — the embedded font contains only glyphs the document actually uses. */
  readonly subsetted?: boolean
}

/** One `<w:font>` entry from `word/fontTable.xml`. */
export interface FontTableEntry {
  readonly name: string
  readonly embedRegular?: EmbeddedFontRef
  readonly embedBold?: EmbeddedFontRef
  readonly embedItalic?: EmbeddedFontRef
  readonly embedBoldItalic?: EmbeddedFontRef
}

// ---------------------------------------------------------------------------
// Internal XML shape
// ---------------------------------------------------------------------------

type XmlScalar = string | number | boolean
type XmlValue = XmlScalar | XmlNode | XmlValue[]

interface XmlNode {
  readonly [key: string]: XmlValue | undefined
}

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Parses the XML content of `word/fontTable.xml`.
 *
 * @param xml - UTF-8 text of the fontTable part.
 * @returns     Every `<w:font>` entry, in document order.
 * @throws      `DocxParseError` on malformed XML.
 */
export function parseFontTable(xml: string): ReadonlyArray<FontTableEntry> {
  assertXmlPartSizeWithinLimit(xml, 'word/fontTable.xml')

  let raw: XmlNode
  try {
    raw = xmlParser.parse(xml) as XmlNode
  } catch (cause) {
    const msg = cause instanceof Error ? cause.message : String(cause)
    throw new DocxParseError(`Failed to parse fontTable XML: ${msg}`, 'word/fontTable.xml')
  }

  const root = asXmlNode(raw['w:fonts'])
  const fontNodes = getNodes(root, 'w:font')

  return fontNodes
    .map(parseFontEntry)
    .filter((entry): entry is FontTableEntry => entry !== undefined)
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

function parseFontEntry(node: XmlNode): FontTableEntry | undefined {
  const name = getAttr(node, 'w:name')
  if (name === undefined || name === '') {
    return undefined
  }

  const embedRegular = parseEmbedRef(getNode(node, 'w:embedRegular'))
  const embedBold = parseEmbedRef(getNode(node, 'w:embedBold'))
  const embedItalic = parseEmbedRef(getNode(node, 'w:embedItalic'))
  const embedBoldItalic = parseEmbedRef(getNode(node, 'w:embedBoldItalic'))

  return {
    name,
    ...(embedRegular !== undefined ? { embedRegular } : {}),
    ...(embedBold !== undefined ? { embedBold } : {}),
    ...(embedItalic !== undefined ? { embedItalic } : {}),
    ...(embedBoldItalic !== undefined ? { embedBoldItalic } : {}),
  }
}

function parseEmbedRef(node: XmlNode | undefined): EmbeddedFontRef | undefined {
  if (node === undefined) {
    return undefined
  }

  const relId = getAttr(node, 'r:id')
  if (relId === undefined || relId === '') {
    return undefined
  }

  const fontKey = getAttr(node, 'w:fontKey')
  const subsettedRaw = getAttr(node, 'w:subsetted')

  return {
    relId,
    ...(fontKey !== undefined && fontKey !== '' ? { fontKey } : {}),
    ...(subsettedRaw !== undefined ? { subsetted: subsettedRaw === '1' || subsettedRaw === 'true' } : {}),
  }
}

function getNodes(parent: XmlNode | undefined, key: string): XmlNode[] {
  return toArray(parent?.[key]).filter((value): value is XmlNode => isXmlNode(value))
}

function getNode(parent: XmlNode | undefined, key: string): XmlNode | undefined {
  return asXmlNode(parent?.[key])
}

function getAttr(node: XmlNode | undefined, key: string): string | undefined {
  if (node === undefined) return undefined
  return toOptionalString(node[`@_${key}`])
}

function toArray<T>(value: T | T[] | undefined | null): T[] {
  if (value === undefined || value === null) return []
  return Array.isArray(value) ? value : [value]
}

function isXmlNode(value: XmlValue | unknown): value is XmlNode {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function asXmlNode(value: XmlValue | unknown): XmlNode | undefined {
  return isXmlNode(value) ? value : undefined
}

function toOptionalString(value: XmlValue | undefined): string | undefined {
  if (value === undefined || value === null) return undefined
  return String(value)
}
