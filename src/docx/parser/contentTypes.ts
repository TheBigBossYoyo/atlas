/**
 * Atlas — DOCX [Content_Types].xml parser (Wave A.1)
 *
 * Parses the `[Content_Types].xml` file that lives at the root of every
 * OOXML package.  Returns an immutable `ContentTypes` object containing
 * the `Default` and `Override` entries.
 */

import { XMLParser } from 'fast-xml-parser'

import { assertXmlPartSizeWithinLimit } from './xmlSizeGuard'

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

export interface ContentTypeDefault {
  readonly extension: string
  readonly contentType: string
}

export interface ContentTypeOverride {
  readonly partName: string
  readonly contentType: string
}

export interface ContentTypes {
  readonly defaults: ReadonlyArray<ContentTypeDefault>
  readonly overrides: ReadonlyArray<ContentTypeOverride>
}

// ---------------------------------------------------------------------------
// Internal XML shape
// ---------------------------------------------------------------------------

interface RawDefault {
  '@_Extension': string
  '@_ContentType': string
}

interface RawOverride {
  '@_PartName': string
  '@_ContentType': string
}

interface RawTypes {
  Types?: {
    Default?: RawDefault | RawDefault[]
    Override?: RawOverride | RawOverride[]
  }
}

// ---------------------------------------------------------------------------
// Parser instance
// ---------------------------------------------------------------------------

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

// ---------------------------------------------------------------------------
// Implementation
// ---------------------------------------------------------------------------

/**
 * Parses the XML content of `[Content_Types].xml`.
 *
 * @param xml - UTF-8 text content of `[Content_Types].xml`.
 * @returns     An immutable `ContentTypes` object.
 */
export function parseContentTypes(xml: string): ContentTypes {
  assertXmlPartSizeWithinLimit(xml, '[Content_Types].xml')
  const raw = xmlParser.parse(xml) as RawTypes

  const rawTypes = raw?.Types

  return {
    defaults: toArray(rawTypes?.Default).map(
      (d): ContentTypeDefault => ({
        extension: d['@_Extension'] ?? '',
        contentType: d['@_ContentType'] ?? '',
      }),
    ),
    overrides: toArray(rawTypes?.Override).map(
      (o): ContentTypeOverride => ({
        partName: o['@_PartName'] ?? '',
        contentType: o['@_ContentType'] ?? '',
      }),
    ),
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function toArray<T>(value: T | T[] | undefined | null): T[] {
  if (value === undefined || value === null) return []
  return Array.isArray(value) ? value : [value]
}
