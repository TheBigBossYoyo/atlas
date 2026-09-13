/**
 * Atlas — DOCX [Content_Types].xml writer (Wave D.3)
 *
 * Inverse of `src/docx/parser/contentTypes.ts`.
 * Emits OOXML `[Content_Types].xml` and provides helpers for managing entries.
 */

import { XMLBuilder } from 'fast-xml-parser'
import type { ContentTypes as ContentTypesPart } from '../parser/contentTypes'

export type { ContentTypes as ContentTypesPart } from '../parser/contentTypes'

// ---------------------------------------------------------------------------
// Standard MIME map
// ---------------------------------------------------------------------------

export const MEDIA_CONTENT_TYPES: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tiff: 'image/tiff',
  svg: 'image/svg+xml',
}

// ---------------------------------------------------------------------------
// XMLBuilder instance
// ---------------------------------------------------------------------------

const xmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  format: false,
  suppressEmptyNode: true,
})

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Emits the XML string for `[Content_Types].xml`.
 * `Default` entries are emitted before `Override` entries per OOXML spec.
 */
export function writeContentTypesXml(types: ContentTypesPart): string {
  const defaults = types.defaults.map((d) => ({
    '@_Extension': d.extension,
    '@_ContentType': d.contentType,
  }))

  const overrides = types.overrides.map((o) => ({
    '@_PartName': o.partName,
    '@_ContentType': o.contentType,
  }))

  const typesNode: Record<string, unknown> = {
    '@_xmlns': CT_NS,
  }
  if (defaults.length > 0) typesNode['Default'] = defaults
  if (overrides.length > 0) typesNode['Override'] = overrides

  const obj = {
    '?xml': { '@_version': '1.0', '@_encoding': 'UTF-8', '@_standalone': 'yes' },
    Types: typesNode,
  }

  const built = xmlBuilder.build(obj) as string
  if (built.startsWith('<?xml')) {
    return built
  }
  return XML_DECL + built
}

/**
 * Pure helper: ensures a `Default` entry for `extension` is present.
 * Uses the standard MIME map for known image extensions.
 * If the extension already has a Default entry, returns the same shape unchanged.
 */
export function ensureMediaContentType(
  types: ContentTypesPart,
  extension: string,
): ContentTypesPart {
  const ext = extension.toLowerCase().replace(/^\./, '')
  const already = types.defaults.some(
    (d) => d.extension.toLowerCase() === ext,
  )
  if (already) return types

  const contentType = MEDIA_CONTENT_TYPES[ext]
  if (contentType === undefined) return types

  return {
    ...types,
    defaults: [...types.defaults, { extension: ext, contentType }],
  }
}

/**
 * Pure helper: adds an Override entry for `partName`.
 */
export function addOverride(
  types: ContentTypesPart,
  partName: string,
  contentType: string,
): ContentTypesPart {
  return {
    ...types,
    overrides: [...types.overrides, { partName, contentType }],
  }
}
