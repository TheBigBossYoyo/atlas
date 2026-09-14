/**
 * Atlas — `word/settings.xml` parser (D17 / DXE-11)
 *
 * `word/settings.xml` passes through the save pipeline byte-for-byte
 * (`docx/index.ts`'s `rawArchive` passthrough) because Atlas's `Document`
 * model doesn't own most of what it contains. The one setting the editor
 * needs to read back is whether Track Changes was left switched on in the
 * source file (`<w:trackChanges/>`, using the OOXML on/off-element
 * convention: present means on unless `w:val` says otherwise) — that value
 * seeds the toolbar's initial toggle state instead of always starting from
 * `false`. Every other setting stays untouched.
 */
import { XMLParser } from 'fast-xml-parser'

import { DocxParseError } from './unzip'
import { assertXmlPartSizeWithinLimit } from './xmlSizeGuard'

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

export interface SettingsPart {
  readonly trackChanges: boolean
}

interface RawSettings {
  'w:settings'?: {
    'w:trackChanges'?: string | number | { readonly '@_w:val'?: string | number | boolean }
  }
}

function isOffValue(value: string | number | boolean): boolean {
  const normalized = String(value).trim().toLowerCase()
  return normalized === 'false' || normalized === '0' || normalized === 'off'
}

/**
 * Parses the XML content of `word/settings.xml`.
 *
 * @param xml - UTF-8 text of the settings part.
 * @returns    `{ trackChanges: boolean }` — `false` when the element is
 *   absent, matching a document that has never had Track Changes enabled.
 * @throws     `DocxParseError` on malformed XML.
 */
export function parseSettings(xml: string): SettingsPart {
  assertXmlPartSizeWithinLimit(xml, 'word/settings.xml')

  let parsed: RawSettings
  try {
    parsed = xmlParser.parse(xml) as RawSettings
  } catch (error) {
    throw new DocxParseError(
      `Failed to parse word/settings.xml: ${error instanceof Error ? error.message : String(error)}`,
      'word/settings.xml',
    )
  }

  const node = parsed['w:settings']?.['w:trackChanges']
  if (node === undefined) {
    return { trackChanges: false }
  }

  if (typeof node === 'object') {
    const val = node['@_w:val']
    return { trackChanges: val === undefined ? true : !isOffValue(val) }
  }

  return { trackChanges: true }
}
