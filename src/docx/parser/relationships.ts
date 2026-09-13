/**
 * Atlas — DOCX relationships parser (Wave A.1)
 *
 * Parses OOXML `.rels` files (both package-level `_rels/.rels` and
 * part-level e.g. `word/_rels/document.xml.rels`).  The XML schema is
 * identical in both cases, so one parser handles all call sites.
 */

import { XMLParser } from 'fast-xml-parser'

import { assertXmlPartSizeWithinLimit } from './xmlSizeGuard'

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/**
 * A single relationship entry from a `.rels` file.
 */
export interface Relationship {
  readonly id: string
  readonly type: string
  readonly target: string
  readonly targetMode?: 'External' | 'Internal'
}

// ---------------------------------------------------------------------------
// Internal XML shape (what fast-xml-parser returns)
// ---------------------------------------------------------------------------

interface RawRelationship {
  '@_Id': string
  '@_Type': string
  '@_Target': string
  '@_TargetMode'?: string
}

interface RawRelationships {
  Relationships?: {
    Relationship?: RawRelationship | RawRelationship[]
  }
}

// ---------------------------------------------------------------------------
// Parser instance (shared — stateless after construction)
// ---------------------------------------------------------------------------

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

// ---------------------------------------------------------------------------
// Implementation
// ---------------------------------------------------------------------------

/**
 * Parses the XML content of any OOXML `.rels` file.
 *
 * @param xml - UTF-8 text content of the `.rels` file.
 * @returns     An immutable array of `Relationship` objects.
 */
export function parseRelationships(xml: string): ReadonlyArray<Relationship> {
  assertXmlPartSizeWithinLimit(xml, '*.rels')
  const raw = xmlParser.parse(xml) as RawRelationships

  const rawRels = raw?.Relationships?.Relationship

  if (rawRels === undefined || rawRels === null) {
    return []
  }

  const items: RawRelationship[] = Array.isArray(rawRels) ? rawRels : [rawRels]

  return items.map((item): Relationship => {
    const targetMode = normaliseTargetMode(item['@_TargetMode'])
    const rel: Relationship = {
      id: item['@_Id'] ?? '',
      type: item['@_Type'] ?? '',
      target: item['@_Target'] ?? '',
      ...(targetMode !== undefined ? { targetMode } : {}),
    }
    return rel
  })
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function normaliseTargetMode(
  raw: string | undefined,
): 'External' | 'Internal' | undefined {
  if (raw === 'External') return 'External'
  if (raw === 'Internal') return 'Internal'
  return undefined
}
