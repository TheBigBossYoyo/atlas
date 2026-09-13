/**
 * Atlas — DOCX relationships writer (Wave D.3)
 *
 * Inverse of `src/docx/parser/relationships.ts`.
 * Emits OOXML `.rels` XML and provides helpers for ID allocation.
 */

import { XMLBuilder } from 'fast-xml-parser'
import type { Relationship } from '../parser/relationships'

// ---------------------------------------------------------------------------
// XMLBuilder instance (shared — stateless after construction)
// ---------------------------------------------------------------------------

const xmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  format: false,
  suppressEmptyNode: true,
})

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Emits the XML string for a `.rels` file.
 */
export function writeRelationshipsXml(rels: ReadonlyArray<Relationship>): string {
  const relationships = rels.map((r) => {
    const attrs: Record<string, string> = {
      '@_Id': r.id,
      '@_Type': r.type,
      '@_Target': r.target,
    }
    if (r.targetMode !== undefined) {
      attrs['@_TargetMode'] = r.targetMode
    }
    return attrs
  })

  const obj = {
    '?xml': { '@_version': '1.0', '@_encoding': 'UTF-8', '@_standalone': 'yes' },
    Relationships: {
      '@_xmlns': RELS_NS,
      Relationship: relationships,
    },
  }

  const built = xmlBuilder.build(obj) as string
  // fast-xml-parser includes the xml declaration when the ?xml key is present;
  // ensure we don't double-emit it.
  if (built.startsWith('<?xml')) {
    return built
  }
  return XML_DECL + built
}

/**
 * Returns the next free `rIdN` string given an existing collection.
 * If the collection is empty, returns `"rId1"`.
 */
export function allocateRelationshipId(existing: ReadonlyArray<Relationship>): string {
  if (existing.length === 0) return 'rId1'

  let max = 0
  for (const rel of existing) {
    const match = /^rId(\d+)$/i.exec(rel.id)
    if (match) {
      const n = parseInt(match[1], 10)
      if (n > max) max = n
    }
  }
  return `rId${max + 1}`
}

/**
 * Pure helper: adds a new relationship to the array, returning the updated
 * array and the allocated ID.
 */
export function addRelationship(
  rels: ReadonlyArray<Relationship>,
  type: string,
  target: string,
  mode?: 'External' | 'Internal',
): { id: string; rels: ReadonlyArray<Relationship> } {
  const id = allocateRelationshipId(rels)
  const newRel: Relationship = {
    id,
    type,
    target,
    ...(mode !== undefined ? { targetMode: mode } : {}),
  }
  return { id, rels: [...rels, newRel] }
}
