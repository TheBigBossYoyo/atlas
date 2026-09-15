/**
 * Atlas — hyperlink insertion helper (D18/DXE-06)
 *
 * Hyperlink insertion needs a relationship (`word/_rels/document.xml.rels`,
 * `TargetMode="External"`) alongside the document edit — bundle-level state
 * the pure Command pipeline doesn't own, the same reason `insertImage.ts`
 * allocates its own relationship id before dispatching a command. This
 * module allocates that relationship id (avoiding collisions with whatever
 * the source document already defines) and applies the resulting
 * `insert-hyperlink` command.
 */

import type { DocxBundle } from '../index'
import type { Document } from '../model/document'
import type { Relationship } from '../parser/relationships'

import { applyCommand } from './commands'
import type { Command, Range } from './commandTypes'
import { allocateRelationshipId } from './relationshipIds'

const HYPERLINK_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'

export interface InsertHyperlinkResult {
  readonly bundle: DocxBundle
  readonly document: Document
  readonly range: Range
  readonly inverse: Command
}

export function insertHyperlinkIntoBundle(
  bundle: DocxBundle,
  range: Range,
  url: string,
): InsertHyperlinkResult {
  const relationships = bundle.relationships ?? []
  const relationshipId = allocateRelationshipId(relationships)

  const nextRelationships: ReadonlyArray<Relationship> = [
    ...relationships,
    { id: relationshipId, type: HYPERLINK_REL_TYPE, target: url, targetMode: 'External' },
  ]

  const result = applyCommand(bundle.document, {
    kind: 'insert-hyperlink',
    range,
    url,
    relationshipId,
  })
  // `insert-hyperlink` always computes a resulting range; the fallback only
  // exists to satisfy applyCommand's general (range is optional for most
  // command kinds) return type.
  const nextRange = result.range ?? range

  return {
    bundle: { ...bundle, document: result.document, relationships: nextRelationships },
    document: result.document,
    range: nextRange,
    inverse: result.inverse,
  }
}
