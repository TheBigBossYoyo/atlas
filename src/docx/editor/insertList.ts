/**
 * Atlas — list-numbering allocation helper (D18/DXE-06)
 *
 * Turning a paragraph into a bulleted/numbered list needs a numbering
 * definition to exist somewhere the *saved* package can find it. Atlas keeps
 * two views of numbering:
 *   - `bundle.numberingPart` (abstractNums + nums) — the source of truth
 *     `saveDocx` actually serializes to `word/numbering.xml`.
 *   - `document.numbering` — a flattened `NumberingDef` map the parser builds
 *     once at load time for the renderer/editor's convenience.
 * A list toggle has to add an entry to *both* (or rendering picks it up but
 * save loses it, or vice versa), which is bundle-level state the pure
 * `insert-list` Command (paragraph numPr only) doesn't own — the same reason
 * image/hyperlink insertion each get their own bundle-aware helper.
 */

import type { DocxBundle } from '../index'
import type { AbstractNum, NumberingPart, NumInstance } from '../parser/numbering'
import type { LvlDef, NumberingDef } from '../model'

export type ListKind = 'bullet' | 'number'

/**
 * Ensures `bundle.numberingPart` and `bundle.document.numbering` both have an
 * entry for `numId`. If the source document already defines this numId
 * (either because it's a real document numbering id, or because a previous
 * insert-list call already created it), reuses it as-is rather than
 * overwriting — repeated toggles of the same list kind converge on one
 * shared definition instead of minting a fresh one each time.
 */
export function ensureListNumbering(bundle: DocxBundle, numId: number, kind: ListKind): DocxBundle {
  const numIdStr = String(numId)
  if (bundle.document.numbering.has(numIdStr)) {
    return bundle
  }

  const abstractNumId = `atlas-list-${numIdStr}`
  const level: LvlDef =
    kind === 'number'
      ? { level: 0, format: 'decimal', text: { value: '%1.', placeholders: [1] }, suffix: 'tab' }
      : { level: 0, format: 'bullet', text: { value: '•', placeholders: [] }, suffix: 'tab' }

  const abstractNum: AbstractNum = { abstractNumId, levels: new Map([[0, level]]) }
  const numInstance: NumInstance = { numId: numIdStr, abstractNumId }

  const previousPart = bundle.numberingPart ?? { abstractNums: new Map(), nums: new Map() }
  const nextNumberingPart: NumberingPart = {
    abstractNums: new Map(previousPart.abstractNums).set(abstractNumId, abstractNum),
    nums: new Map(previousPart.nums).set(numIdStr, numInstance),
  }

  const numberingDef: NumberingDef = { numId: numIdStr, abstractNumId, levels: new Map([[0, level]]) }
  const nextDocumentNumbering = new Map(bundle.document.numbering).set(numIdStr, numberingDef)

  return {
    ...bundle,
    numberingPart: nextNumberingPart,
    document: { ...bundle.document, numbering: nextDocumentNumbering },
  }
}
