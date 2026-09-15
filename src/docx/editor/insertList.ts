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
import type { Document, LvlDef, NumberingDef } from '../model'

export type ListKind = 'bullet' | 'number'

const LIST_FORMAT_BY_KIND: Readonly<Record<ListKind, string>> = {
  bullet: 'bullet',
  number: 'decimal',
}

/**
 * DXE-06/D18 — picks the `numId` to use for a bullet/numbered list toggle
 * (the toolbar's, or DXE-19 rich paste's own list support). A real Word
 * document almost always already defines numId "1" (frequently "2" as well)
 * for its own lists — hardcoding those values here would mean toggling
 * "Bulleted List" on such a document silently reuses whatever list style the
 * document already assigned to numId 1 (rarely an actual bullet format) via
 * `ensureListNumbering`'s "reuse if already defined" rule, instead of
 * creating Atlas's own bullet definition.
 *
 * Reuses an Atlas-created list definition of the matching kind if one
 * already exists in this document (identified by the `atlas-list-` prefix
 * `ensureListNumbering` gives its own `abstractNumId`s, plus a matching
 * level-0 format) so repeated toggles/pastes of the same kind keep
 * converging on one shared definition; otherwise allocates one past every
 * numId already in use, which can never collide with the source document's
 * own numbering or with a different-kind list Atlas already created in this
 * session.
 */
export function pickListNumId(document: Document, kind: ListKind): number {
  const wantedFormat = LIST_FORMAT_BY_KIND[kind]
  let maxNumId = 0

  for (const [numIdStr, def] of document.numbering) {
    const parsed = Number.parseInt(numIdStr, 10)
    if (Number.isFinite(parsed) && parsed > maxNumId) {
      maxNumId = parsed
    }

    if (
      Number.isFinite(parsed) &&
      def.abstractNumId?.startsWith('atlas-list-') === true &&
      def.levels.get(0)?.format === wantedFormat
    ) {
      return parsed
    }
  }

  return maxNumId + 1
}

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
