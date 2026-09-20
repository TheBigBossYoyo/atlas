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
 * already exists in this document (identified by `NumberingDef.atlasManaged`
 * — see that field's doc comment for why this can no longer be a prefix on
 * `abstractNumId` itself — plus a matching level-0 format) so repeated
 * toggles/pastes of the same kind keep converging on one shared definition;
 * otherwise allocates one past every numId already in use, which can never
 * collide with the source document's own numbering or with a different-kind
 * list Atlas already created in this session.
 */
export function pickListNumId(numbering: ReadonlyMap<string, NumberingDef>, kind: ListKind): number {
  const wantedFormat = LIST_FORMAT_BY_KIND[kind]
  let maxNumId = 0

  for (const [numIdStr, def] of numbering) {
    const parsed = Number.parseInt(numIdStr, 10)
    if (Number.isFinite(parsed) && parsed > maxNumId) {
      maxNumId = parsed
    }

    if (
      Number.isFinite(parsed) &&
      def.atlasManaged === true &&
      def.levels.get(0)?.format === wantedFormat
    ) {
      return parsed
    }
  }

  return maxNumId + 1
}

export interface ListNumberingEntry {
  readonly abstractNum: AbstractNum
  readonly numInstance: NumInstance
  readonly numberingDef: NumberingDef
}

/**
 * DOCX-15 fix — `w:abstractNumId` (both the `w:abstractNum` element's own
 * attribute and the `w:num/w:abstractNumId/@w:val` that points at it) is
 * `ST_DecimalNumber` in the OOXML schema: a plain integer, never a string
 * tag. Returns one integer past every abstractNumId already present in the
 * document's numbering part (real ones from the source file and any Atlas
 * minted earlier in this session alike), so a freshly minted id can never
 * collide with one already in use regardless of how sparse or dense the
 * existing ids are. Non-numeric existing ids (there shouldn't be any once
 * this fix ships, but a document round-tripped through an older, buggy
 * Atlas build could still have one on disk) are simply ignored rather than
 * treated as a parse failure — this only needs to find the numeric
 * high-water mark.
 */
function nextAbstractNumId(existingAbstractNumIds: Iterable<string>): number {
  let max = -1
  for (const id of existingAbstractNumIds) {
    const parsed = Number.parseInt(id, 10)
    if (Number.isFinite(parsed) && parsed > max) {
      max = parsed
    }
  }
  return max + 1
}

/**
 * Builds the abstractNum/num/NumberingDef triple for a fresh single-level
 * (level 0 only) bullet or decimal-numbered list definition under `numId`.
 * Pure and side-effect-free so both `ensureListNumbering` (below, for the
 * toolbar's bundle-level list toggle) and DXE-19 rich paste's own
 * `pasteRich.ts` (which mints these against a locally-tracked numbering map
 * rather than a whole `DocxBundle`, potentially several in one paste) can
 * share the exact same definition shape. `existingAbstractNumIds` is every
 * `abstractNumId` already present in the numbering part this entry will be
 * added to — see `nextAbstractNumId` for why the caller supplies this
 * rather than the function tracking it internally: it's the same
 * "bundle-level state the pure helper doesn't own" reason `numId` itself is
 * a parameter rather than self-allocated.
 *
 * `abstractNumId` is kept identical across `abstractNum.abstractNumId`,
 * `numInstance.abstractNumId` and `numberingDef.abstractNumId` — the num
 * instance and the def both have to point at the abstract definition this
 * same call mints, not at some other id.
 */
export function createListNumberingEntry(
  numId: number,
  kind: ListKind,
  existingAbstractNumIds: Iterable<string>,
): ListNumberingEntry {
  const numIdStr = String(numId)
  const abstractNumId = String(nextAbstractNumId(existingAbstractNumIds))
  const level: LvlDef =
    kind === 'number'
      ? { level: 0, format: 'decimal', text: { value: '%1.', placeholders: [1] }, suffix: 'tab' }
      : { level: 0, format: 'bullet', text: { value: '•', placeholders: [] }, suffix: 'tab' }

  return {
    abstractNum: { abstractNumId, levels: new Map([[0, level]]) },
    numInstance: { numId: numIdStr, abstractNumId },
    // atlasManaged — see NumberingDef's doc comment — is what pickListNumId
    // now uses to recognize a definition Atlas minted for reuse, now that
    // abstractNumId itself is a plain integer and can no longer carry an
    // `atlas-list-` prefix.
    numberingDef: { numId: numIdStr, abstractNumId, levels: new Map([[0, level]]), atlasManaged: true },
  }
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

  const previousPart = bundle.numberingPart ?? { abstractNums: new Map(), nums: new Map() }
  const { abstractNum, numInstance, numberingDef } = createListNumberingEntry(
    numId,
    kind,
    previousPart.abstractNums.keys(),
  )

  const nextNumberingPart: NumberingPart = {
    abstractNums: new Map(previousPart.abstractNums).set(abstractNum.abstractNumId, abstractNum),
    nums: new Map(previousPart.nums).set(numInstance.numId, numInstance),
  }

  const nextDocumentNumbering = new Map(bundle.document.numbering).set(numIdStr, numberingDef)

  return {
    ...bundle,
    numberingPart: nextNumberingPart,
    document: { ...bundle.document, numbering: nextDocumentNumbering },
  }
}
