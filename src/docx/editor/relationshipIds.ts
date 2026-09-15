/**
 * Atlas — shared `rIdN` relationship-id allocation (DXE-18/D18/DXE-19).
 *
 * Image insertion, hyperlink insertion, and rich paste (which can insert
 * several of both images and hyperlinks in one go — DXE-19) each need to
 * mint a fresh relationship id that can never collide with one the source
 * document already defines. All three derive it the same way: one past the
 * highest numeric `rId<N>` suffix already in use across
 * `word/_rels/document.xml.rels`.
 */
import type { Relationship } from '../parser/relationships'

/**
 * One past the highest `rId<N>` already in use — the number
 * `allocateRelationshipId` turns into an id for a single insertion, or the
 * starting point for a counter a caller minting several ids in one batch
 * (rich paste's own resource pre-pass) increments itself as it allocates
 * each one.
 */
export function nextRelationshipIdNumber(existing: ReadonlyArray<Relationship>): number {
  let max = 0
  for (const rel of existing) {
    const match = /^rId(\d+)$/.exec(rel.id)
    if (match !== null) {
      const value = Number.parseInt(match[1] ?? '0', 10)
      if (Number.isFinite(value) && value > max) {
        max = value
      }
    }
  }
  return max + 1
}

export function allocateRelationshipId(existing: ReadonlyArray<Relationship>): string {
  return `rId${nextRelationshipIdNumber(existing)}`
}
