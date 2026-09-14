/**
 * Atlas — spellcheck-accept helper (D15/DXE-08)
 *
 * Accepting a spellcheck suggestion used to call only Electron's native
 * `webContents.replaceMisspelling`, which mutates the contentEditable DOM
 * directly (via the OS-level spellchecker) without touching the document
 * model at all — so the fix rendered correctly until the next
 * paginate/re-render, and was silently lost on save. This module builds a
 * real delete+insert editor command pair instead, so the caller can apply it
 * through the normal Command/History pipeline (`applyEditorCommand(s)`),
 * making the correction part of the model and undoable like any other edit.
 */

import type { Document } from '../model'
import type { Command, Position, Range } from './commandTypes'
import { buildReplaceCommands, findAll, type FindMatch } from './Find'

export interface SpellCheckReplacement {
  readonly commands: ReadonlyArray<Command>
  readonly range: Range
}

/**
 * Finds the occurrence of `misspelled` (matched as a whole word,
 * case-sensitively — spellcheck suggestions are for one exact spelling) at
 * or after `hint` (wrapping to the document's first occurrence if none is at
 * or after it, matching `findNext`'s convention elsewhere in this module),
 * and returns the delete+insert commands to replace it plus the resulting
 * cursor position. Returns `null` when there's nothing to do (empty
 * inputs, no matching occurrence, or the "correction" is a no-op).
 */
export function buildSpellCheckReplacement(
  doc: Document,
  misspelled: string,
  replacement: string,
  hint: Position | null,
): SpellCheckReplacement | null {
  if (misspelled.length === 0 || replacement.length === 0 || replacement === misspelled) {
    return null
  }

  const matches = findAll(doc, misspelled, { caseSensitive: true, wholeWord: true, useRegex: false })
  if (matches.length === 0) {
    return null
  }

  const match = pickMatch(matches, hint)
  const commands = buildReplaceCommands(match, replacement)
  const nextOffset = match.range.anchor.charOffset + replacement.length
  const nextPosition: Position = {
    paragraphPath: match.range.anchor.paragraphPath,
    runIndex: match.range.anchor.runIndex,
    charOffset: nextOffset,
  }

  return { commands, range: { anchor: nextPosition, focus: nextPosition } }
}

function pickMatch(matches: ReadonlyArray<FindMatch>, hint: Position | null): FindMatch {
  if (hint === null) {
    return matches[0]
  }

  for (const match of matches) {
    if (comparePositions(match.range.anchor, hint) >= 0) {
      return match
    }
  }

  return matches[0]
}

function compareArrays(a: ReadonlyArray<number>, b: ReadonlyArray<number>): number {
  const len = Math.min(a.length, b.length)
  for (let i = 0; i < len; i += 1) {
    if (a[i] < b[i]) return -1
    if (a[i] > b[i]) return 1
  }
  return a.length - b.length
}

function comparePositions(a: Position, b: Position): number {
  const pathCompare = compareArrays(a.paragraphPath, b.paragraphPath)
  if (pathCompare !== 0) return pathCompare
  if (a.runIndex !== b.runIndex) return a.runIndex - b.runIndex
  return a.charOffset - b.charOffset
}
