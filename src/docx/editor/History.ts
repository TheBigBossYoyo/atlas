import type { Document } from '../model'

import { applyCommand } from './commands'
import type { Command, DeleteRangeCommand, InsertTextCommand, Position, Range } from './commandTypes'

const COALESCE_WINDOW_MS = 1000

type InsertSnapshot = {
  readonly command: InsertTextCommand
  readonly timestamp: number
}

/** An undo-stack entry: the inverse command that undoes it, plus the
 *  revision (see `History`'s own doc comment below) that was current right
 *  *before* this entry's edit was first made — what `currentRevision` is
 *  restored to when this entry is undone. */
type UndoEntry = {
  readonly command: Command
  readonly priorRevision: number
}

/** A redo-stack entry: the command that redoes it, plus the revision that
 *  was current right before the undo that produced this entry — what
 *  `currentRevision` is restored to when this entry is redone. */
type RedoEntry = {
  readonly command: Command
  readonly laterRevision: number
}

export class History {
  private readonly undoStack: UndoEntry[] = []
  private readonly redoStack: RedoEntry[] = []
  private lastInsert: InsertSnapshot | null = null

  // DIRTY-1 — a cheap, O(1) stand-in for reference-equality against the
  // last-saved document (see DocxViewer.tsx's dirty-tracking comment above
  // `lastSavedDocument`). `documentModel` is rebuilt as a brand-new object
  // graph on every edit, undo and redo included — undo/redo apply the
  // *stored inverse command*, they never hand back a previously-held
  // object — so two document states that are byte-for-byte identical (e.g.
  // "typed a character, then undid it") are never `===` to one another.
  //
  // A monotonic revision counter sidesteps that without ever walking the
  // document: every *forward* change — a fresh push, a keystroke coalesced
  // into an existing undo entry, or an external non-undoable edit like
  // adding a comment (see `bumpRevision`) — allocates a brand-new revision
  // number. Undo/redo never allocate; they restore the exact revision
  // number that was current before the entry they're replaying was first
  // created (`priorRevision`/`laterRevision` on each stack entry). So
  // "undo back to exactly what was on disk" restores exactly the revision
  // number that was current at save time, and the caller's dirty check
  // becomes a single integer comparison (`revision !== savedRevision`)
  // instead of a structural walk of a document that can run to hundreds of
  // pages.
  //
  // Alternatives considered (see the task's own framing):
  // - Deep-equal the document against the saved snapshot on every
  //   keystroke: explicitly rejected — O(document size) per keystroke.
  // - Content-addressed hashing computed at save/undo/redo boundaries:
  //   still O(1) per keystroke like this counter, but needs a stable hash
  //   over the *whole* document model (styles/numbering/comments/
  //   footnotes/headers/footers/every section), which is a lot of surface
  //   area to keep in sync as the model grows, for no benefit over a plain
  //   counter here.
  // - Have History retain and return the *original* object for a state it
  //   has already visited, so `===` becomes true again after a full undo:
  //   the smallest-looking option, but it doesn't actually hold up. `redo`
  //   re-derives its document by replaying the *stored command* against
  //   whatever object `undo` just produced, not by handing back a cached
  //   reference — so making redo (or an undo that follows a coalesced or
  //   composite command) reference-exact means caching a document per
  //   stack entry anyway, which is no simpler than a counter and adds a
  //   second document graph alive per edit for the caching to pay off.
  // The counter gets the same O(1) win as that last option without needing
  // every hop in an undo/redo chain to happen to line back up with a
  // cached object.
  private nextRevision = 1
  private currentRevision = 0

  /** The revision of the document state History believes is current now.
   *  Callers treat "dirty" as `getRevision() !== <revision at last save>`. */
  getRevision(): number {
    return this.currentRevision
  }

  /**
   * Allocates a fresh revision for a document change History itself was
   * never told about as a command — DocxViewer's comment and field/TOC
   * edits mutate the document directly without going through
   * `push`/`coalesceWithLast` (they are not undoable through this History
   * at all) — so `getRevision()` must still move forward or such an edit
   * would be silently missed by the dirty check.
   *
   * Also drops the redo stack, same as `push` does for a genuine new edit:
   * every pending redo entry's command was recorded against the document
   * as it stood *before* this change, and replaying one against the
   * document *after* this change is no longer guaranteed to land
   * correctly (e.g. a redo'd insert-text's `at` position could now name
   * the wrong run if this edit shifted paragraph indices).
   */
  bumpRevision(): number {
    this.redoStack.length = 0
    this.lastInsert = null
    this.currentRevision = this.nextRevision++
    return this.currentRevision
  }

  push(inverse: Command): void {
    this.undoStack.push({ command: inverse, priorRevision: this.currentRevision })
    this.redoStack.length = 0
    this.currentRevision = this.nextRevision++
  }

  /**
   * DXE-16 — undo applies the stored inverse and returns not just the
   * resulting document but the selection to restore: `applyCommand` already
   * computes the natural resulting position/range for whatever command it
   * just ran (e.g. a delete-range's own inverse collapses to its anchor), so
   * forwarding it here is what lets the caller (Input.ts) put the caret back
   * where the user was before the edit being undone, instead of leaving
   * whatever selection happened to be active beforehand.
   */
  undo(
    doc: Document,
  ): {
    document: Document
    redoCommand: Command
    range: Range | null
  } | null {
    const entry = this.undoStack.pop()
    if (entry === undefined) {
      return null
    }

    const result = applyCommand(doc, entry.command)
    this.redoStack.push({ command: result.inverse, laterRevision: this.currentRevision })
    this.currentRevision = entry.priorRevision
    this.lastInsert = null

    return {
      document: result.document,
      redoCommand: result.inverse,
      range: result.range ?? null,
    }
  }

  redo(
    doc: Document,
  ): {
    document: Document
    undoCommand: Command
    range: Range | null
  } | null {
    const entry = this.redoStack.pop()
    if (entry === undefined) {
      return null
    }

    const result = applyCommand(doc, entry.command)
    this.undoStack.push({ command: result.inverse, priorRevision: this.currentRevision })
    this.currentRevision = entry.laterRevision
    this.lastInsert = null

    return {
      document: result.document,
      undoCommand: result.inverse,
      range: result.range ?? null,
    }
  }

  coalesceWithLast(cmd: Command, inverse: Command): boolean {
    const now = Date.now()

    if (
      cmd.kind !== 'insert-text' ||
      inverse.kind !== 'delete-range' ||
      this.lastInsert === null ||
      now - this.lastInsert.timestamp > COALESCE_WINDOW_MS
    ) {
      this.lastInsert = cmd.kind === 'insert-text' ? { command: cloneInsertCommand(cmd), timestamp: now } : null
      return false
    }

    const lastUndo = this.undoStack[this.undoStack.length - 1]
    const lastUndoCommand = lastUndo?.command
    if (lastUndo === undefined || lastUndoCommand === undefined || lastUndoCommand.kind !== 'delete-range') {
      this.lastInsert = { command: cloneInsertCommand(cmd), timestamp: now }
      return false
    }

    if (!canCoalesceInsert(this.lastInsert.command, cmd, lastUndoCommand, inverse)) {
      this.lastInsert = { command: cloneInsertCommand(cmd), timestamp: now }
      return false
    }

    this.undoStack[this.undoStack.length - 1] = {
      command: {
        kind: 'delete-range',
        range: {
          anchor: clonePosition(lastUndoCommand.range.anchor),
          focus: clonePosition(inverse.range.focus),
        },
      },
      priorRevision: lastUndo.priorRevision,
    }
    this.redoStack.length = 0
    // DIRTY-1 — the document DID change further (another character was just
    // appended), so identity must still move forward even though this
    // keystroke coalesced into the existing undo entry rather than pushing a
    // new one: otherwise typing "a" (landing right after a save) then "b"
    // (coalescing into the same entry) would still compare equal to the
    // saved revision and wrongly report clean.
    this.currentRevision = this.nextRevision++
    this.lastInsert = {
      command: {
        kind: 'insert-text',
        at: clonePosition(this.lastInsert.command.at),
        text: this.lastInsert.command.text + cmd.text,
      },
      timestamp: now,
    }

    return true
  }
}

function canCoalesceInsert(
  previous: InsertTextCommand,
  current: InsertTextCommand,
  previousUndo: DeleteRangeCommand,
  currentUndo: DeleteRangeCommand,
): boolean {
  return (
    samePath(previous.at.paragraphPath, current.at.paragraphPath) &&
    previous.at.runIndex === current.at.runIndex &&
    current.at.charOffset === previous.at.charOffset + previous.text.length &&
    samePosition(previousUndo.range.focus, currentUndo.range.anchor)
  )
}

function samePath(left: ReadonlyArray<number>, right: ReadonlyArray<number>): boolean {
  if (left.length !== right.length) {
    return false
  }

  return left.every((value, index) => value === right[index])
}

function samePosition(left: Position, right: Position): boolean {
  return (
    samePath(left.paragraphPath, right.paragraphPath) &&
    left.runIndex === right.runIndex &&
    left.charOffset === right.charOffset
  )
}

function cloneInsertCommand(command: InsertTextCommand): InsertTextCommand {
  return {
    kind: 'insert-text',
    at: clonePosition(command.at),
    text: command.text,
  }
}

function clonePosition(position: Position): Position {
  return {
    paragraphPath: Object.freeze([...position.paragraphPath]),
    runIndex: position.runIndex,
    charOffset: position.charOffset,
  }
}
