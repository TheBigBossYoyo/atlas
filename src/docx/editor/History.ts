import type { Document } from '../model'

import { applyCommand } from './commands'
import type { Command, DeleteRangeCommand, InsertTextCommand, Position } from './commandTypes'

const COALESCE_WINDOW_MS = 1000

type InsertSnapshot = {
  readonly command: InsertTextCommand
  readonly timestamp: number
}

export class History {
  private readonly undoStack: Command[] = []
  private readonly redoStack: Command[] = []
  private lastInsert: InsertSnapshot | null = null

  push(inverse: Command): void {
    this.undoStack.push(inverse)
    this.redoStack.length = 0
  }

  undo(
    doc: Document,
  ): {
    document: Document
    redoCommand: Command
  } | null {
    const inverse = this.undoStack.pop()
    if (inverse === undefined) {
      return null
    }

    const result = applyCommand(doc, inverse)
    this.redoStack.push(result.inverse)
    this.lastInsert = null

    return {
      document: result.document,
      redoCommand: result.inverse,
    }
  }

  redo(
    doc: Document,
  ): {
    document: Document
    undoCommand: Command
  } | null {
    const command = this.redoStack.pop()
    if (command === undefined) {
      return null
    }

    const result = applyCommand(doc, command)
    this.undoStack.push(result.inverse)
    this.lastInsert = null

    return {
      document: result.document,
      undoCommand: result.inverse,
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
    if (lastUndo === undefined || lastUndo.kind !== 'delete-range') {
      this.lastInsert = { command: cloneInsertCommand(cmd), timestamp: now }
      return false
    }

    if (!canCoalesceInsert(this.lastInsert.command, cmd, lastUndo, inverse)) {
      this.lastInsert = { command: cloneInsertCommand(cmd), timestamp: now }
      return false
    }

    this.undoStack[this.undoStack.length - 1] = {
      kind: 'delete-range',
      range: {
        anchor: clonePosition(lastUndo.range.anchor),
        focus: clonePosition(inverse.range.focus),
      },
    }
    this.redoStack.length = 0
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
