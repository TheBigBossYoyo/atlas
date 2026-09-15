import type { Document, Paragraph, Run, RunProps, Underline } from '../model'
import type { Command, Position, Range, TrackChangesContext } from './commandTypes'
import { applyCommand, findEnclosingTable, findParagraph, getFlatTextRuns } from './commands'
import type { EnclosingTable } from './commands'
import type { History } from './History'

// ─── Public interfaces ───────────────────────────────────────────────────────

export interface InputContext {
  readonly document: Document
  readonly range: Range | null
  readonly history: History
  /**
   * DXE-11 — when present and `enabled`, every insertion/deletion this module
   * applies directly (typing, Backspace/Delete, word-delete) is recorded as
   * `w:ins`/`w:del` instead of mutating the run in place. Deliberately never
   * forwarded to `history.undo`/`redo` (those replay a *stored inverse*
   * command, which must apply exactly as originally computed — see
   * `commands.ts`'s `applyInsertText`/`applyDeleteSpan` doc comments for why
   * that's what makes undo/redo of a tracked edit come out coherent) nor to
   * `applyFormatToggle` (formatting changes are not yet recorded as
   * `w:rPrChange` — a documented limitation, not an oversight).
   */
  readonly trackChanges?: TrackChangesContext
}

export interface InputResult {
  readonly document: Document
  readonly range: Range | null
}

// ─── Internal types ───────────────────────────────────────────────────────────

type EditableRun = {
  readonly text: string
  readonly run: Run
}

// ─── Utilities ────────────────────────────────────────────────────────────────

/**
 * DXE-21 fix — delegates to commands.ts's `getFlatTextRuns` (hyperlink and
 * atomic image/break/tab runs flattened, same as the edit pipeline) instead
 * of this module's own former hand-rolled "every child must be a plain
 * text-only run or reject the whole paragraph" version. See
 * `getFlatTextRuns`'s own comment for the cross-paragraph data-loss bug that
 * caused. `null` (a paragraph shape this editor doesn't support at all) is
 * intentionally NOT collapsed to `[]` here — callers that treat an empty
 * result as "this paragraph is empty, safe to cross into the next/previous
 * one" (see `wordBoundaryBefore`/`wordBoundaryAfter`) must tell that apart
 * from "couldn't be measured" or they'd repeat the same bug.
 */
function getTextRuns(paragraph: Paragraph): ReadonlyArray<EditableRun> {
  return getFlatTextRuns(paragraph) ?? []
}

function collectParagraphPaths(doc: Document): ReadonlyArray<ReadonlyArray<number>> {
  const paths: Array<ReadonlyArray<number>> = []
  for (let sIdx = 0; sIdx < doc.sections.length; sIdx++) {
    const section = doc.sections[sIdx]
    for (let bIdx = 0; bIdx < section.blocks.length; bIdx++) {
      const block = section.blocks[bIdx]
      if (block.kind === 'paragraph') {
        paths.push(Object.freeze([sIdx, bIdx]))
      }
    }
  }
  return paths
}

function pathEquals(a: ReadonlyArray<number>, b: ReadonlyArray<number>): boolean {
  if (a.length !== b.length) return false
  return a.every((v, i) => v === b[i])
}

function findPathIndex(
  paths: ReadonlyArray<ReadonlyArray<number>>,
  paragraphPath: ReadonlyArray<number>,
): number {
  // direct match
  for (let i = 0; i < paths.length; i++) {
    if (pathEquals(paths[i], paragraphPath)) return i
  }
  // normalise: single-element path [n] → [0, n]
  if (paragraphPath.length === 1) {
    const normalised = Object.freeze([0, paragraphPath[0]])
    for (let i = 0; i < paths.length; i++) {
      if (pathEquals(paths[i], normalised)) return i
    }
  }
  return -1
}

function makePos(
  paragraphPath: ReadonlyArray<number>,
  runIndex: number,
  charOffset: number,
): Position {
  return {
    paragraphPath: Object.freeze([...paragraphPath]),
    runIndex,
    charOffset,
  }
}

function paragraphEndPos(path: ReadonlyArray<number>, para: Paragraph): Position {
  const runs = getTextRuns(para)
  if (runs.length === 0) return makePos(path, 0, 0)
  const lastIdx = runs.length - 1
  return makePos(path, lastIdx, runs[lastIdx].text.length)
}

function samePos(a: Position, b: Position): boolean {
  return (
    a.runIndex === b.runIndex &&
    a.charOffset === b.charOffset &&
    pathEquals(a.paragraphPath, b.paragraphPath)
  )
}

function comparePos(a: Position, b: Position): number {
  for (let i = 0; i < Math.max(a.paragraphPath.length, b.paragraphPath.length); i++) {
    const av = a.paragraphPath[i] ?? -1
    const bv = b.paragraphPath[i] ?? -1
    if (av !== bv) return av - bv
  }
  if (a.runIndex !== b.runIndex) return a.runIndex - b.runIndex
  return a.charOffset - b.charOffset
}

function normaliseRange(range: Range): { start: Position; end: Position } {
  if (comparePos(range.anchor, range.focus) <= 0) {
    return { start: range.anchor, end: range.focus }
  }
  return { start: range.focus, end: range.anchor }
}

function incrementPath(path: ReadonlyArray<number>): ReadonlyArray<number> {
  if (path.length === 0) return Object.freeze([1])
  const next = [...path]
  next[next.length - 1] += 1
  return Object.freeze(next)
}

// ─── Cursor movement helpers (named exports) ─────────────────────────────────

export function moveCursorLeft(pos: Position, doc: Document): Position {
  if (pos.charOffset > 0) {
    return makePos(pos.paragraphPath, pos.runIndex, pos.charOffset - 1)
  }

  const para = findParagraph(doc, pos.paragraphPath)
  if (para && pos.runIndex > 0) {
    const runs = getTextRuns(para)
    const prevIdx = pos.runIndex - 1
    const prev = runs[prevIdx]
    if (prev) {
      return makePos(pos.paragraphPath, prevIdx, prev.text.length)
    }
  }

  // Move to end of previous paragraph
  const allPaths = collectParagraphPaths(doc)
  const pIdx = findPathIndex(allPaths, pos.paragraphPath)
  if (pIdx > 0) {
    const prevPath = allPaths[pIdx - 1]
    const prevPara = findParagraph(doc, prevPath)
    if (prevPara) return paragraphEndPos(prevPath, prevPara)
  }

  return pos // clamped at doc start
}

export function moveCursorRight(pos: Position, doc: Document): Position {
  const para = findParagraph(doc, pos.paragraphPath)
  if (para) {
    const runs = getTextRuns(para)
    const cur = runs[pos.runIndex]
    if (cur && pos.charOffset < cur.text.length) {
      return makePos(pos.paragraphPath, pos.runIndex, pos.charOffset + 1)
    }
    if (pos.runIndex < runs.length - 1) {
      return makePos(pos.paragraphPath, pos.runIndex + 1, 0)
    }
  }

  // Move to start of next paragraph
  const allPaths = collectParagraphPaths(doc)
  const pIdx = findPathIndex(allPaths, pos.paragraphPath)
  if (pIdx >= 0 && pIdx < allPaths.length - 1) {
    return makePos(allPaths[pIdx + 1], 0, 0)
  }

  return pos // clamped at doc end
}

export function moveCursorToLineStart(pos: Position, doc: Document): Position {
  void doc
  return makePos(pos.paragraphPath, 0, 0)
}

export function moveCursorToLineEnd(pos: Position, doc: Document): Position {
  const para = findParagraph(doc, pos.paragraphPath)
  if (!para) return pos
  return paragraphEndPos(pos.paragraphPath, para)
}

export function moveCursorToDocStart(doc: Document): Position {
  const allPaths = collectParagraphPaths(doc)
  if (allPaths.length === 0) return makePos([0, 0], 0, 0)
  return makePos(allPaths[0], 0, 0)
}

export function moveCursorToDocEnd(doc: Document): Position {
  const allPaths = collectParagraphPaths(doc)
  if (allPaths.length === 0) return makePos([0, 0], 0, 0)
  const lastPath = allPaths[allPaths.length - 1]
  const lastPara = findParagraph(doc, lastPath)
  if (!lastPara) return makePos(lastPath, 0, 0)
  return paragraphEndPos(lastPath, lastPara)
}

export function extendOrCollapse(
  currentRange: Range | null,
  newFocus: Position,
  shiftHeld: boolean,
): Range {
  if (shiftHeld && currentRange !== null) {
    return { anchor: currentRange.anchor, focus: newFocus }
  }
  return { anchor: newFocus, focus: newFocus }
}

// ─── Word-boundary helpers ────────────────────────────────────────────────────
//
// DXE-21 — word-boundary delete used to search only the run the cursor sits
// in, so Ctrl+Backspace/Ctrl+Delete silently did nothing once the cursor was
// within a word of a run or paragraph boundary. These now flatten the whole
// paragraph's text (crossing run boundaries transparently) and, when no
// boundary is found within the paragraph, continue the search into the
// adjacent paragraph.

function paragraphFlatText(runs: ReadonlyArray<EditableRun>): string {
  return runs.map((run) => run.text).join('')
}

function flatOffsetToPos(path: ReadonlyArray<number>, runs: ReadonlyArray<EditableRun>, offset: number): Position {
  let remaining = offset
  for (let i = 0; i < runs.length; i += 1) {
    if (remaining <= runs[i].text.length) {
      return makePos(path, i, remaining)
    }
    remaining -= runs[i].text.length
  }
  const lastIndex = runs.length - 1
  return makePos(path, lastIndex < 0 ? 0 : lastIndex, runs[lastIndex]?.text.length ?? 0)
}

function adjacentParagraphBoundary(
  pos: Position,
  doc: Document,
  direction: 'previous' | 'next',
): Position | null {
  const allPaths = collectParagraphPaths(doc)
  const index = findPathIndex(allPaths, pos.paragraphPath)
  if (index < 0) return null

  const targetIndex = direction === 'previous' ? index - 1 : index + 1
  const targetPath = allPaths[targetIndex]
  if (targetPath === undefined) return null

  const targetPara = findParagraph(doc, targetPath)
  if (!targetPara) return null

  return direction === 'previous' ? paragraphEndPos(targetPath, targetPara) : makePos(targetPath, 0, 0)
}

function wordBoundaryBefore(pos: Position, doc: Document): Position | null {
  const para = findParagraph(doc, pos.paragraphPath)
  if (!para) return null
  const runs = getTextRuns(para)

  if (runs.length === 0) {
    // A paragraph with children that still flattened to zero runs is a
    // shape getFlatTextRuns couldn't measure (not genuinely empty) — bail
    // out rather than guessing, so an unsupported shape never gets treated
    // as "empty, safe to delete into" (see getFlatTextRuns's doc comment).
    if (para.children.length > 0) return null
    return adjacentParagraphBoundary(pos, doc, 'previous')
  }

  const flatOffset = positionToFlatOffset(runs, pos)
  const before = paragraphFlatText(runs).slice(0, flatOffset)
  const trimmed = before.trimEnd()

  if (trimmed.length === 0) {
    if (before.length > 0) {
      // Only leading whitespace before the cursor: consume it, landing at the
      // paragraph start, rather than crossing over (matches deleting a run of
      // spaces before reaching the previous word/paragraph).
      return flatOffsetToPos(pos.paragraphPath, runs, 0)
    }
    return adjacentParagraphBoundary(pos, doc, 'previous')
  }

  const lastSpace = trimmed.lastIndexOf(' ')
  const newOffset = lastSpace >= 0 ? lastSpace + 1 : 0
  if (newOffset === flatOffset) return null
  return flatOffsetToPos(pos.paragraphPath, runs, newOffset)
}

function wordBoundaryAfter(pos: Position, doc: Document): Position | null {
  const para = findParagraph(doc, pos.paragraphPath)
  if (!para) return null
  const runs = getTextRuns(para)

  if (runs.length === 0) {
    // See the matching guard in wordBoundaryBefore.
    if (para.children.length > 0) return null
    return adjacentParagraphBoundary(pos, doc, 'next')
  }

  const flatText = paragraphFlatText(runs)
  const flatOffset = positionToFlatOffset(runs, pos)
  const after = flatText.slice(flatOffset)
  const trimmed = after.trimStart()

  if (trimmed.length === 0) {
    if (after.length > 0) {
      return flatOffsetToPos(pos.paragraphPath, runs, flatText.length)
    }
    return adjacentParagraphBoundary(pos, doc, 'next')
  }

  const firstSpace = trimmed.indexOf(' ')
  const newOffset =
    firstSpace >= 0
      ? flatOffset + (after.length - trimmed.length) + firstSpace + 1
      : flatOffset + after.length
  if (newOffset === flatOffset) return null
  return flatOffsetToPos(pos.paragraphPath, runs, newOffset)
}

// ─── Format toggle ────────────────────────────────────────────────────────────

export type FormatKey = 'bold' | 'italic' | 'underline'

function positionToFlatOffset(runs: ReadonlyArray<EditableRun>, pos: Position): number {
  let offset = 0
  for (let i = 0; i < pos.runIndex && i < runs.length; i++) {
    offset += runs[i].text.length
  }
  return offset + pos.charOffset
}

export function isFormatActive(range: Range, doc: Document, key: FormatKey): boolean {
  const para = findParagraph(doc, range.anchor.paragraphPath)
  if (!para) return false
  // only works within a single paragraph (cross-paragraph format is not yet supported)
  if (!pathEquals(range.anchor.paragraphPath, range.focus.paragraphPath)) return false

  const runs = getTextRuns(para)
  const startOffset = Math.min(
    positionToFlatOffset(runs, range.anchor),
    positionToFlatOffset(runs, range.focus),
  )
  const endOffset = Math.max(
    positionToFlatOffset(runs, range.anchor),
    positionToFlatOffset(runs, range.focus),
  )
  if (startOffset === endOffset) return false

  let currentOffset = 0
  for (const { text, run } of runs) {
    const runEnd = currentOffset + text.length
    if (runEnd > startOffset && currentOffset < endOffset) {
      const active =
        key === 'underline' ? run.props?.underline !== undefined : run.props?.[key] === true
      if (!active) return false
    }
    currentOffset = runEnd
  }
  return true
}

export function buildFormatPatch(key: FormatKey, turnOn: boolean): Partial<RunProps> {
  if (key === 'underline') {
    if (turnOn) {
      const u: Underline = { style: 'single' }
      return { underline: u }
    }
    return { underline: undefined }
  }
  return { [key]: turnOn ? true : undefined } as Partial<RunProps>
}

function applyFormatToggle(
  key: FormatKey,
  range: Range | null,
  doc: Document,
  history: History,
): InputResult | null {
  if (!range || samePos(range.anchor, range.focus)) return null
  const active = isFormatActive(range, doc, key)
  const format = buildFormatPatch(key, !active)
  const cmd: Command = { kind: 'apply-run-format', range, format }
  try {
    const result = applyCommand(doc, cmd)
    history.push(result.inverse)
    // USR-09 — formatting splits runs, so the pre-command range addresses
    // stale run indices; keep the command's re-mapped selection instead
    // (returning the old one collapsed the visible selection).
    return { document: result.document, range: result.range ?? range }
  } catch {
    return null
  }
}

// ─── handleBeforeInput ────────────────────────────────────────────────────────

export function handleBeforeInput(
  event: InputEvent,
  ctx: InputContext,
): InputResult | null {
  const { document, range, history, trackChanges } = ctx
  if (range === null) return null

  const focusPos = range.focus

  try {
    switch (event.inputType) {
      case 'insertText': {
        const text = event.data ?? ''
        if (!text) return null

        let workingDoc = document
        let insertAt = focusPos

        // If selection is non-collapsed, delete it first
        if (!samePos(range.anchor, range.focus)) {
          const { start } = normaliseRange(range)
          const deleteCmd: Command = { kind: 'delete-range', range }
          const deleteResult = applyCommand(workingDoc, deleteCmd, trackChanges)
          workingDoc = deleteResult.document
          history.push(deleteResult.inverse)
          insertAt = start
        }

        const cmd: Command = { kind: 'insert-text', at: insertAt, text }
        const result = applyCommand(workingDoc, cmd, trackChanges)
        const coalesced = history.coalesceWithLast(cmd, result.inverse)
        if (!coalesced) history.push(result.inverse)

        // DXE-11 — `applyInsertText` always returns its own precisely
        // computed post-insertion caret position (needed now that a tracked
        // insertion can shift which run/offset the caret lands in — see its
        // own doc comment); `applyCommand`'s general signature just can't
        // express that as non-optional across every command kind, so fall
        // back to the simple same-run computation only in the type-level
        // case this can't actually happen for `insert-text`.
        const fallbackPos = makePos(insertAt.paragraphPath, insertAt.runIndex, insertAt.charOffset + text.length)
        return { document: result.document, range: result.range ?? { anchor: fallbackPos, focus: fallbackPos } }
      }

      case 'insertParagraph': {
        const at = focusPos
        const cmd: Command = { kind: 'insert-paragraph-break', at }
        const result = applyCommand(document, cmd)
        history.push(result.inverse)
        const newPath = incrementPath(at.paragraphPath)
        const newPos = makePos(newPath, 0, 0)
        return { document: result.document, range: { anchor: newPos, focus: newPos } }
      }

      case 'deleteContentBackward': {
        if (!samePos(range.anchor, range.focus)) {
          const { start } = normaliseRange(range)
          const cmd: Command = { kind: 'delete-range', range }
          const result = applyCommand(document, cmd, trackChanges)
          history.push(result.inverse)
          return { document: result.document, range: { anchor: start, focus: start } }
        }
        const before = moveCursorLeft(focusPos, document)
        if (samePos(before, focusPos)) return null
        const deleteRange: Range = { anchor: before, focus: focusPos }
        const cmd: Command = { kind: 'delete-range', range: deleteRange }
        const result = applyCommand(document, cmd, trackChanges)
        history.push(result.inverse)
        return { document: result.document, range: { anchor: before, focus: before } }
      }

      case 'deleteContentForward': {
        if (!samePos(range.anchor, range.focus)) {
          const { start } = normaliseRange(range)
          const cmd: Command = { kind: 'delete-range', range }
          const result = applyCommand(document, cmd, trackChanges)
          history.push(result.inverse)
          return { document: result.document, range: { anchor: start, focus: start } }
        }
        const after = moveCursorRight(focusPos, document)
        if (samePos(after, focusPos)) return null
        const deleteRange: Range = { anchor: focusPos, focus: after }
        const cmd: Command = { kind: 'delete-range', range: deleteRange }
        const result = applyCommand(document, cmd, trackChanges)
        history.push(result.inverse)
        return { document: result.document, range: { anchor: focusPos, focus: focusPos } }
      }

      case 'deleteWordBackward': {
        const wordBefore = wordBoundaryBefore(focusPos, document)
        if (!wordBefore) return null
        const deleteRange: Range = { anchor: wordBefore, focus: focusPos }
        const cmd: Command = { kind: 'delete-range', range: deleteRange }
        const result = applyCommand(document, cmd, trackChanges)
        history.push(result.inverse)
        return { document: result.document, range: { anchor: wordBefore, focus: wordBefore } }
      }

      case 'deleteWordForward': {
        const wordAfter = wordBoundaryAfter(focusPos, document)
        if (!wordAfter) return null
        const deleteRange: Range = { anchor: focusPos, focus: wordAfter }
        const cmd: Command = { kind: 'delete-range', range: deleteRange }
        const result = applyCommand(document, cmd, trackChanges)
        history.push(result.inverse)
        return { document: result.document, range: { anchor: focusPos, focus: focusPos } }
      }

      default:
        return null
    }
  } catch {
    return null
  }
}

// ─── Table cell navigation (DXE-14) ──────────────────────────────────────────

function tableCellStartPath(tablePath: ReadonlyArray<number>, rowIndex: number, cellIndex: number): ReadonlyArray<number> {
  return Object.freeze([...tablePath, rowIndex, cellIndex, 0])
}

/**
 * The selection Tab/Shift+Tab lands on: the *entire* content of the
 * destination cell (mirroring Word, whose Tab both moves focus and primes
 * the cell for immediate overtype) — from the start of its first paragraph
 * to the end of its last.
 */
function tableCellSelectionRange(
  enclosing: EnclosingTable,
  rowIndex: number,
  cellIndex: number,
): Range | null {
  const row = enclosing.table.rows[rowIndex]
  if (row === undefined || row.kind !== 'table-row') return null
  const cell = row.cells[cellIndex]
  if (cell === undefined || cell.kind !== 'table-cell') return null

  const startPath = tableCellStartPath(enclosing.tablePath, rowIndex, cellIndex)
  const anchor = makePos(startPath, 0, 0)

  const lastBlockIndex = cell.blocks.length - 1
  const lastBlock = cell.blocks[lastBlockIndex]
  const endPath = Object.freeze([...enclosing.tablePath, rowIndex, cellIndex, lastBlockIndex])
  const focus = lastBlock?.kind === 'paragraph' ? paragraphEndPos(endPath, lastBlock) : makePos(endPath, 0, 0)

  return { anchor, focus }
}

/**
 * `undefined` when `range.focus` isn't inside a table at all (caller should
 * fall through to its own non-table Tab handling); otherwise the navigation
 * result to return directly from `handleKeyDown` — including `null` for "no
 * further cell to move to" (e.g. Shift+Tab at the table's first cell), which
 * must NOT fall through to inserting a literal tab character the way a
 * non-table Tab press would.
 */
function handleTableTab(
  range: Range,
  document: Document,
  history: History,
  shift: boolean,
): InputResult | null | undefined {
  const enclosing = findEnclosingTable(document, range.focus.paragraphPath)
  if (enclosing === null) {
    return undefined
  }

  const { table, tablePath, rowIndex, cellIndex } = enclosing
  const row = table.rows[rowIndex]
  if (row === undefined || row.kind !== 'table-row') {
    return undefined
  }

  if (!shift) {
    if (cellIndex + 1 < row.cells.length) {
      return { document, range: tableCellSelectionRange(enclosing, rowIndex, cellIndex + 1) ?? range }
    }
    if (rowIndex + 1 < table.rows.length) {
      return { document, range: tableCellSelectionRange(enclosing, rowIndex + 1, 0) ?? range }
    }

    // Last cell of the last row: append a new row and move into it.
    const cmd: Command = { kind: 'insert-table-row', tablePath, at: table.rows.length }
    const result = applyCommand(document, cmd)
    history.push(result.inverse)
    const nextPos = makePos(tableCellStartPath(tablePath, table.rows.length, 0), 0, 0)
    return { document: result.document, range: { anchor: nextPos, focus: nextPos } }
  }

  if (cellIndex > 0) {
    return { document, range: tableCellSelectionRange(enclosing, rowIndex, cellIndex - 1) ?? range }
  }
  if (rowIndex > 0) {
    const previousRow = table.rows[rowIndex - 1]
    if (previousRow === undefined || previousRow.kind !== 'table-row') {
      return null
    }
    return {
      document,
      range: tableCellSelectionRange(enclosing, rowIndex - 1, previousRow.cells.length - 1) ?? range,
    }
  }

  // Already at the table's first cell — Shift+Tab has nowhere to go.
  return null
}

// ─── handleKeyDown ────────────────────────────────────────────────────────────

export function handleKeyDown(
  event: KeyboardEvent,
  ctx: InputContext,
): InputResult | null {
  const { document, range, history } = ctx
  const ctrl = event.ctrlKey || event.metaKey
  const shift = event.shiftKey
  const key = event.key

  try {
    // Undo — DXE-16: restore the selection the edit itself computed (e.g. an
    // insert's inverse collapses to where the user started typing), not
    // whatever the current selection happens to be.
    if (ctrl && key === 'z' && !shift) {
      const result = history.undo(document)
      if (!result) return null
      return { document: result.document, range: result.range ?? range }
    }

    // Redo
    if (ctrl && ((key === 'z' && shift) || key === 'Z' || key === 'y' || key === 'Y')) {
      const result = history.redo(document)
      if (!result) return null
      return { document: result.document, range: result.range ?? range }
    }

    // Select all
    if (ctrl && (key === 'a' || key === 'A')) {
      const start = moveCursorToDocStart(document)
      const end = moveCursorToDocEnd(document)
      return { document, range: { anchor: start, focus: end } }
    }

    // Format toggles
    if (ctrl && (key === 'b' || key === 'B')) {
      return applyFormatToggle('bold', range, document, history)
    }
    if (ctrl && (key === 'i' || key === 'I')) {
      return applyFormatToggle('italic', range, document, history)
    }
    if (ctrl && (key === 'u' || key === 'U')) {
      return applyFormatToggle('underline', range, document, history)
    }

    // Arrow keys — need a range to navigate
    if (!range) return null

    if (key === 'ArrowLeft') {
      if (!shift && !samePos(range.anchor, range.focus)) {
        const { start } = normaliseRange(range)
        return { document, range: { anchor: start, focus: start } }
      }
      const newFocus = moveCursorLeft(range.focus, document)
      return { document, range: extendOrCollapse(range, newFocus, shift) }
    }

    if (key === 'ArrowRight') {
      if (!shift && !samePos(range.anchor, range.focus)) {
        const { end } = normaliseRange(range)
        return { document, range: { anchor: end, focus: end } }
      }
      const newFocus = moveCursorRight(range.focus, document)
      return { document, range: extendOrCollapse(range, newFocus, shift) }
    }

    // Vertical movement deferred (needs paginator)
    if (key === 'ArrowUp' || key === 'ArrowDown' || key === 'PageUp' || key === 'PageDown') {
      // TODO: vertical movement needs paginator output — deferred to a later wave
      return null
    }

    if (key === 'Home') {
      const newFocus = ctrl
        ? moveCursorToDocStart(document)
        : moveCursorToLineStart(range.focus, document)
      return { document, range: extendOrCollapse(range, newFocus, shift) }
    }

    if (key === 'End') {
      const newFocus = ctrl
        ? moveCursorToDocEnd(document)
        : moveCursorToLineEnd(range.focus, document)
      return { document, range: extendOrCollapse(range, newFocus, shift) }
    }

    if (key === 'Backspace') {
      return handleBeforeInput(
        new InputEvent('beforeinput', { inputType: 'deleteContentBackward' }),
        ctx,
      )
    }

    if (key === 'Delete') {
      return handleBeforeInput(
        new InputEvent('beforeinput', { inputType: 'deleteContentForward' }),
        ctx,
      )
    }

    if (key === 'Enter' && !shift) {
      return handleBeforeInput(
        new InputEvent('beforeinput', { inputType: 'insertParagraph' }),
        ctx,
      )
    }

    if (key === 'Enter' && shift) {
      // Shift+Enter: line break — deferred (needs BreakNode insertion)
      return null
    }

    if (key === 'Tab') {
      // DXE-14 — Tab/Shift+Tab inside a table cell moves between cells
      // instead (mirrors Word, and takes priority over the list-indent and
      // plain-tab-character behavior below, which only apply outside a
      // table). Tab in the last cell of the last row appends a new row and
      // moves into it; Shift+Tab at the very first cell is a no-op.
      if (range) {
        const tableResult = handleTableTab(range, document, history, shift)
        if (tableResult !== undefined) {
          return tableResult
        }
      }

      // D18 — Tab/Shift+Tab at the very start of a list paragraph changes its
      // outline level (mirrors Word); anywhere else Tab still inserts a
      // literal tab character, and Shift+Tab outside a list is a no-op.
      if (range && samePos(range.anchor, range.focus) && range.focus.runIndex === 0 && range.focus.charOffset === 0) {
        const paragraph = findParagraph(document, range.focus.paragraphPath)
        if (paragraph?.props?.numPr?.numId !== undefined) {
          const cmd: Command = {
            kind: 'change-list-level',
            paragraphPath: range.focus.paragraphPath,
            delta: shift ? -1 : 1,
          }
          const result = applyCommand(document, cmd)
          history.push(result.inverse)
          return { document: result.document, range }
        }
      }

      if (shift) return null

      return handleBeforeInput(
        new InputEvent('beforeinput', { inputType: 'insertText', data: '\t' }),
        ctx,
      )
    }

    return null
  } catch {
    return null
  }
}
