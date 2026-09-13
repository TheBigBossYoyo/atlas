import type { Document, Paragraph, Run, RunProps, Underline } from '../model'
import type { Command, Position, Range } from './commandTypes'
import { applyCommand, findParagraph } from './commands'
import type { History } from './History'

// ─── Public interfaces ───────────────────────────────────────────────────────

export interface InputContext {
  readonly document: Document
  readonly range: Range | null
  readonly history: History
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

function getTextRuns(paragraph: Paragraph): ReadonlyArray<EditableRun> {
  const result: EditableRun[] = []
  for (const child of paragraph.children) {
    if (child.kind !== 'run') return []
    let text = ''
    for (const c of child.children) {
      if (c.kind !== 'text') return []
      text += c.value
    }
    result.push({ text, run: child })
  }
  return result
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

function wordBoundaryBefore(pos: Position, doc: Document): Position | null {
  const para = findParagraph(doc, pos.paragraphPath)
  if (!para) return null
  const runs = getTextRuns(para)
  const cur = runs[pos.runIndex]
  if (!cur) return null

  const before = cur.text.slice(0, pos.charOffset)
  const trimmed = before.trimEnd()
  const lastSpace = trimmed.lastIndexOf(' ')
  const newOffset = lastSpace >= 0 ? lastSpace + 1 : 0
  if (newOffset === pos.charOffset) return null
  return makePos(pos.paragraphPath, pos.runIndex, newOffset)
}

function wordBoundaryAfter(pos: Position, doc: Document): Position | null {
  const para = findParagraph(doc, pos.paragraphPath)
  if (!para) return null
  const runs = getTextRuns(para)
  const cur = runs[pos.runIndex]
  if (!cur) return null

  const after = cur.text.slice(pos.charOffset)
  const trimmed = after.trimStart()
  const firstSpace = trimmed.indexOf(' ')
  const newOffset =
    firstSpace >= 0
      ? pos.charOffset + (after.length - trimmed.length) + firstSpace + 1
      : pos.charOffset + after.length
  if (newOffset === pos.charOffset) return null
  return makePos(pos.paragraphPath, pos.runIndex, newOffset)
}

// ─── Format toggle ────────────────────────────────────────────────────────────

type FormatKey = 'bold' | 'italic' | 'underline'

function positionToFlatOffset(runs: ReadonlyArray<EditableRun>, pos: Position): number {
  let offset = 0
  for (let i = 0; i < pos.runIndex && i < runs.length; i++) {
    offset += runs[i].text.length
  }
  return offset + pos.charOffset
}

function isFormatActive(range: Range, doc: Document, key: FormatKey): boolean {
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

function buildFormatPatch(key: FormatKey, turnOn: boolean): Partial<RunProps> {
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
    return { document: result.document, range }
  } catch {
    return null
  }
}

// ─── handleBeforeInput ────────────────────────────────────────────────────────

export function handleBeforeInput(
  event: InputEvent,
  ctx: InputContext,
): InputResult | null {
  const { document, range, history } = ctx
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
          const deleteResult = applyCommand(workingDoc, deleteCmd)
          workingDoc = deleteResult.document
          history.push(deleteResult.inverse)
          insertAt = start
        }

        const cmd: Command = { kind: 'insert-text', at: insertAt, text }
        const result = applyCommand(workingDoc, cmd)
        const coalesced = history.coalesceWithLast(cmd, result.inverse)
        if (!coalesced) history.push(result.inverse)

        const newOffset = insertAt.charOffset + text.length
        const newPos = makePos(insertAt.paragraphPath, insertAt.runIndex, newOffset)
        return { document: result.document, range: { anchor: newPos, focus: newPos } }
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
          const result = applyCommand(document, cmd)
          history.push(result.inverse)
          return { document: result.document, range: { anchor: start, focus: start } }
        }
        const before = moveCursorLeft(focusPos, document)
        if (samePos(before, focusPos)) return null
        const deleteRange: Range = { anchor: before, focus: focusPos }
        const cmd: Command = { kind: 'delete-range', range: deleteRange }
        const result = applyCommand(document, cmd)
        history.push(result.inverse)
        return { document: result.document, range: { anchor: before, focus: before } }
      }

      case 'deleteContentForward': {
        if (!samePos(range.anchor, range.focus)) {
          const { start } = normaliseRange(range)
          const cmd: Command = { kind: 'delete-range', range }
          const result = applyCommand(document, cmd)
          history.push(result.inverse)
          return { document: result.document, range: { anchor: start, focus: start } }
        }
        const after = moveCursorRight(focusPos, document)
        if (samePos(after, focusPos)) return null
        const deleteRange: Range = { anchor: focusPos, focus: after }
        const cmd: Command = { kind: 'delete-range', range: deleteRange }
        const result = applyCommand(document, cmd)
        history.push(result.inverse)
        return { document: result.document, range: { anchor: focusPos, focus: focusPos } }
      }

      case 'deleteWordBackward': {
        const wordBefore = wordBoundaryBefore(focusPos, document)
        if (!wordBefore) return null
        const deleteRange: Range = { anchor: wordBefore, focus: focusPos }
        const cmd: Command = { kind: 'delete-range', range: deleteRange }
        const result = applyCommand(document, cmd)
        history.push(result.inverse)
        return { document: result.document, range: { anchor: wordBefore, focus: wordBefore } }
      }

      case 'deleteWordForward': {
        const wordAfter = wordBoundaryAfter(focusPos, document)
        if (!wordAfter) return null
        const deleteRange: Range = { anchor: focusPos, focus: wordAfter }
        const cmd: Command = { kind: 'delete-range', range: deleteRange }
        const result = applyCommand(document, cmd)
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
    // Undo
    if (ctrl && key === 'z' && !shift) {
      const result = history.undo(document)
      if (!result) return null
      return { document: result.document, range }
    }

    // Redo
    if (ctrl && ((key === 'z' && shift) || key === 'Z' || key === 'y' || key === 'Y')) {
      const result = history.redo(document)
      if (!result) return null
      return { document: result.document, range }
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

    if (key === 'Tab' && !shift) {
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
