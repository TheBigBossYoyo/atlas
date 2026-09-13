/**
 * Atlas — Find & Replace logic (Wave C.6)
 *
 * Pure functions. No side effects. No `any`.
 */

import type { Document, Block, Paragraph, ParagraphChild, Run } from '../model/document'
import type { Position, Range } from './Selection'
import type { Command, DeleteRangeCommand, InsertTextCommand } from './commandTypes'

// ---------------------------------------------------------------------------
// Public interfaces
// ---------------------------------------------------------------------------

export interface FindOptions {
  readonly caseSensitive: boolean
  readonly wholeWord: boolean
  readonly useRegex: boolean
}

export interface FindMatch {
  readonly range: Range
  readonly text: string
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/** Collect text-bearing runs from a paragraph, ignoring non-Run children. */
interface RunSlice {
  readonly run: Run
  readonly runIndex: number
  /** Offset of first char in the concatenated paragraph string. */
  readonly offset: number
}

function collectRunSlices(paragraph: Paragraph): ReadonlyArray<RunSlice> {
  const slices: RunSlice[] = []
  let offset = 0

  for (let i = 0; i < paragraph.children.length; i++) {
    const child: ParagraphChild = paragraph.children[i]
    if (child.kind === 'run') {
      slices.push({ run: child, runIndex: i, offset })
      // advance offset by text length of this run
      for (const rc of child.children) {
        if (rc.kind === 'text') {
          offset += rc.value.length
        } else if (rc.kind === 'tab') {
          offset += 1 // treat tab as single char
        }
      }
    } else if (child.kind === 'hyperlink') {
      // include runs inside hyperlinks, indexed at their hyperlink position
      for (const hc of child.children) {
        if (hc.kind === 'run') {
          slices.push({ run: hc, runIndex: i, offset })
          for (const rc of hc.children) {
            if (rc.kind === 'text') {
              offset += rc.value.length
            } else if (rc.kind === 'tab') {
              offset += 1
            }
          }
        }
      }
    }
  }

  return slices
}

/** Concatenate visible text of a paragraph. */
function paragraphText(slices: ReadonlyArray<RunSlice>): string {
  const parts: string[] = []
  for (const s of slices) {
    for (const rc of s.run.children) {
      if (rc.kind === 'text') {
        parts.push(rc.value)
      } else if (rc.kind === 'tab') {
        parts.push('\t')
      }
    }
  }
  return parts.join('')
}

/** Build a RegExp from query and options. Returns null if pattern is invalid. */
function buildRegex(query: string, opts: FindOptions): RegExp | null {
  if (query.length === 0) return null
  try {
    const flags = opts.caseSensitive ? 'g' : 'gi'
    let pattern = opts.useRegex ? query : query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    if (opts.wholeWord) {
      pattern = `\\b${pattern}\\b`
    }
    return new RegExp(pattern, flags)
  } catch {
    return null
  }
}

/**
 * Map an absolute char offset within the concatenated paragraph text back to a Position.
 * charOffset is the offset within the run text, not the run's absolute position.
 */
function absoluteOffsetToPosition(
  paragraphPath: ReadonlyArray<number>,
  slices: ReadonlyArray<RunSlice>,
  absoluteOffset: number,
): Position {
  // Find which run contains this offset
  for (let i = 0; i < slices.length; i++) {
    const slice = slices[i]
    // compute length of this run's text
    let runLen = 0
    for (const rc of slice.run.children) {
      if (rc.kind === 'text') runLen += rc.value.length
      else if (rc.kind === 'tab') runLen += 1
    }
    const sliceEnd = slice.offset + runLen
    if (absoluteOffset < sliceEnd || (i === slices.length - 1 && absoluteOffset <= sliceEnd)) {
      const charOffset = absoluteOffset - slice.offset
      return {
        paragraphPath: paragraphPath as ReadonlyArray<number>,
        runIndex: slice.runIndex,
        charOffset,
      }
    }
  }
  // Fallback: clamp to last run end
  const last = slices[slices.length - 1]
  if (last === undefined) {
    return { paragraphPath: paragraphPath as ReadonlyArray<number>, runIndex: 0, charOffset: 0 }
  }
  let runLen = 0
  for (const rc of last.run.children) {
    if (rc.kind === 'text') runLen += rc.value.length
    else if (rc.kind === 'tab') runLen += 1
  }
  return {
    paragraphPath: paragraphPath as ReadonlyArray<number>,
    runIndex: last.runIndex,
    charOffset: runLen,
  }
}

/** Walk all paragraphs in the document in order, yielding [paragraphPath, paragraph]. */
function* walkParagraphs(
  blocks: ReadonlyArray<Block>,
  pathPrefix: ReadonlyArray<number>,
): Generator<[ReadonlyArray<number>, Paragraph]> {
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i]
    const path = [...pathPrefix, i]
    if (block.kind === 'paragraph') {
      yield [path, block]
    } else if (block.kind === 'table') {
      for (let r = 0; r < block.rows.length; r++) {
        const row = block.rows[r]
        if (row.kind === 'table-row') {
          for (let c = 0; c < row.cells.length; c++) {
            const cell = row.cells[c]
            if (cell.kind === 'table-cell') {
              yield* walkParagraphs(cell.blocks, [...path, r, c])
            }
          }
        }
      }
    }
  }
}

function* walkDocumentParagraphs(
  doc: Document,
): Generator<[ReadonlyArray<number>, Paragraph]> {
  for (let s = 0; s < doc.sections.length; s++) {
    const section = doc.sections[s]
    // We use section index as prefix, but paragraphPath convention from commands uses
    // [sectionIndex, ...blockPath]. Checking commands.ts to align:
    // commands.ts findParagraph uses path[0] as section index.
    yield* walkParagraphs(section.blocks, [s])
  }
}

/** Find all matches within a single paragraph. */
function findInParagraph(
  paragraphPath: ReadonlyArray<number>,
  paragraph: Paragraph,
  regex: RegExp,
): ReadonlyArray<FindMatch> {
  const slices = collectRunSlices(paragraph)
  if (slices.length === 0) return []

  const text = paragraphText(slices)
  regex.lastIndex = 0

  const matches: FindMatch[] = []
  let match: RegExpExecArray | null

  while ((match = regex.exec(text)) !== null) {
    const start = match.index
    const end = start + match[0].length

    const anchor = absoluteOffsetToPosition(paragraphPath, slices, start)
    const focus = absoluteOffsetToPosition(paragraphPath, slices, end)

    matches.push({
      range: { anchor, focus },
      text: match[0],
    })

    // Prevent infinite loop on zero-length matches
    if (match[0].length === 0) {
      regex.lastIndex++
    }
  }

  return matches
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

export function findAll(
  doc: Document,
  query: string,
  opts: FindOptions,
): ReadonlyArray<FindMatch> {
  const regex = buildRegex(query, opts)
  if (regex === null) return []

  const results: FindMatch[] = []

  for (const [paragraphPath, paragraph] of walkDocumentParagraphs(doc)) {
    const found = findInParagraph(paragraphPath, paragraph, regex)
    results.push(...found)
  }

  return results
}

export function findNext(
  doc: Document,
  query: string,
  opts: FindOptions,
  from: Position | null,
): FindMatch | null {
  const all = findAll(doc, query, opts)
  if (all.length === 0) return null

  if (from === null) return all[0]

  // Find first match whose anchor is strictly after `from`
  for (const m of all) {
    const cmp = comparePositionsSimple(m.range.anchor, from)
    if (cmp > 0) return m
  }

  // Wrap around
  return all[0]
}

export function findPrev(
  doc: Document,
  query: string,
  opts: FindOptions,
  from: Position | null,
): FindMatch | null {
  const all = findAll(doc, query, opts)
  if (all.length === 0) return null

  if (from === null) return all[all.length - 1]

  // Find last match whose anchor is strictly before `from`
  let prev: FindMatch | null = null
  for (const m of all) {
    const cmp = comparePositionsSimple(m.range.anchor, from)
    if (cmp < 0) prev = m
  }

  // Wrap around
  return prev ?? all[all.length - 1]
}

export function buildReplaceCommands(
  match: FindMatch,
  replacement: string,
): ReadonlyArray<Command> {
  const deleteCmd: DeleteRangeCommand = {
    kind: 'delete-range',
    range: match.range,
  }
  const insertCmd: InsertTextCommand = {
    kind: 'insert-text',
    at: match.range.anchor,
    text: replacement,
  }
  return [deleteCmd, insertCmd]
}

export function buildReplaceAllCommands(
  doc: Document,
  query: string,
  opts: FindOptions,
  replacement: string,
): ReadonlyArray<Command> {
  const all = findAll(doc, query, opts)

  // Reverse order so applying from end backward keeps earlier offsets stable
  const reversed = [...all].reverse()

  const commands: Command[] = []
  for (const match of reversed) {
    commands.push(...buildReplaceCommands(match, replacement))
  }
  return commands
}

// ---------------------------------------------------------------------------
// Internal position comparison (avoids import cycle with Selection.ts)
// ---------------------------------------------------------------------------

function compareArrays(a: ReadonlyArray<number>, b: ReadonlyArray<number>): number {
  const len = Math.min(a.length, b.length)
  for (let i = 0; i < len; i++) {
    if (a[i] < b[i]) return -1
    if (a[i] > b[i]) return 1
  }
  return a.length - b.length
}

function comparePositionsSimple(a: Position, b: Position): number {
  const pc = compareArrays(a.paragraphPath, b.paragraphPath)
  if (pc !== 0) return pc
  if (a.runIndex !== b.runIndex) return a.runIndex - b.runIndex
  return a.charOffset - b.charOffset
}
