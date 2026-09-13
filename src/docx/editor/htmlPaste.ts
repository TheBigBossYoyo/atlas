/**
 * Atlas — paste-from-Word/HTML helper (Wave E.5)
 *
 * Converts pasted HTML or plain text into an ordered sequence of editor
 * Commands plus a final cursor position.  Pasting is implemented as:
 *
 *   1. (optional) delete the current selection
 *   2. for each paragraph in the paste:
 *        a. insert its plain text at the current cursor
 *        b. (between paragraphs) insert a paragraph break
 *        c. (optional) apply bold/italic/underline run formatting
 *
 * The helper is intentionally environment-agnostic: it accepts the HTML as a
 * string and uses the global `DOMParser` (available in browsers, jsdom, and
 * happy-dom) to parse it.  When run in a non-DOM environment the helper falls
 * back to a plain-text path.
 *
 * Scope (MVP):
 *   - paragraph blocks: <p>, <div>, <h1..h6>, <li>, <br>
 *   - inline formatting: <b>/<strong>, <i>/<em>, <u>
 *   - everything else collapses to its text content (whitespace normalised)
 *
 * Out of scope for E.5 (deferred): tables, lists with numbering, hyperlinks,
 * images, font/colour, paragraph alignment.
 */

import type { ParaProps, RunProps } from '../model'

import type { Command, Position, Range } from './commandTypes'

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

export interface ParsedRun {
  readonly text: string
  readonly format: Partial<RunProps>
}

export interface ParsedParagraph {
  readonly runs: ReadonlyArray<ParsedRun>
  readonly paraFormat?: Partial<ParaProps>
}

export interface PasteCommandsResult {
  readonly commands: ReadonlyArray<Command>
  readonly finalCursor: Position
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Parse a plain-text string into paragraphs.  Splits on `\r?\n`, drops a
 * trailing empty paragraph created by a terminal newline, and otherwise
 * preserves empty paragraphs (so a blank line stays a blank line).
 */
export function textToParagraphs(text: string): ReadonlyArray<ParsedParagraph> {
  if (text.length === 0) {
    return []
  }

  const lines = text.split(/\r?\n/)
  // A trailing newline in the source produces a terminal empty entry — drop
  // it so "foo\n" becomes one paragraph, not two.
  if (lines.length > 1 && lines[lines.length - 1] === '') {
    lines.pop()
  }

  return lines.map(line => ({
    runs: line.length === 0 ? [] : [{ text: line, format: {} }],
  }))
}

/**
 * Parse an HTML fragment into paragraphs with inline formatting.  When the
 * environment lacks `DOMParser` (or parsing fails) the helper falls back to
 * `textToParagraphs` on the stripped text content.
 */
export function htmlToParagraphs(html: string): ReadonlyArray<ParsedParagraph> {
  if (html.length === 0) {
    return []
  }

  const parser = getDomParser()
  if (parser === null) {
    return textToParagraphs(stripTags(html))
  }

  let doc: Document
  try {
    doc = parser.parseFromString(html, 'text/html')
  } catch {
    return textToParagraphs(stripTags(html))
  }

  const body = doc.body
  if (body === null) {
    return textToParagraphs(stripTags(html))
  }

  const builder = new ParagraphBuilder()
  walk(body, { bold: false, italic: false, underline: false }, builder)
  return builder.finish()
}

/**
 * Build the ordered command sequence to apply a parsed paste at `cursor`,
 * optionally replacing `selection` first.  Returns the command list plus the
 * final cursor position (where the caret should land after applying them).
 */
export function buildPasteCommands(
  paragraphs: ReadonlyArray<ParsedParagraph>,
  cursor: Position,
  selection: Range | null,
): PasteCommandsResult {
  const commands: Command[] = []

  let insertAt: Position = cursor
  if (selection !== null && !positionsEqual(selection.anchor, selection.focus)) {
    commands.push({ kind: 'delete-range', range: selection })
    insertAt = orderedStart(selection)
  }

  if (paragraphs.length === 0) {
    return { commands, finalCursor: insertAt }
  }

  let pos: Position = insertAt

  paragraphs.forEach((paragraph, paragraphIndex) => {
    if (paragraphIndex > 0) {
      commands.push({ kind: 'insert-paragraph-break', at: pos })
      pos = startOfNextParagraph(pos)
    }

    paragraph.runs.forEach(run => {
      if (run.text.length === 0) {
        return
      }

      const textStart: Position = pos
      commands.push({ kind: 'insert-text', at: textStart, text: run.text })

      const textEnd: Position = {
        paragraphPath: textStart.paragraphPath,
        runIndex: textStart.runIndex,
        charOffset: textStart.charOffset + run.text.length,
      }

      if (hasAnyFormat(run.format)) {
        commands.push({
          kind: 'apply-run-format',
          range: { anchor: textStart, focus: textEnd },
          format: run.format,
        })
      }

      pos = textEnd
    })
  })

  return { commands, finalCursor: pos }
}

// ---------------------------------------------------------------------------
// Internals — HTML walker
// ---------------------------------------------------------------------------

interface FormatState {
  readonly bold: boolean
  readonly italic: boolean
  readonly underline: boolean
}

const PARAGRAPH_TAGS = new Set([
  'P',
  'DIV',
  'LI',
  'H1',
  'H2',
  'H3',
  'H4',
  'H5',
  'H6',
  'BLOCKQUOTE',
  'PRE',
])

const SKIP_TAGS = new Set(['SCRIPT', 'STYLE', 'NOSCRIPT', 'TEMPLATE'])

class ParagraphBuilder {
  private readonly paragraphs: ParsedParagraph[] = []
  private currentRuns: ParsedRun[] = []

  appendText(text: string, state: FormatState): void {
    const normalised = text.replace(/\s+/g, ' ')
    if (normalised.length === 0) {
      return
    }

    const format = stateToFormat(state)
    const last = this.currentRuns[this.currentRuns.length - 1]
    if (last !== undefined && shallowFormatEqual(last.format, format)) {
      this.currentRuns[this.currentRuns.length - 1] = {
        text: last.text + normalised,
        format,
      }
      return
    }

    this.currentRuns.push({ text: normalised, format })
  }

  breakParagraph(): void {
    const trimmed = trimRuns(this.currentRuns)
    if (trimmed.length === 0) {
      // Suppress empty paragraphs that come from block boundaries when no
      // visible content has been accumulated yet (e.g. entering the first
      // <p>, or back-to-back block close/open).
      this.currentRuns = []
      return
    }
    this.paragraphs.push({ runs: trimmed })
    this.currentRuns = []
  }

  finish(): ReadonlyArray<ParsedParagraph> {
    if (this.currentRuns.length > 0) {
      this.paragraphs.push({ runs: trimRuns(this.currentRuns) })
      this.currentRuns = []
    }
    return this.paragraphs
  }
}

function walk(node: Node, state: FormatState, builder: ParagraphBuilder): void {
  if (node.nodeType === 3 /* TEXT_NODE */) {
    const text = node.nodeValue ?? ''
    builder.appendText(text, state)
    return
  }

  if (node.nodeType !== 1 /* ELEMENT_NODE */) {
    return
  }

  const element = node as Element
  const tag = element.tagName.toUpperCase()

  if (SKIP_TAGS.has(tag)) {
    return
  }

  if (tag === 'BR') {
    builder.breakParagraph()
    return
  }

  const isBlock = PARAGRAPH_TAGS.has(tag)
  const nextState: FormatState = {
    bold: state.bold || tag === 'B' || tag === 'STRONG',
    italic: state.italic || tag === 'I' || tag === 'EM',
    underline: state.underline || tag === 'U',
  }

  if (isBlock) {
    // Close any in-progress paragraph before entering a new block-level
    // element so its content starts on a fresh paragraph.
    builder.breakParagraph()
  }

  const children = element.childNodes
  for (let i = 0; i < children.length; i += 1) {
    const child = children.item(i)
    if (child !== null) {
      walk(child, nextState, builder)
    }
  }

  if (isBlock) {
    builder.breakParagraph()
  }
}

function stateToFormat(state: FormatState): Partial<RunProps> {
  const format: { -readonly [K in keyof RunProps]?: RunProps[K] } = {}
  if (state.bold) {
    format.bold = true
  }
  if (state.italic) {
    format.italic = true
  }
  if (state.underline) {
    format.underline = { style: 'single' }
  }
  return format
}

function shallowFormatEqual(a: Partial<RunProps>, b: Partial<RunProps>): boolean {
  return (
    Boolean(a.bold) === Boolean(b.bold) &&
    Boolean(a.italic) === Boolean(b.italic) &&
    Boolean(a.underline) === Boolean(b.underline)
  )
}

function trimRuns(runs: ReadonlyArray<ParsedRun>): ReadonlyArray<ParsedRun> {
  // Drop pure-whitespace runs at the paragraph boundaries that sneak in from
  // tag-internal whitespace.  Preserve internal runs verbatim.
  const start = runs.findIndex(run => run.text.trim().length > 0)
  if (start === -1) {
    return []
  }

  let end = runs.length - 1
  while (end > start && runs[end].text.trim().length === 0) {
    end -= 1
  }

  const sliced = runs.slice(start, end + 1)
  // Trim leading whitespace on the first run and trailing whitespace on the
  // last run, mirroring browser paste behaviour.
  const head = sliced[0]
  const tail = sliced[sliced.length - 1]
  const trimmed: ParsedRun[] = sliced.map(r => ({ ...r }))
  trimmed[0] = { ...head, text: head.text.replace(/^\s+/, '') }
  trimmed[trimmed.length - 1] = {
    ...tail,
    text: tail.text.replace(/\s+$/, ''),
  }
  return trimmed.filter(run => run.text.length > 0)
}

// ---------------------------------------------------------------------------
// Internals — misc
// ---------------------------------------------------------------------------

function hasAnyFormat(format: Partial<RunProps>): boolean {
  return Object.keys(format).length > 0
}

function positionsEqual(a: Position, b: Position): boolean {
  if (a.runIndex !== b.runIndex || a.charOffset !== b.charOffset) {
    return false
  }
  if (a.paragraphPath.length !== b.paragraphPath.length) {
    return false
  }
  for (let i = 0; i < a.paragraphPath.length; i += 1) {
    if (a.paragraphPath[i] !== b.paragraphPath[i]) {
      return false
    }
  }
  return true
}

function comparePositions(a: Position, b: Position): number {
  const len = Math.max(a.paragraphPath.length, b.paragraphPath.length)
  for (let i = 0; i < len; i += 1) {
    const av = a.paragraphPath[i] ?? -1
    const bv = b.paragraphPath[i] ?? -1
    if (av !== bv) {
      return av - bv
    }
  }
  if (a.runIndex !== b.runIndex) {
    return a.runIndex - b.runIndex
  }
  return a.charOffset - b.charOffset
}

function orderedStart(range: Range): Position {
  return comparePositions(range.anchor, range.focus) <= 0 ? range.anchor : range.focus
}

function startOfNextParagraph(pos: Position): Position {
  const path = [...pos.paragraphPath]
  if (path.length === 0) {
    return { paragraphPath: Object.freeze([0]), runIndex: 0, charOffset: 0 }
  }
  path[path.length - 1] = (path[path.length - 1] ?? 0) + 1
  return {
    paragraphPath: Object.freeze(path),
    runIndex: 0,
    charOffset: 0,
  }
}

function getDomParser(): DOMParser | null {
  if (typeof DOMParser === 'undefined') {
    return null
  }
  return new DOMParser()
}

function stripTags(html: string): string {
  return html
    .replace(/<\s*br\s*\/?\s*>/gi, '\n')
    .replace(/<\/(p|div|li|h[1-6]|blockquote|pre)>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
}
