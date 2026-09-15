/**
 * Atlas — rich paste-from-Word/HTML parsing (DXE-19)
 *
 * Parses pasted HTML into an ordered sequence of `PasteBlock`s (paragraphs
 * and tables, interleaved in document order), extending htmlPaste.ts's MVP
 * paragraph/bold/italic/underline walker with:
 *   - tables, including horizontally merged cells (`colspan`) — vertical
 *     merge (`rowspan`) is a documented remaining gap: a row-spanning cell
 *     is parsed as an ordinary single-row cell, so the rows below it come
 *     out one cell short rather than reconstructed as a `vMerge` group.
 *   - hyperlinks (`<a href>`), restricted to http(s)/mailto — the same
 *     scheme allowlist `PageView.tsx` already applies when *rendering* a
 *     hyperlink, applied here at parse time instead so a `javascript:`/
 *     `data:` href from untrusted clipboard HTML never reaches the model.
 *   - inline text color / highlight, from `style="color:...;
 *     background-color:..."` or a legacy `<font color>` attribute — an
 *     unrecognized background-color (outside OOXML's fixed `w:highlight`
 *     enum) falls back to `shd` (arbitrary-color shading), which is closer
 *     to what the pasted color actually looked like than dropping it.
 *   - lists (`<ul>`/`<ol>`/`<li>`), tracking nesting depth for fidelity even
 *     though `pasteRich.ts`'s `ensureListNumbering` integration (matching
 *     the toolbar's own list support) currently only defines level 0 — a
 *     paste of a nested list still marks every level as a list item, just
 *     without a distinct indent/marker per level yet.
 *   - inline images, `data:image/{png,jpeg,gif}` sources only. An `http(s)`
 *     or `file:` `<img src>` is intentionally never read: that would mean
 *     this editor reaching out to the filesystem or network on behalf of
 *     untrusted clipboard HTML, which `insertImage.ts`'s own
 *     file-picker-only design deliberately avoids elsewhere too. Real
 *     Word/browser copy already inlines the image as a `data:` URL for
 *     exactly this content (as opposed to a *link to* an image), so this
 *     covers the common paste case.
 *
 * A `<table>` nested inside a pasted table's cell is dropped (its text
 * content is not flattened in either) — a documented gap alongside vertical
 * merge, rather than attempted here.
 */

import { hexColor, type RunProps } from '../model'

import { parseCssColor, toHighlightColor } from './colorMapping'

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

export interface PasteListInfo {
  readonly ordered: boolean
  /** Nesting depth, 0 for a top-level `<li>`. Parsed for fidelity but not
   * yet reflected in the actual list indent — see the module doc comment. */
  readonly level: number
}

export interface PasteTextRun {
  readonly kind: 'text'
  readonly text: string
  readonly format: Partial<RunProps>
  /** Set only for a run inside an `<a href>` with a safe (http/https/
   * mailto) scheme. */
  readonly href?: string
}

export interface PasteImageRun {
  readonly kind: 'image'
  /** A `data:image/{png,jpeg,gif};base64,...` URL — see the module doc
   * comment for why other `<img src>` schemes never reach this far. */
  readonly dataUrl: string
}

export type PasteRun = PasteTextRun | PasteImageRun

export interface PasteParagraph {
  readonly kind: 'paragraph'
  readonly runs: ReadonlyArray<PasteRun>
  readonly list?: PasteListInfo
}

export interface PasteTableCell {
  readonly paragraphs: ReadonlyArray<PasteParagraph>
  /** From the source `colspan` attribute (default 1, clamped to >= 1). */
  readonly colSpan: number
}

export interface PasteTableRow {
  readonly cells: ReadonlyArray<PasteTableCell>
}

export interface PasteTable {
  readonly kind: 'table'
  readonly rows: ReadonlyArray<PasteTableRow>
}

export type PasteBlock = PasteParagraph | PasteTable

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Parses an HTML fragment into paste blocks. Falls back to a single
 * plain-text paragraph sequence (via the caller's own `text/plain` clipboard
 * data being the more appropriate source in that case — this function
 * itself just returns `[]`) when `DOMParser` is unavailable or parsing
 * fails, mirroring `htmlPaste.ts`'s own fallback contract.
 */
export function htmlToPasteBlocks(html: string): ReadonlyArray<PasteBlock> {
  if (html.length === 0) {
    return []
  }

  const parser = getDomParser()
  if (parser === null) {
    return []
  }

  let doc: Document
  try {
    doc = parser.parseFromString(html, 'text/html')
  } catch {
    return []
  }

  if (doc.body === null) {
    return []
  }

  const builder = new PasteBlockBuilder()
  walkNode(doc.body, { format: {} }, builder)
  return builder.finish()
}

// ---------------------------------------------------------------------------
// Internals — safety
// ---------------------------------------------------------------------------

const SAFE_HYPERLINK_SCHEMES: ReadonlySet<string> = new Set(['http:', 'https:', 'mailto:'])

function isSafeHyperlinkHref(href: string): boolean {
  try {
    return SAFE_HYPERLINK_SCHEMES.has(new URL(href).protocol)
  } catch {
    return false
  }
}

const DATA_IMAGE_RE = /^data:image\/(png|jpe?g|gif)(?:;[^,]*)?;base64,/i

function isSupportedImageDataUrl(src: string): boolean {
  return DATA_IMAGE_RE.test(src)
}

// ---------------------------------------------------------------------------
// Internals — color parsing
// ---------------------------------------------------------------------------

function parseInlineStyle(styleAttr: string | null): Map<string, string> {
  const declarations = new Map<string, string>()
  if (styleAttr === null) {
    return declarations
  }

  for (const rule of styleAttr.split(';')) {
    const colonIndex = rule.indexOf(':')
    if (colonIndex === -1) {
      continue
    }
    const property = rule.slice(0, colonIndex).trim().toLowerCase()
    const value = rule.slice(colonIndex + 1).trim()
    if (property.length > 0 && value.length > 0) {
      declarations.set(property, value)
    }
  }

  return declarations
}

// ---------------------------------------------------------------------------
// Internals — walker state
// ---------------------------------------------------------------------------

interface WalkState {
  readonly format: Partial<RunProps>
  readonly href?: string
  readonly list?: PasteListInfo
}

function withFormat(state: WalkState, patch: Partial<RunProps>): WalkState {
  return { ...state, format: { ...state.format, ...patch } }
}

const PARAGRAPH_TAGS: ReadonlySet<string> = new Set([
  'P',
  'DIV',
  'H1',
  'H2',
  'H3',
  'H4',
  'H5',
  'H6',
  'BLOCKQUOTE',
  'PRE',
])

const SKIP_TAGS: ReadonlySet<string> = new Set(['SCRIPT', 'STYLE', 'NOSCRIPT', 'TEMPLATE'])

// ---------------------------------------------------------------------------
// Internals — paragraph/block builder
// ---------------------------------------------------------------------------

class PasteBlockBuilder {
  private readonly blocks: PasteBlock[] = []
  private currentRuns: PasteRun[] = []
  private currentList: PasteListInfo | undefined

  setListContext(list: PasteListInfo | undefined): void {
    this.currentList = list
  }

  appendText(text: string, state: WalkState): void {
    const normalised = text.replace(/\s+/g, ' ')
    if (normalised.length === 0) {
      return
    }

    const last = this.currentRuns[this.currentRuns.length - 1]
    if (
      last !== undefined &&
      last.kind === 'text' &&
      shallowFormatEqual(last.format, state.format) &&
      last.href === state.href
    ) {
      this.currentRuns[this.currentRuns.length - 1] = { ...last, text: last.text + normalised }
      return
    }

    this.currentRuns.push({
      kind: 'text',
      text: normalised,
      format: state.format,
      ...(state.href !== undefined ? { href: state.href } : {}),
    })
  }

  appendImage(dataUrl: string): void {
    this.currentRuns.push({ kind: 'image', dataUrl })
  }

  breakParagraph(): void {
    const trimmed = trimRuns(this.currentRuns)
    this.currentRuns = []
    if (trimmed.length === 0) {
      return
    }

    this.blocks.push({
      kind: 'paragraph',
      runs: trimmed,
      ...(this.currentList !== undefined ? { list: this.currentList } : {}),
    })
  }

  pushTable(table: PasteTable): void {
    this.breakParagraph()
    if (table.rows.length > 0) {
      this.blocks.push(table)
    }
  }

  finish(): ReadonlyArray<PasteBlock> {
    this.breakParagraph()
    return this.blocks
  }
}

function shallowFormatEqual(a: Partial<RunProps>, b: Partial<RunProps>): boolean {
  return (
    Boolean(a.bold) === Boolean(b.bold) &&
    Boolean(a.italic) === Boolean(b.italic) &&
    Boolean(a.underline) === Boolean(b.underline) &&
    a.color === b.color &&
    a.highlight === b.highlight &&
    JSON.stringify(a.shd) === JSON.stringify(b.shd)
  )
}

function trimRuns(runs: ReadonlyArray<PasteRun>): ReadonlyArray<PasteRun> {
  const isBlankText = (run: PasteRun): boolean => run.kind === 'text' && run.text.trim().length === 0

  const start = runs.findIndex((run) => !isBlankText(run))
  if (start === -1) {
    return []
  }

  let end = runs.length - 1
  while (end > start && isBlankText(runs[end])) {
    end -= 1
  }

  const sliced = runs.slice(start, end + 1).map((run) => ({ ...run }))
  const head = sliced[0]
  if (head.kind === 'text') {
    sliced[0] = { ...head, text: head.text.replace(/^\s+/, '') }
  }
  const tail = sliced[sliced.length - 1]
  if (tail.kind === 'text') {
    sliced[sliced.length - 1] = { ...tail, text: tail.text.replace(/\s+$/, '') }
  }

  return sliced.filter((run) => run.kind === 'image' || run.text.length > 0)
}

// ---------------------------------------------------------------------------
// Internals — DOM walk
// ---------------------------------------------------------------------------

function applyElementColors(element: Element, state: WalkState): WalkState {
  const style = parseInlineStyle(element.getAttribute('style'))
  let next = state

  const colorValue = style.get('color')
  if (colorValue !== undefined) {
    const hex = parseCssColor(colorValue)
    if (hex !== null) {
      next = withFormat(next, { color: hexColor(hex) })
    }
  } else if (element.tagName === 'FONT') {
    const attr = element.getAttribute('color')
    if (attr !== null) {
      const hex = parseCssColor(attr.startsWith('#') ? attr : `#${attr}`)
      if (hex !== null) {
        next = withFormat(next, { color: hexColor(hex) })
      }
    }
  }

  const bgValue = style.get('background-color')
  if (bgValue !== undefined) {
    const hex = parseCssColor(bgValue)
    if (hex !== null) {
      const highlight = toHighlightColor(hex)
      next =
        highlight !== null
          ? withFormat(next, { highlight })
          : withFormat(next, { shd: { fill: hexColor(hex) } })
    }
  }

  return next
}

function walkNode(node: Node, state: WalkState, builder: PasteBlockBuilder): void {
  if (node.nodeType === 3 /* TEXT_NODE */) {
    builder.appendText(node.nodeValue ?? '', state)
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

  if (tag === 'IMG') {
    const src = element.getAttribute('src')
    if (src !== null && isSupportedImageDataUrl(src)) {
      builder.appendImage(src)
    }
    return
  }

  if (tag === 'TABLE') {
    const table = parseTableElement(element, state)
    if (table !== null) {
      builder.pushTable(table)
    }
    return
  }

  if (tag === 'UL' || tag === 'OL') {
    const nextState: WalkState = {
      ...state,
      list: { ordered: tag === 'OL', level: (state.list?.level ?? -1) + 1 },
    }
    walkChildren(element, nextState, builder)
    return
  }

  if (tag === 'LI') {
    builder.breakParagraph()
    builder.setListContext(state.list)
    walkChildren(element, state, builder)
    builder.breakParagraph()
    builder.setListContext(undefined)
    return
  }

  let nextState = state
  if (tag === 'B' || tag === 'STRONG') {
    nextState = withFormat(nextState, { bold: true })
  }
  if (tag === 'I' || tag === 'EM') {
    nextState = withFormat(nextState, { italic: true })
  }
  if (tag === 'U') {
    nextState = withFormat(nextState, { underline: { style: 'single' } })
  }
  if (tag === 'A') {
    const href = element.getAttribute('href')
    if (href !== null && isSafeHyperlinkHref(href)) {
      nextState = { ...nextState, href }
    }
  }
  nextState = applyElementColors(element, nextState)

  const isBlock = PARAGRAPH_TAGS.has(tag)
  if (isBlock) {
    builder.breakParagraph()
  }

  walkChildren(element, nextState, builder)

  if (isBlock) {
    builder.breakParagraph()
  }
}

function walkChildren(element: Element, state: WalkState, builder: PasteBlockBuilder): void {
  const children = element.childNodes
  for (let i = 0; i < children.length; i += 1) {
    const child = children.item(i)
    if (child !== null) {
      walkNode(child, state, builder)
    }
  }
}

// ---------------------------------------------------------------------------
// Internals — table parsing
// ---------------------------------------------------------------------------

function collectTableRowElements(tableEl: Element): ReadonlyArray<Element> {
  const rows: Element[] = []

  const collectFrom = (parent: Element): void => {
    for (const child of Array.from(parent.children)) {
      if (child.tagName === 'TR') {
        rows.push(child)
      }
    }
  }

  for (const child of Array.from(tableEl.children)) {
    if (child.tagName === 'TR') {
      rows.push(child)
    } else if (child.tagName === 'THEAD' || child.tagName === 'TBODY' || child.tagName === 'TFOOT') {
      collectFrom(child)
    }
  }

  return rows
}

function parseTableCellContent(cellEl: Element, state: WalkState): ReadonlyArray<PasteParagraph> {
  const builder = new PasteBlockBuilder()
  walkChildren(cellEl, state, builder)
  return builder.finish().filter((block): block is PasteParagraph => block.kind === 'paragraph')
}

function parseTableElement(tableEl: Element, state: WalkState): PasteTable | null {
  const rowElements = collectTableRowElements(tableEl)
  if (rowElements.length === 0) {
    return null
  }

  const rows: PasteTableRow[] = []
  for (const rowEl of rowElements) {
    const cellElements = Array.from(rowEl.children).filter(
      (child) => child.tagName === 'TD' || child.tagName === 'TH',
    )
    if (cellElements.length === 0) {
      continue
    }

    const cells: PasteTableCell[] = cellElements.map((cellEl) => {
      const colSpanAttr = cellEl.getAttribute('colspan')
      const parsedSpan = colSpanAttr === null ? 1 : Number.parseInt(colSpanAttr, 10)
      const colSpan = Number.isFinite(parsedSpan) && parsedSpan > 0 ? parsedSpan : 1
      const paragraphs = parseTableCellContent(cellEl, state)
      return {
        paragraphs: paragraphs.length > 0 ? paragraphs : [{ kind: 'paragraph', runs: [] }],
        colSpan,
      }
    })

    rows.push({ cells })
  }

  return rows.length > 0 ? { kind: 'table', rows } : null
}

// ---------------------------------------------------------------------------
// Internals — misc
// ---------------------------------------------------------------------------

function getDomParser(): DOMParser | null {
  if (typeof DOMParser === 'undefined') {
    return null
  }
  return new DOMParser()
}
