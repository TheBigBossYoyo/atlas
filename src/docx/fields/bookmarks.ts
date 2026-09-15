/**
 * Atlas — bookmark text + page resolution for REF/PAGEREF fields
 * (DEFER-5 / DXS-20)
 *
 * `evaluateFieldText`'s REF/PAGEREF cases (see `./evaluate.ts`) consume a
 * `FieldEvaluationContext.bookmarkText`/`bookmarkPage` map that has to be
 * built from the actual document somewhere: `collectBookmarkMaps` walks the
 * whole document once, pairing each `w:bookmarkStart`/`w:bookmarkEnd` by
 * `id` (a bookmark's name lives only on its `start` boundary — see
 * `Bookmark`'s own doc comment) and collecting the plain text that falls
 * between them. This covers a bookmark spanning multiple paragraphs, nested
 * inside a hyperlink, or wrapping a field's own cached result (a field
 * nested inside a bookmarked range contributes its cached display text,
 * the same recursive treatment `paragraphText.ts` gives a field for TOC
 * heading text).
 *
 * A bookmark's PAGEREF page number is the page its START boundary falls
 * on (`pageOf`, when supplied — the same `(sectionIndex, blockIndex) ->
 * page` resolver `toc.ts`/`DocxViewer` build from pagination output),
 * looked up only for a bookmark whose start sits directly in a body
 * section's own paragraph list. A bookmark nested inside a table cell (or
 * one that lives in a header/footer/footnote/endnote, walked here for TEXT
 * only) has no single well-defined body page — same reasoning as
 * `updateFields`'s table-cell/header PAGE-field handling — so it gets a
 * `bookmarkText` entry but no `bookmarkPage` one.
 *
 * A bookmark left open with no matching end (malformed source) never
 * reaches `finished` and so is simply absent from both maps — the same
 * honest "leave the field's cached result unchanged" fallback
 * `evaluateRefField`/`evaluatePageRefField` already give a bookmark name
 * with no entry at all.
 */
import type {
  Block,
  Document,
  HyperlinkChild,
  ParagraphChild,
  Run,
  RunChild,
  TableChild,
  TableRowChild,
} from '../model'

import type { TocPageResolver } from './toc'

export interface BookmarkMaps {
  readonly bookmarkText: ReadonlyMap<string, string>
  readonly bookmarkPage: ReadonlyMap<string, number>
}

/** A body paragraph's address, for PAGEREF page lookup — `undefined` outside a body section (see this module's doc comment). */
interface BodyPath {
  readonly sectionIndex: number
  readonly blockIndex: number
}

interface OpenBookmark {
  readonly name: string
  readonly textParts: string[]
  readonly page: number | undefined
}

interface WalkState {
  readonly open: Map<string, OpenBookmark>
  readonly finishedText: Map<string, string>
  readonly finishedPage: Map<string, number>
}

export function collectBookmarkMaps(document: Document, pageOf?: TocPageResolver): BookmarkMaps {
  const state: WalkState = { open: new Map(), finishedText: new Map(), finishedPage: new Map() }

  document.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      walkBlock(block, state, { sectionIndex, blockIndex }, pageOf)
    })
  })

  for (const header of document.headers.values()) {
    walkBlocks(header.blocks, state, undefined, pageOf)
  }
  for (const footer of document.footers.values()) {
    walkBlocks(footer.blocks, state, undefined, pageOf)
  }
  for (const footnote of document.footnotes.values()) {
    walkBlocks(footnote.blocks, state, undefined, pageOf)
  }
  for (const endnote of document.endnotes.values()) {
    walkBlocks(endnote.blocks, state, undefined, pageOf)
  }

  return { bookmarkText: state.finishedText, bookmarkPage: state.finishedPage }
}

function walkBlocks(
  blocks: ReadonlyArray<Block>,
  state: WalkState,
  path: BodyPath | undefined,
  pageOf: TocPageResolver | undefined,
): void {
  for (const block of blocks) {
    walkBlock(block, state, path, pageOf)
  }
}

function walkBlock(
  block: Block,
  state: WalkState,
  path: BodyPath | undefined,
  pageOf: TocPageResolver | undefined,
): void {
  if (block.kind === 'paragraph') {
    for (const child of block.children) {
      walkParagraphChild(child, state, path, pageOf)
    }
    return
  }

  if (block.kind === 'table') {
    for (const row of block.rows) {
      walkTableRow(row, state, pageOf)
    }
  }
}

function walkTableRow(row: TableChild, state: WalkState, pageOf: TocPageResolver | undefined): void {
  if (row.kind !== 'table-row') {
    return
  }
  for (const cell of row.cells) {
    walkTableCell(cell, state, pageOf)
  }
}

function walkTableCell(cell: TableRowChild, state: WalkState, pageOf: TocPageResolver | undefined): void {
  if (cell.kind !== 'table-cell') {
    return
  }
  // A table cell's paragraphs have no body-section page address of their
  // own — same reasoning as `updateFields`'s table-cell PAGE handling.
  walkBlocks(cell.blocks, state, undefined, pageOf)
}

function walkParagraphChild(
  child: ParagraphChild,
  state: WalkState,
  path: BodyPath | undefined,
  pageOf: TocPageResolver | undefined,
): void {
  switch (child.kind) {
    case 'run':
      appendText(state, runPlainText(child))
      return
    case 'hyperlink':
      for (const hyperlinkChild of child.children) {
        walkHyperlinkChild(hyperlinkChild, state, path, pageOf)
      }
      return
    case 'field':
      for (const resultChild of child.result) {
        walkParagraphChild(resultChild, state, path, pageOf)
      }
      return
    case 'bookmark':
      handleBookmark(child.boundary, child.id, child.name, state, path, pageOf)
      return
    case 'ins-revision':
      for (const revisionChild of child.children as ReadonlyArray<ParagraphChild>) {
        walkParagraphChild(revisionChild, state, path, pageOf)
      }
      return
    default:
      // 'del-revision' (tracked-deleted text isn't visible content, matching
      // `paragraphText.ts`), 'comment-range'/'comment-reference'/
      // 'footnote-reference'/'endnote-reference'/'unknown': no text, and
      // none of these can themselves carry a nested bookmark boundary.
      return
  }
}

function walkHyperlinkChild(
  child: HyperlinkChild,
  state: WalkState,
  path: BodyPath | undefined,
  pageOf: TocPageResolver | undefined,
): void {
  switch (child.kind) {
    case 'run':
      appendText(state, runPlainText(child))
      return
    case 'bookmark':
      handleBookmark(child.boundary, child.id, child.name, state, path, pageOf)
      return
    case 'field':
      for (const resultChild of child.result) {
        walkParagraphChild(resultChild, state, path, pageOf)
      }
      return
    default:
      return
  }
}

function handleBookmark(
  boundary: 'start' | 'end',
  id: string,
  name: string | undefined,
  state: WalkState,
  path: BodyPath | undefined,
  pageOf: TocPageResolver | undefined,
): void {
  if (boundary === 'start') {
    if (name === undefined) {
      return
    }
    const page = path !== undefined ? pageOf?.(path.sectionIndex, path.blockIndex) : undefined
    state.open.set(id, { name, textParts: [], page })
    return
  }

  const entry = state.open.get(id)
  if (entry === undefined) {
    return
  }
  state.open.delete(id)

  // First occurrence wins for a (theoretically invalid, but never trusted
  // to be impossible) duplicate bookmark name — Word itself doesn't allow
  // creating a second bookmark under a name already in use.
  if (state.finishedText.has(entry.name)) {
    return
  }
  state.finishedText.set(entry.name, entry.textParts.join(''))
  if (entry.page !== undefined) {
    state.finishedPage.set(entry.name, entry.page)
  }
}

function appendText(state: WalkState, text: string): void {
  if (text === '' || state.open.size === 0) {
    return
  }
  for (const entry of state.open.values()) {
    entry.textParts.push(text)
  }
}

function runPlainText(run: Run): string {
  return run.children.map(runChildText).join('')
}

function runChildText(child: RunChild): string {
  switch (child.kind) {
    case 'text':
      return child.value
    case 'tab':
      return '\t'
    default:
      return ''
  }
}
