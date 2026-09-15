/**
 * Atlas — table of contents generation + "Update table of contents"
 * command (DEFER-5 / DXS-20)
 *
 * A TOC field's cached result is fundamentally different from every other
 * field type's: an ordinary field caches one text value, but a TOC caches
 * a whole list of heading-derived entries (each its own line, optionally
 * hyperlinked, with a page number) — this module's `collectTocEntries`
 * (the pure generator) and `updateTableOfContents` (the command wiring it
 * into a `Document`) are deliberately separate from `updateFields.ts`,
 * which explicitly skips `TOC` fields and defers to this module instead.
 *
 * Scope, documented rather than silently approximated:
 *  - Heading detection uses each paragraph's CASCADE-RESOLVED
 *    `ParaProps.outlineLvl` (direct formatting, else its style chain) —
 *    this covers both a `\o`-style "built-in Heading styles" TOC and a
 *    `\u`-style "any paragraph with an applied outline level" TOC
 *    identically, since Atlas's model doesn't distinguish "outline level
 *    came from style vs. direct formatting" after cascade resolution. Real
 *    Word can in principle produce different entries for `\o` vs `\u` on a
 *    document mixing style-based and direct-formatting outline levels in
 *    unusual ways; this scope treats them the same.
 *  - `\z` (hide page numbers/leaders in web layout) has no effect: Atlas
 *    doesn't have a distinct web-layout TOC rendering mode.
 *  - `\h` hyperlinks an entry only when the heading paragraph itself starts
 *    with a bookmark (real Word auto-generates a `_Toc*` bookmark at each
 *    heading when `\h` is used — Atlas doesn't synthesize one for a
 *    heading that lacks it, so such an entry renders as plain text
 *    instead of a broken link).
 *  - The regenerated TOC is written back as ONE paragraph (the `TOC`
 *    `Field`'s own `result`, entries separated by line breaks) rather than
 *    Word's usual one-`TOCn`-styled-paragraph-per-entry layout spanning
 *    many paragraphs — matching how this branch's field model represents
 *    every field (as one `ParagraphChild`, not a multi-paragraph span).
 *    `findTocField` therefore only ever locates a single-paragraph TOC
 *    field; a real Word-authored multi-paragraph TOC isn't modeled as a
 *    `Field` at all (its begin/instrText runs have no matching `end`
 *    within the same paragraph, so `groupComplexFieldRuns` leaves them as
 *    ordinary run content — preserved on save, just not recognized as a
 *    field to update, an honest "not yet supported" rather than data
 *    loss).
 */
import type { Block, Document, Field, Paragraph, ParagraphChild, RunChild } from '../model'
import { resolveParaProps } from '../parser/cascade'

import { parseFieldInstruction } from './instruction'
import { paragraphPlainText } from './paragraphText'

export interface TocOptions {
  /** 1-based, inclusive (a `Heading1` paragraph is level 1). */
  readonly minLevel: number
  readonly maxLevel: number
  readonly hyperlink: boolean
}

const DEFAULT_MIN_LEVEL = 1
const DEFAULT_MAX_LEVEL = 3

export function parseTocOptions(instruction: string): TocOptions {
  const parsed = parseFieldInstruction(instruction)
  const range = parsed.switches.get('o')
  const { minLevel, maxLevel } = typeof range === 'string' ? parseLevelRange(range) : defaultLevelRange()

  return {
    minLevel,
    maxLevel,
    hyperlink: parsed.switches.has('h'),
  }
}

function defaultLevelRange(): { minLevel: number; maxLevel: number } {
  return { minLevel: DEFAULT_MIN_LEVEL, maxLevel: DEFAULT_MAX_LEVEL }
}

function parseLevelRange(range: string): { minLevel: number; maxLevel: number } {
  const trimmed = range.trim()
  const rangeMatch = /^(\d+)\s*-\s*(\d+)$/.exec(trimmed)
  if (rangeMatch !== null) {
    return { minLevel: Number.parseInt(rangeMatch[1], 10), maxLevel: Number.parseInt(rangeMatch[2], 10) }
  }

  const singleMatch = /^(\d+)$/.exec(trimmed)
  if (singleMatch !== null) {
    return { minLevel: DEFAULT_MIN_LEVEL, maxLevel: Number.parseInt(singleMatch[1], 10) }
  }

  return defaultLevelRange()
}

export interface TocEntry {
  /** 1-based heading level. */
  readonly level: number
  readonly text: string
  /** The heading's own leading bookmark name, if it has one — see this module's doc comment on `\h` scope. */
  readonly bookmarkName?: string
  readonly pageNumber?: number
}

/**
 * Resolves a body paragraph's page for a TOC entry. `sectionIndex`/
 * `blockIndex` address it the same way `FieldEvaluationContext.currentPageOf`
 * addresses a `PAGE` field's paragraph (`[sectionIndex, blockIndex]`) —
 * callers can share one implementation for both.
 */
export type TocPageResolver = (sectionIndex: number, blockIndex: number) => number | undefined

export function collectTocEntries(
  document: Document,
  options: TocOptions,
  pageOf?: TocPageResolver,
): ReadonlyArray<TocEntry> {
  const docDefaults = { pPr: document.defaults?.paragraph }
  const entries: TocEntry[] = []

  document.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      const entry = tocEntryForBlock(block, options, document, docDefaults)
      if (entry === undefined) {
        return
      }
      const pageNumber = pageOf?.(sectionIndex, blockIndex)
      entries.push(pageNumber !== undefined ? { ...entry, pageNumber } : entry)
    })
  })

  return entries
}

function tocEntryForBlock(
  block: Block,
  options: TocOptions,
  document: Document,
  docDefaults: { pPr?: Paragraph['props'] },
): TocEntry | undefined {
  if (block.kind !== 'paragraph') {
    return undefined
  }

  const resolved = resolveParaProps(block.props, block.props?.pStyle, document.styles, docDefaults)
  if (resolved.outlineLvl === undefined) {
    return undefined
  }

  const level = resolved.outlineLvl + 1
  if (level < options.minLevel || level > options.maxLevel) {
    return undefined
  }

  const text = paragraphPlainText(block).trim()
  if (text === '') {
    return undefined
  }

  const bookmarkName = leadingBookmarkName(block)
  return {
    level,
    text,
    ...(bookmarkName !== undefined ? { bookmarkName } : {}),
  }
}

function leadingBookmarkName(paragraph: Paragraph): string | undefined {
  for (const child of paragraph.children) {
    if (child.kind === 'bookmark' && child.boundary === 'start' && child.name !== undefined) {
      return child.name
    }
  }
  return undefined
}

/**
 * Renders entries into one field's worth of `ParagraphChild` content — see
 * this module's doc comment on why a single paragraph rather than Word's
 * usual one-paragraph-per-entry layout. Nesting is a plain leading-tab
 * indent per level below 1 (a readable approximation, not Word's real
 * TOCn-style-driven indent amounts) and each line is `text` + a tab +
 * the page number (or nothing, when no page is known yet).
 */
export function renderTocFieldResult(
  entries: ReadonlyArray<TocEntry>,
  options: TocOptions,
): ReadonlyArray<ParagraphChild> {
  const children: ParagraphChild[] = []

  entries.forEach((entry, index) => {
    if (index > 0) {
      children.push({ kind: 'run', children: [{ kind: 'break', breakType: 'line' }] })
    }

    const indent = '\t'.repeat(Math.max(0, entry.level - 1))
    const lineChildren: RunChild[] = [
      { kind: 'text', value: `${indent}${entry.text}` },
      { kind: 'tab' },
      { kind: 'text', value: entry.pageNumber !== undefined ? String(entry.pageNumber) : '' },
    ]

    if (options.hyperlink && entry.bookmarkName !== undefined) {
      children.push({
        kind: 'hyperlink',
        anchor: entry.bookmarkName,
        children: [{ kind: 'run', children: lineChildren }],
      })
    } else {
      children.push({ kind: 'run', children: lineChildren })
    }
  })

  return children
}

interface TocFieldLocation {
  readonly sectionIndex: number
  readonly blockIndex: number
  readonly childIndex: number
  readonly field: Field
}

function findTocField(document: Document): TocFieldLocation | undefined {
  for (let sectionIndex = 0; sectionIndex < document.sections.length; sectionIndex += 1) {
    const blocks = document.sections[sectionIndex]?.blocks ?? []
    for (let blockIndex = 0; blockIndex < blocks.length; blockIndex += 1) {
      const block = blocks[blockIndex]
      if (block === undefined || block.kind !== 'paragraph') {
        continue
      }
      for (let childIndex = 0; childIndex < block.children.length; childIndex += 1) {
        const child = block.children[childIndex]
        if (child !== undefined && child.kind === 'field' && child.fieldType === 'TOC') {
          return { sectionIndex, blockIndex, childIndex, field: child }
        }
      }
    }
  }
  return undefined
}

export interface UpdateTableOfContentsResult {
  readonly document: Document
  /** `false` when the document has no (single-paragraph, see this module's doc comment) TOC field to update — `document` is returned unchanged. */
  readonly updated: boolean
}

export function updateTableOfContents(
  document: Document,
  pageOf?: TocPageResolver,
): UpdateTableOfContentsResult {
  const location = findTocField(document)
  if (location === undefined) {
    return { document, updated: false }
  }
  // `w:fldLock` (`Field.locked`): the same "author explicitly froze this
  // field" reasoning `updateFields.ts` honors for every other field type
  // applies here too — Word's own "Update Table of Contents" leaves a
  // locked TOC field's entries untouched rather than regenerating them.
  if (location.field.locked === true) {
    return { document, updated: false }
  }

  const options = parseTocOptions(location.field.instruction)
  const entries = collectTocEntries(document, options, pageOf)
  const result = renderTocFieldResult(entries, options)
  const updatedField: Field = { ...location.field, result, raw: undefined }

  const sections = document.sections.map((section, sectionIndex) => {
    if (sectionIndex !== location.sectionIndex) {
      return section
    }
    return { ...section, blocks: replaceFieldAt(section.blocks, location.blockIndex, location.childIndex, updatedField) }
  })

  return { document: { ...document, sections }, updated: true }
}

function replaceFieldAt(
  blocks: ReadonlyArray<Block>,
  blockIndex: number,
  childIndex: number,
  field: Field,
): ReadonlyArray<Block> {
  return blocks.map((block, index) => {
    if (index !== blockIndex || block.kind !== 'paragraph') {
      return block
    }
    const children = block.children.map((child, cIndex) => (cIndex === childIndex ? field : child))
    return { ...block, children }
  })
}
