/**
 * Atlas — referenced-font collection (D22 / DXP-18)
 *
 * `DocxViewer.tsx`'s `preloadDocxFonts` eagerly loads all 5 bundled font
 * families (~7MB) on every DOCX open, regardless of which ones the
 * document actually uses — a Times-New-Roman-only document still blocks
 * first paint on loading Calibri/Cambria/Courier New/Arial too.
 *
 * This module provides the pure, testable piece that fix needs: given a
 * parsed `Document`, which Word font names does it actually reference, and
 * which bundled substitute families does that resolve to. The caller
 * (DocxViewer, which owns the actual `FontFace` registration and
 * background-backfill scheduling) can load only `resolveReferencedFontFamilies`'s
 * result up front and load the rest of `FONT_FAMILIES` lazily afterward.
 */

import { FONT_FAMILIES, resolveFontFamily, type FontFamily } from './families'
import type {
  Block,
  Document,
  FontSet,
  HyperlinkChild,
  ParagraphChild,
  RunProps,
  TableChild,
  TableRowChild,
} from '../model'

function addFontSet(names: Set<string>, rFonts: FontSet | undefined): void {
  if (rFonts === undefined) {
    return
  }
  if (rFonts.ascii !== undefined) names.add(rFonts.ascii)
  if (rFonts.hAnsi !== undefined) names.add(rFonts.hAnsi)
  if (rFonts.cs !== undefined) names.add(rFonts.cs)
  if (rFonts.eastAsia !== undefined) names.add(rFonts.eastAsia)
}

function addRunProps(names: Set<string>, props: RunProps | undefined): void {
  if (props !== undefined) {
    addFontSet(names, props.rFonts)
  }
}

function collectFromParagraphChild(names: Set<string>, node: ParagraphChild): void {
  switch (node.kind) {
    case 'run':
      addRunProps(names, node.props)
      return
    case 'hyperlink':
      for (const child of node.children) {
        collectFromHyperlinkChild(names, child)
      }
      return
    case 'ins-revision':
    case 'del-revision':
      for (const child of node.children) {
        collectFromParagraphChild(names, child)
      }
      return
    default:
      return
  }
}

function collectFromHyperlinkChild(names: Set<string>, node: HyperlinkChild): void {
  if (node.kind === 'run') {
    addRunProps(names, node.props)
  }
}

function collectFromBlocks(names: Set<string>, blocks: ReadonlyArray<Block>): void {
  for (const block of blocks) {
    collectFromBlock(names, block)
  }
}

function collectFromBlock(names: Set<string>, block: Block): void {
  if (block.kind === 'paragraph') {
    for (const child of block.children) {
      collectFromParagraphChild(names, child)
    }
    return
  }

  if (block.kind === 'table') {
    for (const row of block.rows) {
      collectFromTableChild(names, row)
    }
  }
}

function collectFromTableChild(names: Set<string>, row: TableChild): void {
  if (row.kind !== 'table-row') {
    return
  }
  for (const cell of row.cells) {
    collectFromTableRowChild(names, cell)
  }
}

function collectFromTableRowChild(names: Set<string>, cell: TableRowChild): void {
  if (cell.kind === 'table-cell') {
    collectFromBlocks(names, cell.blocks)
  }
}

/**
 * Collects every distinct Word font family name referenced anywhere in a
 * parsed `Document` — document defaults, styles' own run formatting, and
 * direct run formatting throughout the body, headers, footers, footnotes,
 * endnotes, and comments.
 *
 * Theme-font placeholders (`w:rFonts`'s `asciiTheme`/`hAnsiTheme`/etc, e.g.
 * "minorHAnsi") are intentionally left unresolved here — resolving those to
 * a concrete family name is `parser/theme.ts`'s `resolveThemeFont`'s job,
 * which needs the loaded theme this module doesn't have. A caller that
 * wants full fidelity should also feed the theme's own resolved
 * major/minor font names in.
 */
export function collectReferencedFontFamilies(document: Document): ReadonlyArray<string> {
  const names = new Set<string>()

  addFontSet(names, document.defaults?.run?.rFonts)

  for (const style of document.styles.values()) {
    addFontSet(names, style.run?.rFonts)
  }

  for (const section of document.sections) {
    collectFromBlocks(names, section.blocks)
  }
  for (const header of document.headers.values()) {
    collectFromBlocks(names, header.blocks)
  }
  for (const footer of document.footers.values()) {
    collectFromBlocks(names, footer.blocks)
  }
  for (const footnote of document.footnotes.values()) {
    collectFromBlocks(names, footnote.blocks)
  }
  for (const endnote of document.endnotes.values()) {
    collectFromBlocks(names, endnote.blocks)
  }
  for (const comment of document.comments.values()) {
    collectFromBlocks(names, comment.body)
  }

  return Array.from(names)
}

/**
 * Resolves `collectReferencedFontFamilies`'s result to the distinct set of
 * bundled `FontFamily` substitutes a document actually needs — the input
 * `loadFontsForFamilies`-style callers should load up front, backfilling
 * the remainder of `FONT_FAMILIES` lazily afterward (DXP-18).
 */
export function resolveReferencedFontFamilies(document: Document): ReadonlyArray<FontFamily> {
  const seen = new Set<FontFamily>()
  const result: FontFamily[] = []

  for (const name of collectReferencedFontFamilies(document)) {
    const resolved = resolveFontFamily(name)
    if (resolved !== null && !seen.has(resolved)) {
      seen.add(resolved)
      result.push(resolved)
    }
  }

  return result
}

/** The bundled families NOT referenced by `document` — safe to backfill lazily. */
export function resolveUnreferencedFontFamilies(document: Document): ReadonlyArray<FontFamily> {
  const referenced = new Set(resolveReferencedFontFamilies(document))
  return FONT_FAMILIES.filter((family) => !referenced.has(family))
}
