import { halfPoint, hexColor, twip, type Document, type HighlightColor, type JustifyContent } from '../model'

import { findParagraph } from './commands'
import type { Command, Range } from './commandTypes'
import { normalizeRange } from './Selection'
import type { ToolbarCommand } from './toolbar/toolbarTypes'

function getSelectionRange(selection: Range | null): Range | null {
  return selection
}

function getPrimaryParagraphPath(selection: Range | null): ReadonlyArray<number> | null {
  return selection?.focus.paragraphPath ?? selection?.anchor.paragraphPath ?? null
}

function pathsEqual(a: ReadonlyArray<number>, b: ReadonlyArray<number>): boolean {
  return a.length === b.length && a.every((segment, index) => segment === b[index])
}

function collectTopLevelParagraphPaths(document: Document): ReadonlyArray<ReadonlyArray<number>> {
  const paths: Array<ReadonlyArray<number>> = []

  document.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      if (block.kind === 'paragraph') {
        paths.push(Object.freeze([sectionIndex, blockIndex]))
      }
    })
  })

  return paths
}

/**
 * Every paragraph a selection touches, not just its two endpoints — a
 * multi-paragraph bullet/numbered-list toggle or alignment change must apply
 * to every paragraph in between, not skip them. Falls back to the raw
 * anchor/focus paragraph paths (deduplicated) when either endpoint isn't
 * among the document's top-level paragraphs (e.g. one sits inside a table
 * cell), since that's a case this simple contiguous-range enumeration
 * doesn't cover.
 */
function getParagraphPaths(
  selection: Range | null,
  document: Document,
): ReadonlyArray<ReadonlyArray<number>> {
  if (selection === null) {
    return []
  }

  const normalized = normalizeRange(selection)
  const allPaths = collectTopLevelParagraphPaths(document)
  const startIndex = allPaths.findIndex((path) => pathsEqual(path, normalized.start.paragraphPath))
  const endIndex = allPaths.findIndex((path) => pathsEqual(path, normalized.end.paragraphPath))

  if (startIndex === -1 || endIndex === -1) {
    return pathsEqual(normalized.start.paragraphPath, normalized.end.paragraphPath)
      ? [normalized.start.paragraphPath]
      : [normalized.start.paragraphPath, normalized.end.paragraphPath]
  }

  const [lo, hi] = startIndex <= endIndex ? [startIndex, endIndex] : [endIndex, startIndex]
  return allPaths.slice(lo, hi + 1)
}

function toAlignment(align: 'left' | 'center' | 'right' | 'justify'): JustifyContent {
  switch (align) {
    case 'left':
      return 'start'
    case 'center':
      return 'center'
    case 'right':
      return 'end'
    case 'justify':
      return 'both'
  }
}

function toHighlightColor(colorHex: string): HighlightColor | null {
  switch (colorHex.trim().toLowerCase()) {
    case '#000000':
      return 'black'
    case '#0000ff':
      return 'blue'
    case '#00ffff':
      return 'cyan'
    case '#00008b':
      return 'darkBlue'
    case '#008b8b':
      return 'darkCyan'
    case '#a9a9a9':
    case '#666666':
      return 'darkGray'
    case '#006400':
      return 'darkGreen'
    case '#8b008b':
      return 'darkMagenta'
    case '#8b0000':
    case '#980000':
      return 'darkRed'
    case '#b8860b':
    case '#ff9900':
      return 'darkYellow'
    case '#00ff00':
      return 'green'
    case '#d3d3d3':
    case '#cccccc':
    case '#d9d9d9':
    case '#efefef':
      return 'lightGray'
    case '#ff00ff':
      return 'magenta'
    case 'transparent':
    case 'none':
      return 'none'
    case '#ff0000':
      return 'red'
    case '#ffffff':
      return 'white'
    case '#ffff00':
      return 'yellow'
    default:
      return null
  }
}

/**
 * DXE-11 — finds the nearest tracked-change revision to resolve for
 * accept/reject-change: the backend (`accept-revision`/`reject-revision`)
 * already resolves a revision addressed by an exact `{paragraphPath,
 * childIndex}`, but the toolbar only has the current selection. Since a
 * paragraph containing an `ins-revision`/`del-revision` child falls outside
 * the run-index addressing scheme entirely (`getEditableRuns` only flattens
 * plain runs and hyperlinks), the best a selection-driven Accept/Reject
 * button can do is resolve to the first revision in the selection's
 * paragraph — sufficient for the common case of one open revision at a time,
 * with Accept All/Reject All covering documents with several.
 */
function findRevisionAtSelection(
  document: Document,
  selection: Range | null,
): { paragraphPath: ReadonlyArray<number>; childIndex: number } | null {
  const paragraphPath = getPrimaryParagraphPath(selection)
  if (paragraphPath === null) {
    return null
  }

  const paragraph = findParagraph(document, paragraphPath)
  if (paragraph === null) {
    return null
  }

  const childIndex = paragraph.children.findIndex(
    (child) => child.kind === 'ins-revision' || child.kind === 'del-revision',
  )

  return childIndex === -1 ? null : { paragraphPath, childIndex }
}

export function toolbarToCommand(
  toolbarCmd: ToolbarCommand,
  selection: Range | null,
  document: Document,
): Command | null {
  switch (toolbarCmd.kind) {
    case 'set-font-family': {
      const range = getSelectionRange(selection)
      if (range === null) {
        return null
      }

      return {
        kind: 'apply-run-format',
        range,
        format: {
          rFonts: {
            ascii: toolbarCmd.family,
            hAnsi: toolbarCmd.family,
            eastAsia: toolbarCmd.family,
            cs: toolbarCmd.family,
          },
        },
      }
    }
    case 'set-font-size': {
      const range = getSelectionRange(selection)
      if (range === null) {
        return null
      }

      return {
        kind: 'apply-run-format',
        range,
        format: {
          sz: halfPoint(toolbarCmd.sizePt * 2),
          szCs: halfPoint(toolbarCmd.sizePt * 2),
        },
      }
    }
    case 'toggle-bold': {
      const range = getSelectionRange(selection)
      return range === null ? null : { kind: 'apply-run-format', range, format: { bold: true } }
    }
    case 'toggle-italic': {
      const range = getSelectionRange(selection)
      return range === null ? null : { kind: 'apply-run-format', range, format: { italic: true } }
    }
    case 'toggle-underline': {
      const range = getSelectionRange(selection)
      return range === null
        ? null
        : { kind: 'apply-run-format', range, format: { underline: { style: 'single' } } }
    }
    case 'toggle-strike': {
      const range = getSelectionRange(selection)
      return range === null ? null : { kind: 'apply-run-format', range, format: { strike: true } }
    }
    case 'toggle-subscript': {
      const range = getSelectionRange(selection)
      return range === null
        ? null
        : { kind: 'apply-run-format', range, format: { vertAlign: 'subscript' } }
    }
    case 'toggle-superscript': {
      const range = getSelectionRange(selection)
      return range === null
        ? null
        : { kind: 'apply-run-format', range, format: { vertAlign: 'superscript' } }
    }
    case 'set-font-color': {
      const range = getSelectionRange(selection)
      return range === null
        ? null
        : { kind: 'apply-run-format', range, format: { color: hexColor(toolbarCmd.colorHex) } }
    }
    case 'set-highlight-color': {
      const range = getSelectionRange(selection)
      const highlight = toHighlightColor(toolbarCmd.colorHex)

      return range === null || highlight === null
        ? null
        : { kind: 'apply-run-format', range, format: { highlight } }
    }
    case 'set-alignment': {
      const paragraphPaths = getParagraphPaths(selection, document)
      return paragraphPaths.length === 0
        ? null
        : {
            kind: 'apply-para-format',
            paragraphPaths,
            format: { jc: toAlignment(toolbarCmd.align) },
          }
    }
    case 'set-line-spacing': {
      const paragraphPaths = getParagraphPaths(selection, document)
      return paragraphPaths.length === 0
        ? null
        : {
            kind: 'apply-para-format',
            paragraphPaths,
            format: {
              spacing: {
                line: twip(toolbarCmd.spacing * 240),
                lineRule: 'auto',
              },
            },
          }
    }
    case 'toggle-bullet-list': {
      const paragraphPaths = getParagraphPaths(selection, document)
      return paragraphPaths.length === 0
        ? null
        : { kind: 'insert-list', paragraphPaths, numId: 1, level: 0 }
    }
    case 'toggle-numbered-list': {
      const paragraphPaths = getParagraphPaths(selection, document)
      return paragraphPaths.length === 0
        ? null
        : { kind: 'insert-list', paragraphPaths, numId: 2, level: 0 }
    }
    case 'change-indent': {
      const paragraphPath = getPrimaryParagraphPath(selection)
      return paragraphPath === null
        ? null
        : { kind: 'change-list-level', paragraphPath, delta: toolbarCmd.delta }
    }
    case 'apply-style': {
      const paragraphPath = getPrimaryParagraphPath(selection)
      return paragraphPath === null
        ? null
        : { kind: 'apply-style', paragraphPath, styleId: toolbarCmd.styleId }
    }
    case 'insert-table': {
      const focus = selection?.focus ?? selection?.anchor
      return focus === undefined
        ? null
        : { kind: 'insert-table', at: focus, rows: toolbarCmd.rows, cols: toolbarCmd.cols }
    }
    case 'insert-image':
      return null
    case 'insert-hyperlink':
      return null
    case 'insert-header':
      return null
    case 'insert-footer':
      return null
    case 'insert-page-break':
      return null
    case 'insert-comment':
      return null
    case 'set-margins':
      return null
    case 'set-orientation':
      return null
    case 'set-page-size':
      return null
    case 'set-columns':
      return null
    case 'toggle-spell-check':
      return null
    case 'toggle-track-changes':
      return null
    case 'accept-change': {
      const target = findRevisionAtSelection(document, selection)
      return target === null ? null : { kind: 'accept-revision', ...target }
    }
    case 'reject-change': {
      const target = findRevisionAtSelection(document, selection)
      return target === null ? null : { kind: 'reject-revision', ...target }
    }
    case 'open-comments-pane':
      return null
    case 'undo':
      return null
    case 'redo':
      return null
    case 'open-find-replace':
      return null
  }
}
