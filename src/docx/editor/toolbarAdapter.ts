import {
  eighthPoint,
  halfPoint,
  hexColor,
  twip,
  type Border,
  type BorderSet,
  type Document,
  type JustifyContent,
  type TableProps,
} from '../model'

import { toHighlightColor } from './colorMapping'
import { findEnclosingTable, findParagraph } from './commands'
import type { Command, Range } from './commandTypes'
import { pickListNumId } from './insertList'
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

/**
 * DXE-14 — resolves a table-editing toolbar/context-menu command against the
 * cell the cursor is currently in. There's no rectangular multi-cell mouse
 * selection yet (documented remaining scope for DXE-14), so "merge" acts on
 * the current cell and its immediate right-hand neighbor — the single most
 * common real case (merging a header cell with the one beside it) — and
 * "split" reverses a merge on the current cell back into single-column
 * cells. Returns `null` when the cursor isn't inside a table at all, or
 * (merge/split) when the specific operation isn't valid at that cell, so the
 * UI's disabled state (`ToolbarState.insideTable`) and this resolution never
 * disagree about *whether* a table command can run, only about the finer
 * per-command validity `applyCommand` itself already enforces.
 */
function resolveTableCommand(
  kind:
    | 'insert-table-row-above'
    | 'insert-table-row-below'
    | 'insert-table-column-left'
    | 'insert-table-column-right'
    | 'delete-table-row'
    | 'delete-table-column'
    | 'delete-table'
    | 'merge-table-cell-right'
    | 'split-table-cell',
  selection: Range | null,
  document: Document,
): Command | null {
  const paragraphPath = getPrimaryParagraphPath(selection)
  if (paragraphPath === null) {
    return null
  }

  const enclosing = findEnclosingTable(document, paragraphPath)
  if (enclosing === null) {
    return null
  }

  const { tablePath, rowIndex, cellIndex } = enclosing

  switch (kind) {
    case 'insert-table-row-above':
      return { kind: 'insert-table-row', tablePath, at: rowIndex }
    case 'insert-table-row-below':
      return { kind: 'insert-table-row', tablePath, at: rowIndex + 1 }
    case 'insert-table-column-left':
      return { kind: 'insert-table-column', tablePath, at: cellIndex }
    case 'insert-table-column-right':
      return { kind: 'insert-table-column', tablePath, at: cellIndex + 1 }
    case 'delete-table-row':
      return { kind: 'delete-table-row', tablePath, rowIndex }
    case 'delete-table-column':
      return { kind: 'delete-table-column', tablePath, columnIndex: cellIndex }
    case 'delete-table':
      return { kind: 'delete-table', tablePath }
    case 'merge-table-cell-right':
      return { kind: 'merge-table-cells', tablePath, rowIndex, fromCellIndex: cellIndex, toCellIndex: cellIndex + 1 }
    case 'split-table-cell':
      return { kind: 'split-table-cell', tablePath, rowIndex, cellIndex }
  }
}

/**
 * DXE-14 — the two border sets the properties dialog's on/off checkbox
 * toggles between. `'single'`/half-point 4 (2pt in eighth-points) mirrors
 * Word's own default table border when one is turned on from scratch;
 * `'none'` (rather than `'nil'`, which means "unset — inherit") explicitly
 * suppresses every edge, matching what "no borders" means to a user picking
 * it from a dialog.
 */
const TABLE_BORDER_ON: Border = Object.freeze({ style: 'single', size: eighthPoint(4), color: hexColor('#000000') })
const TABLE_BORDER_OFF: Border = Object.freeze({ style: 'none' })

function buildTableBorderSet(bordersOn: boolean): BorderSet {
  const edge = bordersOn ? TABLE_BORDER_ON : TABLE_BORDER_OFF
  return Object.freeze({ top: edge, bottom: edge, left: edge, right: edge, insideH: edge, insideV: edge })
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
        : { kind: 'insert-list', paragraphPaths, numId: pickListNumId(document.numbering, 'bullet'), level: 0 }
    }
    case 'toggle-numbered-list': {
      const paragraphPaths = getParagraphPaths(selection, document)
      return paragraphPaths.length === 0
        ? null
        : { kind: 'insert-list', paragraphPaths, numId: pickListNumId(document.numbering, 'number'), level: 0 }
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
    case 'insert-table-row-above':
    case 'insert-table-row-below':
    case 'insert-table-column-left':
    case 'insert-table-column-right':
    case 'delete-table-row':
    case 'delete-table-column':
    case 'delete-table':
    case 'merge-table-cell-right':
    case 'split-table-cell':
      return resolveTableCommand(toolbarCmd.kind, selection, document)
    case 'set-table-properties': {
      const paragraphPath = getPrimaryParagraphPath(selection)
      const enclosing = paragraphPath === null ? null : findEnclosingTable(document, paragraphPath)
      if (enclosing === null) {
        return null
      }

      const props: TableProps = {
        ...(toolbarCmd.widthTwips !== null ? { tblW: { type: 'dxa' as const, value: twip(toolbarCmd.widthTwips) } } : {}),
        jc: toAlignment(toolbarCmd.alignment),
        tblBorders: buildTableBorderSet(toolbarCmd.bordersOn),
      }

      return { kind: 'apply-table-props', tablePath: enclosing.tablePath, props }
    }
    case 'insert-image':
      return null
    case 'insert-hyperlink':
      return null
    case 'insert-header':
      return null
    case 'insert-footer':
      return null
    case 'insert-page-break': {
      const focus = selection?.focus ?? selection?.anchor
      return focus === undefined
        ? null
        : { kind: 'insert-inline', at: focus, child: { kind: 'break', breakType: 'page' } }
    }
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
    case 'accept-all-changes':
      return { kind: 'accept-all-revisions' }
    case 'reject-all-changes':
      return { kind: 'reject-all-revisions' }
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
