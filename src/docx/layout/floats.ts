/**
 * Atlas — DOCX anchored (floating) drawing geometry (Wave 3-D, D4 remainder /
 * DXL-03 / DXP-09).
 *
 * Pure, side-effect-free geometry: given a laid-out `Page` and the
 * `Document` it came from, computes each anchored drawing's absolute
 * page-rectangle (from its `wp:positionH`/`wp:positionV` metadata and the
 * page/margin/column/paragraph frames it can be relative to), plus the
 * horizontal exclusion a `square`/`topAndBottom`-wrapped float imposes on a
 * given line.
 *
 * Deliberately decoupled from `paginate.ts`/`breakLines.ts` (owned by the
 * docx-pagination branch): `computePageFloats` derives everything it needs
 * from a `Page` already produced by `paginate()`, so the renderer
 * (`PageView.tsx`) can call it standalone. See this file's trailing
 * "Integration notes" comment for exactly how a future text-wrap pass in
 * `breakLines.ts` would consume `getLineExclusions`.
 */
import type {
  Document,
  Drawing,
  DrawingHorizontalAlign,
  DrawingPositionH,
  DrawingPositionV,
  DrawingVerticalAlign,
  DrawingWrap,
  HyperlinkChild,
  ParagraphChild,
} from '../model'

import type { ColumnBox, Page, PageLineRef } from './pageTypes'

const EMUS_PER_POINT = 12700

/** A page-absolute rectangle, in the same pt coordinate frame as `Page.sizePt`/`marginsPt`. */
export interface FloatRectPt {
  readonly leftPt: number
  readonly topPt: number
  readonly widthPt: number
  readonly heightPt: number
}

/** One anchored drawing's resolved page placement, ready to render. */
export interface PageFloat {
  readonly drawing: Drawing
  /** The anchoring paragraph's block index within `document.sections[page.sectionIndex].blocks`. */
  readonly blockIndex: number
  readonly rect: FloatRectPt
  readonly behindDoc: boolean
  readonly wrap: DrawingWrap | undefined
}

/** One horizontal band a wrapped float removes from a line's available width. */
export interface LineExclusion {
  readonly leftPt: number
  readonly rightPt: number
}

// ---------------------------------------------------------------------------
// Anchor rectangle geometry
// ---------------------------------------------------------------------------

export interface AnchorRectInput {
  readonly positionH: DrawingPositionH | undefined
  readonly positionV: DrawingPositionV | undefined
  readonly widthPt: number
  readonly heightPt: number
  readonly pageRect: FloatRectPt
  readonly marginRect: FloatRectPt
  readonly columnRect: FloatRectPt
  /**
   * The anchoring paragraph's reference frame for `relativeFrom: 'paragraph'
   * | 'line'` (positionV) — approximated as the paragraph's first laid-out
   * line on this page (Atlas doesn't track a finer per-character anchor
   * point). Also used as the fallback frame for positionH's `'character'`
   * relativeFrom, for the same reason.
   */
  readonly paragraphRect: FloatRectPt
}

/**
 * Computes an anchored drawing's absolute page rectangle from its
 * `wp:positionH`/`wp:positionV` metadata and the candidate reference frames
 * it can be relative to. Pure function — no dependency on `Page`/`Document`
 * shapes, so it's directly unit-testable against hand-built rectangles.
 */
export function computeAnchorRect(input: AnchorRectInput): FloatRectPt {
  const hFrame = resolveHorizontalFrame(input.positionH?.relativeFrom, input)
  const vFrame = resolveVerticalFrame(input.positionV?.relativeFrom, input)
  const leftPt = resolveAxisOffset(
    input.positionH?.align,
    input.positionH?.offsetEmu,
    hFrame.leftPt,
    hFrame.widthPt,
    input.widthPt,
  )
  const topPt = resolveAxisOffset(
    input.positionV?.align,
    input.positionV?.offsetEmu,
    vFrame.topPt,
    vFrame.heightPt,
    input.heightPt,
  )

  return { leftPt, topPt, widthPt: input.widthPt, heightPt: input.heightPt }
}

function resolveHorizontalFrame(
  relativeFrom: DrawingPositionH['relativeFrom'] | undefined,
  input: AnchorRectInput,
): FloatRectPt {
  switch (relativeFrom) {
    case 'page':
      return input.pageRect
    case 'column':
      return input.columnRect
    case 'character':
      // Atlas doesn't track a run's exact horizontal offset within its
      // line at this layer — approximate with the anchoring line's own
      // start, matching the common case of an anchor at a paragraph/run
      // boundary.
      return input.paragraphRect
    case 'margin':
    case 'leftMargin':
    case 'rightMargin':
    case 'insideMargin':
    case 'outsideMargin':
    default:
      // 'insideMargin'/'outsideMargin' (mirrored, odd/even-page margins)
      // and the undefined/malformed fallback all resolve to the plain
      // margin rect — Atlas doesn't model mirrored-margin sections.
      return input.marginRect
  }
}

function resolveVerticalFrame(
  relativeFrom: DrawingPositionV['relativeFrom'] | undefined,
  input: AnchorRectInput,
): FloatRectPt {
  switch (relativeFrom) {
    case 'page':
      return input.pageRect
    case 'paragraph':
    case 'line':
      return input.paragraphRect
    case 'margin':
    case 'topMargin':
    case 'bottomMargin':
    case 'insideMargin':
    case 'outsideMargin':
    default:
      return input.marginRect
  }
}

function resolveAxisOffset(
  align: DrawingHorizontalAlign | DrawingVerticalAlign | undefined,
  offsetEmu: number | undefined,
  frameStartPt: number,
  frameSizePt: number,
  drawingSizePt: number,
): number {
  if (offsetEmu !== undefined) {
    return frameStartPt + offsetEmu / EMUS_PER_POINT
  }

  switch (align) {
    case 'center':
      return frameStartPt + (frameSizePt - drawingSizePt) / 2
    case 'right':
    case 'bottom':
    case 'outside':
      return frameStartPt + (frameSizePt - drawingSizePt)
    case 'left':
    case 'top':
    case 'inside':
    default:
      return frameStartPt
  }
}

// ---------------------------------------------------------------------------
// Line exclusions (square / topAndBottom wrap)
// ---------------------------------------------------------------------------

function emuToPt(emu: number | undefined): number {
  return emu === undefined ? 0 : emu / EMUS_PER_POINT
}

/**
 * A sentinel `LineExclusion` meaning "no valid text position on this line at
 * all" (a `topAndBottom`-wrapped float spanning the line's full height) —
 * any real column width subtracted against `[-Infinity, Infinity]` leaves
 * nothing, forcing the consumer to skip the line/push it below the float,
 * without this module needing to know the column's own width.
 */
const FULL_WIDTH_EXCLUSION: LineExclusion = {
  leftPt: Number.NEGATIVE_INFINITY,
  rightPt: Number.POSITIVE_INFINITY,
}

/**
 * Returns the horizontal band(s) `pageFloats` remove from a line spanning
 * `[lineTopPt, lineTopPt + lineHeightPt)`, for `square`/`tight`/
 * `topAndBottom`-wrapped floats only (`none` imposes no exclusion — text
 * simply overlaps it; `through` is approximated as `square` since Atlas has
 * no picture-contour data to wrap tighter than the bounding box).
 *
 * Integration note (for the docx-pagination branch, which owns
 * `breakLines.ts`): see this file's trailing comment for the couple of
 * lines that would call this from `getLineLimit`/`buildRawLines`.
 */
export function getLineExclusions(
  pageFloats: ReadonlyArray<PageFloat>,
  lineTopPt: number,
  lineHeightPt: number,
): ReadonlyArray<LineExclusion> {
  const lineBottomPt = lineTopPt + lineHeightPt
  const exclusions: LineExclusion[] = []

  for (const pageFloat of pageFloats) {
    const mode = pageFloat.wrap?.mode ?? 'square'
    if (mode === 'none') {
      continue
    }

    const distTPt = emuToPt(pageFloat.wrap?.distTEmu)
    const distBPt = emuToPt(pageFloat.wrap?.distBEmu)
    const floatTopPt = pageFloat.rect.topPt - distTPt
    const floatBottomPt = pageFloat.rect.topPt + pageFloat.rect.heightPt + distBPt

    if (lineBottomPt <= floatTopPt || lineTopPt >= floatBottomPt) {
      continue // no vertical overlap with this line
    }

    if (mode === 'topAndBottom') {
      exclusions.push(FULL_WIDTH_EXCLUSION)
      continue
    }

    // 'square' and the 'tight'/'through' approximation both exclude the
    // float's bounding box (± its distance-from-text margins).
    const distLPt = emuToPt(pageFloat.wrap?.distLEmu)
    const distRPt = emuToPt(pageFloat.wrap?.distREmu)
    const floatLeftPt = pageFloat.rect.leftPt - distLPt
    const floatRightPt = pageFloat.rect.leftPt + pageFloat.rect.widthPt + distRPt
    const side = pageFloat.wrap?.side ?? 'bothSides'

    if (side === 'left') {
      // Text is allowed only to the left of the float: exclude from the
      // float's own left edge out to the column's right edge.
      exclusions.push({ leftPt: floatLeftPt, rightPt: Number.POSITIVE_INFINITY })
    } else if (side === 'right') {
      exclusions.push({ leftPt: Number.NEGATIVE_INFINITY, rightPt: floatRightPt })
    } else {
      // 'bothSides', and 'largest' approximated the same way (Atlas would
      // need the column's width here to pick a side — out of scope for
      // this pure-geometry pass; documented in notesForMerger).
      exclusions.push({ leftPt: floatLeftPt, rightPt: floatRightPt })
    }
  }

  return exclusions
}

// ---------------------------------------------------------------------------
// Per-page float computation
// ---------------------------------------------------------------------------

/**
 * Walks a paragraph's children (including hyperlink/tracked-change
 * wrappers) for every `layout: 'anchor'` `Drawing` it directly contains.
 * Mirrors `PageView.tsx`'s `appendParagraphChildRuns` traversal shape but
 * collects drawings instead of run metadata — kept local rather than
 * shared, since the two walks serve different, unrelated purposes.
 */
function collectAnchorDrawings(children: ReadonlyArray<ParagraphChild>): ReadonlyArray<Drawing> {
  const result: Drawing[] = []

  for (const child of children) {
    if (child.kind === 'run') {
      for (const runChild of child.children) {
        if (runChild.kind === 'drawing' && runChild.layout === 'anchor') {
          result.push(runChild)
        }
      }
      continue
    }

    if (child.kind === 'hyperlink') {
      result.push(...collectAnchorDrawingsFromHyperlink(child.children))
      continue
    }

    if (child.kind === 'ins-revision' || child.kind === 'del-revision') {
      result.push(...collectAnchorDrawings(child.children as ReadonlyArray<ParagraphChild>))
    }
  }

  return result
}

function collectAnchorDrawingsFromHyperlink(
  children: ReadonlyArray<HyperlinkChild>,
): ReadonlyArray<Drawing> {
  const result: Drawing[] = []

  for (const child of children) {
    if (child.kind !== 'run') {
      continue
    }
    for (const runChild of child.children) {
      if (runChild.kind === 'drawing' && runChild.layout === 'anchor') {
        result.push(runChild)
      }
    }
  }

  return result
}

/** Finds the first line of `page`'s lines whose `paragraphPath` is exactly `[blockIndex]`. */
function findParagraphAnchorLine(
  page: Page,
  blockIndex: number,
): { readonly column: ColumnBox; readonly lineRef: PageLineRef } | undefined {
  for (const column of page.columns) {
    for (const lineRef of column.lines) {
      if (lineRef.lineIndex === 0 && pathIsBlock(lineRef.paragraphPath, blockIndex)) {
        return { column, lineRef }
      }
    }
  }

  return undefined
}

function pathIsBlock(path: ReadonlyArray<number>, blockIndex: number): boolean {
  return path.length === 1 && path[0] === blockIndex
}

/**
 * Computes every anchored drawing's page rectangle for one laid-out `Page`.
 *
 * Only top-level body paragraphs are considered — an anchor drawing inside
 * a table cell can't be placed from `Page` alone today, since
 * `PageTableRowRef`'s cell content lines carry no `paragraphPath` back to
 * the source table/cell/paragraph (unlike top-level `PageLineRef`s). That's
 * a real gap (documented for the merger), not a silent one: such a drawing
 * still parses/serializes correctly via `Drawing.positionH`/`V`/`wrap`
 * (DXP-09), it just isn't floated on screen yet.
 *
 * A paragraph whose anchor drawing has no matching line on this page
 * (because the paragraph landed on a different page, or hasn't been placed
 * at all) simply contributes no float — matching Word's own behavior that a
 * floating picture only appears on the page its anchor paragraph is on.
 */
export function computePageFloats(page: Page, document: Document): ReadonlyArray<PageFloat> {
  const section = document.sections[page.sectionIndex]
  if (section === undefined) {
    return []
  }

  const pageRect: FloatRectPt = { leftPt: 0, topPt: 0, widthPt: page.sizePt.width, heightPt: page.sizePt.height }
  const marginRect: FloatRectPt = {
    leftPt: page.marginsPt.left,
    topPt: page.marginsPt.top,
    widthPt: Math.max(0, page.sizePt.width - page.marginsPt.left - page.marginsPt.right),
    heightPt: Math.max(0, page.sizePt.height - page.marginsPt.top - page.marginsPt.bottom),
  }

  const floats: PageFloat[] = []

  section.blocks.forEach((block, blockIndex) => {
    if (block.kind !== 'paragraph') {
      return
    }

    const anchorDrawings = collectAnchorDrawings(block.children)
    if (anchorDrawings.length === 0) {
      return
    }

    const anchorLine = findParagraphAnchorLine(page, blockIndex)
    if (anchorLine === undefined) {
      return // this paragraph didn't land on this page
    }

    const columnRect: FloatRectPt = {
      leftPt: anchorLine.column.leftPt,
      topPt: page.marginsPt.top,
      widthPt: anchorLine.column.widthPt,
      heightPt: Math.max(0, page.sizePt.height - page.marginsPt.top - page.marginsPt.bottom),
    }
    // `PageLineRef.topPt` is relative to the content area (below the
    // header); `PageView.tsx` renders it inside a column `<div>` already
    // offset by `marginsPt.top`, so the true page-absolute top is their
    // sum. `leftPt` is already page-absolute (PageView subtracts
    // `column.leftPt` back out only to make it column-relative for
    // rendering inside that div) — see paginate.ts's `placeLine`.
    const paragraphRect: FloatRectPt = {
      leftPt: anchorLine.lineRef.leftPt,
      topPt: page.marginsPt.top + anchorLine.lineRef.topPt,
      widthPt: anchorLine.lineRef.line.width,
      heightPt: anchorLine.lineRef.line.lineHeight,
    }

    for (const drawing of anchorDrawings) {
      const widthPt = resolveExtentPt(drawing.extent?.cx)
      const heightPt = resolveExtentPt(drawing.extent?.cy)
      const rect = computeAnchorRect({
        positionH: drawing.positionH,
        positionV: drawing.positionV,
        widthPt,
        heightPt,
        pageRect,
        marginRect,
        columnRect,
        paragraphRect,
      })

      floats.push({
        drawing,
        blockIndex,
        rect,
        behindDoc: drawing.behindDoc ?? false,
        wrap: drawing.wrap,
      })
    }
  })

  return floats
}

const DEFAULT_DRAWING_SIZE_PT = 96

function resolveExtentPt(emu: number | undefined): number {
  return typeof emu === 'number' && emu > 0 ? emu / EMUS_PER_POINT : DEFAULT_DRAWING_SIZE_PT
}

// ---------------------------------------------------------------------------
// Integration notes for the docx-pagination branch (paginate.ts/breakLines.ts)
// ---------------------------------------------------------------------------
//
// This module intentionally computes floats from an already-finished `Page`
// rather than threading them through pagination itself, so it can ship
// without touching paginate.ts/breakLines.ts (both hotspots shared by
// several wave3-D tasks). Wiring real square/topAndBottom text wrap-around
// is a small, well-isolated follow-up once this lands:
//
// 1. Pagination would need to compute a page's floats *before* laying out
//    that page's lines (a chicken-and-egg step this module doesn't have to
//    solve, since it runs *after* layout) — the practical approach is a
//    two-pass paginate: lay out each page once to learn where paragraphs
//    land, call `computePageFloats`, then re-run `breakLines` for
//    paragraphs whose lines overlap a float's vertical span with a
//    per-line width reduced by `getLineExclusions`.
// 2. The one call site: in `breakLines.ts`'s `getLineLimit` (or
//    `buildRawLines`'s per-line width computation), subtract each
//    `LineExclusion` that falls within `[column.leftPt, column.leftPt +
//    column.widthPt)` from the available width — e.g. narrow the line's
//    right edge when the exclusion sits against the column's right margin,
//    or split the line into a left/right segment when the exclusion sits in
//    the middle (Atlas's current single-segment-per-line model would need a
//    LineBox extension for that last case — a `leadingOffsetPt` per line is
//    the minimal version, sufficient for the common "image at the column
//    edge" case even before full split-line support exists).
// 3. `PageFloat.behindDoc` and `wrap.mode === 'none'` floats need no
//    exclusion at all (already handled — `getLineExclusions` skips `none`
//    outright); only `square`/`tight`/`through`/`topAndBottom` floats
//    affect line width.
