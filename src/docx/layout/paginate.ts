import type {
  Block,
  Document,
  Footnote,
  FooterReference,
  HeaderReference,
  HyperlinkChild,
  LvlDef,
  ParaProps,
  Paragraph,
  ParagraphChild,
  Run,
  RunProps,
  Section,
  SectionProps,
  SectionVerticalAlign,
  Table,
  Tab,
} from '../model'
import { resolveParaProps, resolveRunProps } from '../parser/cascade'

import { breakLines, resolveLeftIndentPt, resolveLineIndentExtraPt } from './breakLines'
import { itemizeRuns } from './itemize'
import { layoutTable } from './layoutTable'
import {
  createNumberingCounterState,
  MARKER_RUN_INDEX,
  resolveListMarker,
  type NumberingCounterState,
} from './listMarkers'
import {
  assignEndnoteMark,
  assignFootnoteMark,
  collectNoteReferences,
  createNoteNumberingState,
  renumberFootnotesEachPage,
  resetEndnoteCounter,
  resetFootnoteCounter,
  type NoteNumberingState,
} from './noteNumbering'
import type {
  ColumnBox,
  Page,
  PageFootnoteLineRef,
  PageLineRef,
  PageTableRef,
  PageTableRowRef,
  PaginatorInput,
} from './pageTypes'
import { PaginationCancelledError } from './pageTypes'
import type { LaidOutCell, LaidOutRow, LaidOutTable } from './tableTypes'
import type { EffectiveParaProps, EffectiveRunProps, LineBox, LineItem, TabStop } from './types'

/**
 * Yield to the event loop so the renderer can paint the loading indicator
 * and respond to user input between heavy itemize/measure blocks. Uses
 * requestAnimationFrame when available (aligns with browser paint cadence)
 * and falls back to setTimeout(0) in Node tests / non-DOM environments.
 */
async function nextPaint(): Promise<void> {
  await new Promise<void>((resolve) => {
    if (typeof window !== 'undefined' && typeof window.requestAnimationFrame === 'function') {
      window.requestAnimationFrame(() => resolve())
      return
    }
    setTimeout(resolve, 0)
  })
}

/**
 * Yield to the paint loop every YIELD_EVERY blocks while paginating long
 * documents (LaBoetie: 1494 blocks across 1 section). Lower numbers keep
 * the UI snappier; higher numbers reduce total wall time. 24 was chosen
 * empirically: ~50ms of itemize work per batch on typical hardware.
 */
const YIELD_EVERY = 24

async function checkpoint(input: PaginatorInput, tick: number): Promise<void> {
  if (input.shouldCancel !== undefined && input.shouldCancel()) {
    throw new PaginationCancelledError()
  }
  if (tick > 0 && tick % YIELD_EVERY === 0) {
    await nextPaint()
  }
}

function countTotalBlocks(input: PaginatorInput): number {
  let total = 0
  for (const section of input.document.sections) {
    total += section.blocks.length
  }
  return total
}

const DEFAULT_PAGE_WIDTH_PT = 612
const DEFAULT_PAGE_HEIGHT_PT = 792
const DEFAULT_MARGIN_PT = 72
const DEFAULT_HEADER_FOOTER_DISTANCE_PT = 36
const DEFAULT_COLUMN_SPACE_PT = 18
const EMPTY_LINES: ReadonlyArray<LineBox> = []
const EMPTY_TABS: ReadonlyArray<Tab> = []
const EMPTY_HEADER_REFERENCES: ReadonlyArray<HeaderReference> = []
const EMPTY_FOOTER_REFERENCES: ReadonlyArray<FooterReference> = []
const EMPTY_PARAGRAPH_PATH: ReadonlyArray<number> = []
const EMPTY_FOOTNOTE_CONTENT: ReadonlyMap<string, ReadonlyArray<LineBox>> = new Map()
/** Fixed reservation (in points) for the short rule Word draws above a page's footnote area, added once per page the first time a footnote is reserved on it. */
const FOOTNOTE_SEPARATOR_RESERVE_PT = 14
/** Vertical gap between two different footnotes' text within the same page's footnote area. */
const FOOTNOTE_INTER_NOTE_GAP_PT = 4

type ResolvedPageSize = {
  width: number
  height: number
}

type ResolvedPageMargins = {
  top: number
  right: number
  bottom: number
  left: number
  header: number
  footer: number
  gutter: number
}

type ResolvedSectionLayout = {
  sizePt: ResolvedPageSize
  marginsPt: ResolvedPageMargins
  columnWidthPt: number
  columnLeftsPt: ReadonlyArray<number>
  columnCount: number
  sectionType: SectionProps['type']
  titlePage: boolean
  headerReferences: ReadonlyArray<HeaderReference>
  footerReferences: ReadonlyArray<FooterReference>
  vAlign: SectionVerticalAlign | undefined
}

type ParagraphUnit = {
  kind: 'paragraph'
  paragraphPath: ReadonlyArray<number>
  lines: ReadonlyArray<LineBox>
  keepWithNext: boolean
  keepLines: boolean
  widowControl: boolean
  pageBreakBefore: boolean
  /** Resolved alignment/indent/spacing, consumed by `computeLineLeftOffsetPt`. */
  paraProps: EffectiveParaProps
  /** The column width `breakLines` measured this paragraph's lines against. */
  columnWidthPt: number
  /**
   * Blank space reserved above this paragraph's first line: the larger of
   * this paragraph's `spacing.before` and the previous paragraph's
   * `spacing.after` (0 across a table, or when `contextualSpacing` applies
   * to two consecutive paragraphs sharing a style). Applied at placement
   * time, and only when the paragraph doesn't happen to start at the very
   * top of a page/column (see `placeLine`) — matching Word's suppression
   * of spacing at a page/column top.
   */
  leadingGapPt: number
}

type TableUnit = {
  kind: 'table'
  paragraphPath: ReadonlyArray<number>
  lines: ReadonlyArray<LineBox>
  keepWithNext: false
  keepLines: false
  widowControl: false
  pageBreakBefore: false
  /**
   * Present only for the real (non-injected-callback) path. When set, the
   * post-pagination pass converts the synthetic per-row lines placed for
   * this unit into a `PageTableRef` per page/column. Absent for the legacy
   * injected-callback path which keeps the synthetic-line-only behavior
   * for backward compatibility with existing tests.
   */
  laidOutTable?: LaidOutTable
}

type LayoutUnit = ParagraphUnit | TableUnit

type MutableColumn = {
  widthPt: number
  leftPt: number
  lines: PageLineRef[]
  usedHeightPt: number
}

type ActivePage = {
  sectionIndex: number
  sizePt: ResolvedPageSize
  marginsPt: ResolvedPageMargins
  headerLines: ReadonlyArray<LineBox>
  footerLines: ReadonlyArray<LineBox>
  columns: MutableColumn[]
  contentTopPt: number
  contentHeightPt: number
  currentColumnIndex: number
  /** D24/DXL-17 — the owning section's `w:vAlign`, applied as a post-placement offset in `finalizePage`. */
  vAlign: SectionVerticalAlign | undefined
  /**
   * D11 milestone 3 — running total of bottom-of-page space reserved for
   * this page's footnote area (including the separator), grown as a
   * footnote-referencing line is committed via `placeLine`. Body content
   * fit-checks subtract this from `contentHeightPt` (see
   * `effectiveContentHeightPt`) so later lines naturally leave room for it.
   */
  footnoteAreaHeightPt: number
  /** Footnote ids reserved on this page so far, in first-appearance order. */
  footnoteIds: string[]
  /** This page's section's referenced footnotes, pre-laid-out once per section (see `buildFootnoteContentLines`). */
  footnoteContentById: ReadonlyMap<string, ReadonlyArray<LineBox>>
}

export async function paginate(input: PaginatorInput): Promise<ReadonlyArray<Page>> {
  const pages: Page[] = []
  const evenAndOddHeaders = input.evenAndOddHeaders === true
  const headerFooterLines = mergeHeaderFooterLines(
    await buildDefaultHeaderFooterLines(input),
    input.headerFooterLines,
  )
  const totalBlocks = countTotalBlocks(input)
  const progressState = { completedBlocks: 0, tick: 0 }
  /**
   * Maps `paragraphPath` (stringified) of a TableUnit to its `LaidOutTable`.
   * Populated during section unit building (real path only — legacy injected
   * `tableLayout` callbacks do not emit a LaidOutTable and remain
   * synthetic-line-only). Consumed by `attachPageTables` after pagination
   * to convert placed synthetic table lines into `PageTableRef`s.
   */
  const tableMetaByPath = new Map<string, LaidOutTable>()
  const styleCache = createStyleResolutionCache()
  // Numbering counters persist across the whole document (a list can
  // continue across a section break), so this is created once here rather
  // than per-section.
  const listCounterState = createNumberingCounterState()
  // Footnote/endnote numbering (D11 milestones 2/5) — same "persists across
  // the whole document" reasoning as listCounterState; `eachSect` restart
  // resets the footnote counter at the top of each section's iteration
  // below rather than replacing this state object.
  const noteState = createNoteNumberingState()
  const footnoteRestart = input.footnoteNumbering?.restart ?? 'continuous'
  // Endnotes only ever support 'continuous'/'eachSect' in practice (Word's
  // own "Endnote numbering" option omits "restart each page" entirely,
  // unlike its footnote counterpart) — and structurally couldn't mean
  // anything else here regardless: endnotes are placed once, together, at
  // the very end of the whole document (`placeEndnotes`, after this loop),
  // never per-page, so there is no "page" for an `eachPage` reset to key
  // off. An `eachPage` value is therefore treated the same as `continuous`.
  const endnoteRestart = input.endnoteNumbering?.restart === 'eachSect' ? 'eachSect' : 'continuous'

  // Carries across the `for` loop below (rather than being section-local)
  // so a 'continuous'/'nextColumn' section break (D9) can keep flowing
  // into the page the PREVIOUS section left in progress instead of always
  // starting a fresh one. `undefined` only before the first section.
  let currentPage: ActivePage | undefined
  // Captured on every iteration so endnote placement (after the loop) can
  // reuse the LAST section's page geometry — endnotes render at the end of
  // the document using that section's layout, not a layout of their own.
  let lastSectionLayout: ResolvedSectionLayout | undefined

  for (const [sectionIndex, section] of input.document.sections.entries()) {
    const sectionLayout = resolveSectionLayout(section)
    lastSectionLayout = sectionLayout

    if (footnoteRestart === 'eachSect' || footnoteRestart === 'eachPage') {
      // `eachPage` can't be resolved until pages exist (see
      // `noteNumbering.ts`'s module doc) — resetting per-section here still
      // gives it a reasonable placeholder numbering to measure/render
      // against before the post-layout `applyEachPageFootnoteRestart` pass
      // relabels the displayed text per page.
      resetFootnoteCounter(noteState, input.footnoteNumbering?.start ?? 1)
    }

    if (endnoteRestart === 'eachSect') {
      resetEndnoteCounter(noteState, input.endnoteNumbering?.start ?? 1)
    }

    const sectionFootnoteIds: string[] = []
    const units = await buildSectionUnits(
      input,
      section,
      sectionIndex,
      sectionLayout.columnWidthPt,
      totalBlocks,
      progressState,
      styleCache,
      listCounterState,
      noteState,
      sectionFootnoteIds,
    )

    for (const unit of units) {
      if (unit.kind === 'table' && unit.laidOutTable !== undefined) {
        tableMetaByPath.set(paragraphPathKey(unit.paragraphPath), unit.laidOutTable)
      }
    }

    const footnoteContentById = await buildSectionFootnoteContent(
      input,
      sectionFootnoteIds,
      resolveFullContentWidthPt(sectionLayout),
      styleCache,
      noteState,
    )

    let pageNumberInSection = 0
    const openNewPage = (): ActivePage => {
      pageNumberInSection += 1
      return createActivePage(
        sectionLayout,
        sectionIndex,
        pageNumberInSection,
        pages.length + 1,
        headerFooterLines,
        evenAndOddHeaders,
        footnoteContentById,
      )
    }

    // A section break's own type decides whether this section continues
    // flowing into the previous section's last page (D9) or always starts
    // a fresh one — the very first section never has a "previous" page to
    // continue.
    const continuation = sectionIndex === 0 ? 'fresh' : resolveSectionContinuation(section.props.type)

    // `sectionPage` (definitely `ActivePage`, unlike the loop-spanning
    // `currentPage`) carries this section's in-progress page through the
    // rest of this iteration, including the nested unit-placement loop
    // below — reassigning `currentPage` itself across a loop boundary
    // doesn't narrow its `ActivePage | undefined` type back down reliably.
    let sectionPage: ActivePage

    if (continuation === 'continuous' && currentPage !== undefined) {
      sectionPage = applyContinuousSectionGeometry(currentPage, sectionLayout, sectionIndex)
      sectionPage.footnoteContentById = mergeFootnoteContent(sectionPage.footnoteContentById, footnoteContentById)
    } else if (continuation === 'nextColumn' && currentPage !== undefined) {
      sectionPage = forceColumnBreak(currentPage, pages, openNewPage)
      sectionPage.sectionIndex = sectionIndex
      sectionPage.vAlign = sectionLayout.vAlign
      sectionPage.footnoteContentById = mergeFootnoteContent(sectionPage.footnoteContentById, footnoteContentById)
    } else {
      if (currentPage !== undefined) {
        pages.push(finalizePage(currentPage, pages.length))
      }
      if (sectionIndex > 0) {
        pageNumberInSection = addParityPaddingPages(
          pages,
          sectionLayout,
          sectionIndex,
          pageNumberInSection,
          headerFooterLines,
          evenAndOddHeaders,
        )
      }
      sectionPage = openNewPage()
    }

    if (units.length === 0) {
      currentPage = sectionPage
      continue
    }

    let pendingGroup: LayoutUnit[] = []

    for (const unit of units) {
      pendingGroup.push(unit)

      if (!unit.keepWithNext) {
        sectionPage = flushPendingGroup(pendingGroup, sectionPage, pages, openNewPage)
        pendingGroup = []
      }
    }

    if (pendingGroup.length > 0) {
      sectionPage = flushPendingGroup(pendingGroup, sectionPage, pages, openNewPage)
    }

    currentPage = sectionPage
  }

  // D11 milestone 4 — endnotes render once, at the very end of the whole
  // document (not per-page/per-section like footnotes), continuing to flow
  // from wherever the last section's content left off using that section's
  // own page geometry.
  if (noteState.endnoteOrder.length > 0 && lastSectionLayout !== undefined) {
    currentPage = await placeEndnotes(
      input,
      noteState,
      lastSectionLayout,
      currentPage,
      pages,
      headerFooterLines,
      evenAndOddHeaders,
      styleCache,
    )
  }

  if (currentPage !== undefined) {
    pages.push(finalizePage(currentPage, pages.length))
  }

  const finishedPages = attachPageTables(pages, tableMetaByPath)

  return footnoteRestart === 'eachPage'
    ? applyEachPageFootnoteRestart(finishedPages, input.footnoteNumbering?.numFmt, input.footnoteNumbering?.start ?? 1)
    : finishedPages
}

/**
 * D11 milestone 1 — builds header/footer content directly from
 * `document.headers`/`document.footers` (paragraphs only; a table inside a
 * header/footer is not laid out here, an accepted scope limit), itemized/
 * broken the same way body paragraphs are.
 *
 * Built per distinct (header/footer id, section content width) pair rather
 * than per id alone: the common case is every section sharing the same
 * width, so the SAME header/footer id is normally laid out exactly once,
 * but a document mixing portrait and landscape sections (or otherwise
 * varying margins) that reuses the same header/footer id across both would
 * otherwise have that id's content measured/wrapped against only the
 * FIRST section's width and then rendered — visibly mis-wrapped — on
 * every other section using a different width. Only ids actually
 * referenced by some section are built at all, so a document with many
 * unused parts (common after several rounds of Word editing) doesn't pay
 * to lay out all of them.
 */
async function buildDefaultHeaderFooterLines(
  input: PaginatorInput,
): Promise<ReadonlyMap<string, ReadonlyArray<LineBox>>> {
  const referencedIdsByWidth = new Map<number, Set<string>>()
  for (const section of input.document.sections) {
    const references = [
      ...(section.props.headerReference ?? EMPTY_HEADER_REFERENCES),
      ...(section.props.footerReference ?? EMPTY_FOOTER_REFERENCES),
    ]
    if (references.length === 0) {
      continue
    }

    const contentWidthPt = resolveFullContentWidthPt(resolveSectionLayout(section))
    const ids = referencedIdsByWidth.get(contentWidthPt) ?? new Set<string>()
    for (const reference of references) {
      ids.add(reference.id)
    }
    referencedIdsByWidth.set(contentWidthPt, ids)
  }

  if (referencedIdsByWidth.size === 0) {
    return EMPTY_FOOTNOTE_CONTENT
  }

  const styleCache = createStyleResolutionCache()
  const result = new Map<string, ReadonlyArray<LineBox>>()

  for (const [contentWidthPt, ids] of referencedIdsByWidth) {
    for (const id of ids) {
      const part = input.document.headers.get(id) ?? input.document.footers.get(id)
      if (part === undefined) {
        continue
      }
      result.set(
        headerFooterContentKey(id, contentWidthPt),
        await buildBlockGroupLines(input, part.blocks, contentWidthPt, styleCache),
      )
    }
  }

  return result
}

/** Composite key for `buildDefaultHeaderFooterLines`'s per-width cache — see its doc comment. */
function headerFooterContentKey(id: string, contentWidthPt: number): string {
  return `${id}::${contentWidthPt}`
}

function mergeHeaderFooterLines(
  builtIn: ReadonlyMap<string, ReadonlyArray<LineBox>>,
  override: ReadonlyMap<string, ReadonlyArray<LineBox>> | undefined,
): ReadonlyMap<string, ReadonlyArray<LineBox>> {
  if (override === undefined || override.size === 0) {
    return builtIn
  }
  return new Map([...builtIn, ...override])
}

function mergeFootnoteContent(
  existing: ReadonlyMap<string, ReadonlyArray<LineBox>>,
  additional: ReadonlyMap<string, ReadonlyArray<LineBox>>,
): ReadonlyMap<string, ReadonlyArray<LineBox>> {
  if (additional.size === 0) {
    return existing
  }
  return new Map([...existing, ...additional])
}

/** Full page-width content measurement (margins/gutter only — ignores column count), used for headers/footers/footnotes/endnotes, which never participate in multi-column body layout. */
function resolveFullContentWidthPt(sectionLayout: ResolvedSectionLayout): number {
  return Math.max(
    0,
    sectionLayout.sizePt.width - sectionLayout.marginsPt.left - sectionLayout.marginsPt.right - sectionLayout.marginsPt.gutter,
  )
}

/**
 * Lays out a header/footer's paragraph blocks (tables are skipped — see
 * `buildDefaultHeaderFooterLines`'s doc comment) into a flat `LineBox[]`,
 * independent of the page-placement machinery paginate.ts otherwise uses:
 * a header/footer's content isn't subject to page breaking, so it doesn't
 * need `ParagraphUnit`'s keep-together/leading-gap bookkeeping, just the
 * same itemize/measure pipeline body paragraphs go through.
 *
 * Each returned line still carries its own `leftOffsetPt` (indent +
 * center/right alignment — see `computeParagraphLineOffsetPt`): unlike body
 * content, this flat line list never passes through `placeLineSlice`'s
 * per-line placement (the only other place that offset gets computed), so
 * without it every header/footer paragraph would render flush left
 * regardless of its own indent or `w:jc` — silently dropping a `center`- or
 * `right`-aligned header/footer paragraph's alignment (a common real-world
 * case: e.g. a simple centered or right-aligned page-number footer).
 */
async function buildBlockGroupLines(
  input: PaginatorInput,
  blocks: ReadonlyArray<Block>,
  contentWidthPt: number,
  styleCache: StyleResolutionCache,
): Promise<ReadonlyArray<LineBox>> {
  const lines: LineBox[] = []

  for (const block of blocks) {
    if (block.kind !== 'paragraph') {
      continue
    }

    const paraProps = applyNumberingIndentFallback(
      resolveEffectiveParaProps(block.props, input.document, styleCache),
      input.document,
    )
    const tabStops = buildTabStops(paraProps, false)

    const paragraphLines = await breakLines({
      paragraph: block,
      paraProps,
      runs: collectParagraphRuns(block.children, block.props?.pStyle, input.document, styleCache),
      availableWidth: contentWidthPt,
      fontResolver: input.fontResolver,
      tabStops,
      theme: input.theme,
    })

    lines.push(...attachParagraphLineOffsets(paragraphLines, paraProps, contentWidthPt))
  }

  return lines
}

/** Stashes each line's `leftOffsetPt` (see `LineBox`'s own doc comment) for content that renders its lines directly instead of through `placeLineSlice`. */
function attachParagraphLineOffsets(
  lines: ReadonlyArray<LineBox>,
  paraProps: EffectiveParaProps,
  columnWidthPt: number,
): ReadonlyArray<LineBox> {
  return lines.map((line, lineIndex) => ({
    ...line,
    leftOffsetPt: computeParagraphLineOffsetPt(paraProps, columnWidthPt, lineIndex, line),
  }))
}

/**
 * D11 milestone 4 — endnotes render once at the very end of the document
 * (a materially different placement rule from footnotes: not per-page, and
 * genuinely allowed to split across a page break like ordinary body
 * content). Reuses the SAME `flushPendingGroup`/`placeUnit` machinery
 * body content flows through, continuing from `currentPage` when there is
 * one (appending after the last section's content) or opening a fresh page
 * using the last section's own geometry otherwise.
 *
 * Endnote paragraphs are given a synthetic two-element `paragraphPath`
 * (`[-1, unitIndex]`) — real body/table paths are always a single-element
 * `[blockIndex]`, so this can never collide with (and be mistaken for) real
 * body content by `attachPageTables`'s path-matching or the renderer's
 * `data-paragraph-path`. In-place endnote editing is out of scope (mirrors
 * D11's plan note that header/footer in-place editing is a stretch item);
 * this path exists only to give the renderer somewhere to draw the text.
 */
async function placeEndnotes(
  input: PaginatorInput,
  noteState: NoteNumberingState,
  sectionLayout: ResolvedSectionLayout,
  currentPage: ActivePage | undefined,
  pages: Page[],
  headerFooterLines: ReadonlyMap<string, ReadonlyArray<LineBox>>,
  evenAndOddHeaders: boolean,
  styleCache: StyleResolutionCache,
): Promise<ActivePage> {
  const contentWidthPt = resolveFullContentWidthPt(sectionLayout)
  const units = await buildEndnoteUnits(input, noteState, contentWidthPt, styleCache)

  let pageNumberInSection = 0
  const openNewPage = (): ActivePage => {
    pageNumberInSection += 1
    return createActivePage(
      sectionLayout,
      currentPage?.sectionIndex ?? 0,
      pageNumberInSection,
      pages.length + 1,
      headerFooterLines,
      evenAndOddHeaders,
      EMPTY_FOOTNOTE_CONTENT,
    )
  }

  let activePage = currentPage ?? openNewPage()
  let pendingGroup: LayoutUnit[] = []

  for (const unit of units) {
    pendingGroup.push(unit)
    if (!unit.keepWithNext) {
      activePage = flushPendingGroup(pendingGroup, activePage, pages, openNewPage)
      pendingGroup = []
    }
  }

  if (pendingGroup.length > 0) {
    activePage = flushPendingGroup(pendingGroup, activePage, pages, openNewPage)
  }

  return activePage
}

async function buildEndnoteUnits(
  input: PaginatorInput,
  noteState: NoteNumberingState,
  contentWidthPt: number,
  styleCache: StyleResolutionCache,
): Promise<ReadonlyArray<ParagraphUnit>> {
  const units: ParagraphUnit[] = []
  let unitIndex = 0

  for (const id of noteState.endnoteOrder) {
    const endnote = input.document.endnotes.get(id)
    if (endnote === undefined) {
      continue
    }

    const paragraphs = endnote.blocks.filter((block): block is Paragraph => block.kind === 'paragraph')
    const mark = noteState.endnoteMarkById.get(id) ?? ''

    for (const [paragraphIndex, paragraph] of paragraphs.entries()) {
      const paraProps = applyNumberingIndentFallback(
        resolveEffectiveParaProps(paragraph.props, input.document, styleCache),
        input.document,
      )
      const leadingItems =
        paragraphIndex === 0 ? await buildNoteMarkerLeadingItems(input, 'endnote', id, mark) : []
      const tabStops = buildTabStops(paraProps, leadingItems.length > 0)

      const lines = await breakLines({
        paragraph,
        paraProps,
        runs: collectParagraphRuns(paragraph.children, paragraph.props?.pStyle, input.document, styleCache),
        availableWidth: contentWidthPt,
        fontResolver: input.fontResolver,
        tabStops,
        theme: input.theme,
        ...(leadingItems.length > 0 ? { leadingItems } : {}),
      })

      units.push({
        kind: 'paragraph',
        paragraphPath: [-1, unitIndex],
        lines,
        keepWithNext: paraProps.keepNext === true,
        keepLines: paraProps.keepLines === true,
        widowControl: paraProps.widowControl !== false,
        pageBreakBefore: false,
        paraProps,
        columnWidthPt: contentWidthPt,
        leadingGapPt:
          paragraphIndex === 0
            ? Math.max(twipToPt(paraProps.spacing?.before), FOOTNOTE_INTER_NOTE_GAP_PT)
            : twipToPt(paraProps.spacing?.before),
      })
      unitIndex += 1
    }
  }

  return units
}

/**
 * D11 milestone 5 — `eachPage` footnote restart. Layout already ran with
 * placeholder continuous numbers (needed to itemize/measure marker text at
 * all — see `noteNumbering.ts`'s module doc); this pass relabels the
 * DISPLAYED marker text only, restarting at `start` on every page using
 * that page's own footnote order (`Page.footnoteLines` is already in
 * first-reference order, courtesy of `reserveFootnotesForLine`).
 */
function applyEachPageFootnoteRestart(
  pages: ReadonlyArray<Page>,
  numFmt: string | undefined,
  start: number,
): ReadonlyArray<Page> {
  const idsByPage = pages.map(collectPageFootnoteIdOrder)
  const marksById = renumberFootnotesEachPage(idsByPage, numFmt, start)

  return pages.map((page) => ({
    ...page,
    columns: page.columns.map((column) => ({
      ...column,
      lines: column.lines.map((lineRef) => ({ ...lineRef, line: relabelBodyNoteMarks(lineRef.line, marksById) })),
    })),
    footnoteLines: relabelFootnoteAreaMarks(page.footnoteLines, marksById),
  }))
}

function collectPageFootnoteIdOrder(page: Page): ReadonlyArray<string> {
  const order: string[] = []
  const seen = new Set<string>()
  for (const footnoteLine of page.footnoteLines) {
    if (!seen.has(footnoteLine.noteId)) {
      seen.add(footnoteLine.noteId)
      order.push(footnoteLine.noteId)
    }
  }
  return order
}

function relabelBodyNoteMarks(line: LineBox, marksById: ReadonlyMap<string, string>): LineBox {
  let changed = false
  const items = line.items.map((item) => {
    if (item.kind !== 'word' || item.noteRef?.kind !== 'footnote') {
      return item
    }
    const mark = marksById.get(item.noteRef.id)
    if (mark === undefined || mark === item.text) {
      return item
    }
    changed = true
    return { ...item, text: mark }
  })
  return changed ? { ...line, items } : line
}

function relabelFootnoteAreaMarks(
  footnoteLines: ReadonlyArray<PageFootnoteLineRef>,
  marksById: ReadonlyMap<string, string>,
): ReadonlyArray<PageFootnoteLineRef> {
  const relabeledForNote = new Set<string>()

  return footnoteLines.map((footnoteLine) => {
    if (relabeledForNote.has(footnoteLine.noteId)) {
      return footnoteLine
    }
    relabeledForNote.add(footnoteLine.noteId)

    const mark = marksById.get(footnoteLine.noteId)
    if (mark === undefined) {
      return footnoteLine
    }

    const items = footnoteLine.line.items.map((item) =>
      item.kind === 'word' && item.runIndex === MARKER_RUN_INDEX ? { ...item, text: mark } : item,
    )
    return { ...footnoteLine, line: { ...footnoteLine.line, items } }
  })
}

function paragraphPathKey(path: ReadonlyArray<number>): string {
  return path.join('.')
}

/**
 * Post-pagination pass: walks each finalized page/column, finds consecutive
 * runs of `PageLineRef`s whose `paragraphPath` matches a known table, and
 * replaces them with a single `PageTableRef`. The synthetic placeholder
 * lines emitted by `buildTableUnit` are removed; the renderer instead draws
 * a real `<table>` for each `PageTableRef`.
 *
 * Repeated header rows: when a table is split across pages and its source
 * `LaidOutTable.repeatHeaderRowCount > 0`, the first N header rows are
 * prepended to every continuation slice (the renderer treats these as
 * `isRepeatedHeader=true` rows but otherwise renders them identically).
 * The repeated headers do NOT consume additional vertical space in the
 * post-pass: the column's pre-existing line-placement already accounted
 * for each placed synthetic line's height, so adding visually repeated
 * header rows on continuations may exceed the original column slot. This
 * trade-off is acceptable for read-only rendering and will be refined
 * during the in-cell editing pass (3.0.10).
 */
function attachPageTables(
  pages: ReadonlyArray<Page>,
  tableMetaByPath: ReadonlyMap<string, LaidOutTable>,
): ReadonlyArray<Page> {
  if (tableMetaByPath.size === 0) {
    return pages
  }

  return pages.map((page) => ({
    ...page,
    columns: page.columns.map((column) => extractTablesFromColumn(column, tableMetaByPath)),
  }))
}

function extractTablesFromColumn(
  column: ColumnBox,
  tableMetaByPath: ReadonlyMap<string, LaidOutTable>,
): ColumnBox {
  const survivingLines: PageLineRef[] = []
  const tables: PageTableRef[] = [...column.tables]
  let cursor = 0

  while (cursor < column.lines.length) {
    const lineRef = column.lines[cursor]
    const key = paragraphPathKey(lineRef.paragraphPath)
    const laidOutTable = tableMetaByPath.get(key)

    if (laidOutTable === undefined) {
      survivingLines.push(lineRef)
      cursor += 1
      continue
    }

    // Found the start of a table slice in this column. Greedily consume all
    // consecutive lines belonging to the same table (same paragraphPath).
    let end = cursor + 1
    while (
      end < column.lines.length &&
      paragraphPathKey(column.lines[end].paragraphPath) === key
    ) {
      end += 1
    }

    const sliceLineRefs = column.lines.slice(cursor, end)
    const tableRef = buildPageTableRefFromSlice(
      sliceLineRefs,
      laidOutTable,
      column.leftPt,
      column.widthPt,
    )

    if (tableRef !== undefined) {
      tables.push(tableRef)
    }

    cursor = end
  }

  return {
    widthPt: column.widthPt,
    leftPt: column.leftPt,
    lines: survivingLines,
    tables,
  }
}

function buildPageTableRefFromSlice(
  sliceLineRefs: ReadonlyArray<PageLineRef>,
  laidOutTable: LaidOutTable,
  columnLeftPt: number,
  columnWidthPt: number,
): PageTableRef | undefined {
  if (sliceLineRefs.length === 0) {
    return undefined
  }

  // For empty tables `buildTableUnit` emits a single placeholder line with
  // height 0. Skip rendering — nothing meaningful to show.
  if (laidOutTable.rows.length === 0) {
    return undefined
  }

  const firstLine = sliceLineRefs[0]

  // Each placed synthetic line corresponds to exactly one source row, in
  // document order. We use the synthetic-line's index within the unit's
  // `lines` array (which paginate threads via `lineIndex` on PageLineRef)
  // to recover which rows landed in this slice.
  const sourceRowIndices = sliceLineRefs.map((lineRef) => lineRef.lineIndex)
  const firstSourceRowIndex = sourceRowIndices[0]
  const isContinuation = firstSourceRowIndex > 0

  const rows: PageTableRowRef[] = []

  // Repeated header rows on continuation slices: prepend the first
  // `repeatHeaderRowCount` rows from the source table, marked as
  // `isRepeatedHeader=true`. Skip duplicates if those rows are already
  // present in this slice (e.g. very first slice contains them naturally).
  if (isContinuation && laidOutTable.repeatHeaderRowCount > 0) {
    for (let headerIndex = 0; headerIndex < laidOutTable.repeatHeaderRowCount; headerIndex += 1) {
      if (sourceRowIndices.includes(headerIndex)) {
        continue
      }
      const headerRow = laidOutTable.rows[headerIndex]
      if (headerRow === undefined) {
        continue
      }
      rows.push({
        sourceRowIndex: headerIndex,
        isHeaderRow: headerRow.isHeader,
        isRepeatedHeader: true,
        row: headerRow,
      })
    }
  }

  for (const sourceRowIndex of sourceRowIndices) {
    const row = laidOutTable.rows[sourceRowIndex]
    if (row === undefined) {
      continue
    }
    rows.push({
      sourceRowIndex,
      isHeaderRow: row.isHeader,
      isRepeatedHeader: false,
      row,
    })
  }

  if (rows.length === 0) {
    return undefined
  }

  const heightPt = rows.reduce((total, rowRef) => total + rowRef.row.heightPt, 0)

  return {
    blockPath: firstLine.paragraphPath,
    topPt: firstLine.topPt,
    leftPt: columnLeftPt,
    widthPt: Math.min(laidOutTable.widthPt, columnWidthPt),
    heightPt,
    columnWidthsPt: laidOutTable.columnWidthsPt,
    borders: laidOutTable.borders,
    shadingFill: laidOutTable.shadingFill,
    rows,
    isContinuation,
  }
}

function addParityPaddingPages(
  pages: Page[],
  sectionLayout: ResolvedSectionLayout,
  sectionIndex: number,
  pageNumberInSection: number,
  headerFooterLines: ReadonlyMap<string, ReadonlyArray<LineBox>>,
  evenAndOddHeaders: boolean,
): number {
  if (sectionLayout.sectionType !== 'evenPage' && sectionLayout.sectionType !== 'oddPage') {
    return pageNumberInSection
  }

  const nextPhysicalPageNumber = pages.length + 1
  const targetParity = sectionLayout.sectionType === 'evenPage' ? 0 : 1

  if (nextPhysicalPageNumber % 2 === targetParity) {
    return pageNumberInSection
  }

  // A blank filler page carries no body content, so it never references a
  // footnote — an empty map is correct here, not a placeholder.
  const paddingPage = createActivePage(
    sectionLayout,
    sectionIndex,
    pageNumberInSection + 1,
    nextPhysicalPageNumber,
    headerFooterLines,
    evenAndOddHeaders,
    EMPTY_FOOTNOTE_CONTENT,
  )

  pages.push(finalizePage(paddingPage, pages.length))
  return pageNumberInSection + 1
}

function flushPendingGroup(
  group: ReadonlyArray<LayoutUnit>,
  currentPage: ActivePage,
  pages: Page[],
  openNewPage: () => ActivePage,
): ActivePage {
  if (group.length === 0) {
    return currentPage
  }

  if (group.some((unit) => unit.pageBreakBefore) && !isPageEmpty(currentPage)) {
    pages.push(finalizePage(currentPage, pages.length))
    currentPage = openNewPage()
  }

  const requiresGroupedPlacement = group.length > 1 || group.some((unit) => unit.keepWithNext)

  if (requiresGroupedPlacement && !isPageEmpty(currentPage) && !groupFitsOnPage(group, currentPage)) {
    pages.push(finalizePage(currentPage, pages.length))
    currentPage = openNewPage()
  }

  for (const unit of group) {
    currentPage = placeUnit(unit, currentPage, pages, openNewPage)
  }

  return currentPage
}

function placeUnit(
  unit: LayoutUnit,
  currentPage: ActivePage,
  pages: Page[],
  openNewPage: () => ActivePage,
): ActivePage {
  if (unit.lines.length === 0) {
    return currentPage
  }

  if (unit.keepLines && !isPageEmpty(currentPage) && !linesFitOnPage(unit.lines, currentPage)) {
    pages.push(finalizePage(currentPage, pages.length))
    currentPage = openNewPage()
  }

  let lineOffset = 0

  while (lineOffset < unit.lines.length) {
    const remainingLines = unit.lines.slice(lineOffset)
    const forcedBreakOffset = findForcedBreakLineOffset(remainingLines)
    let fitCount = countLinesThatFitOnPage(remainingLines, currentPage)

    if (forcedBreakOffset !== undefined && forcedBreakOffset < fitCount) {
      // D10: a manual page/column break (Ctrl+Enter / Ctrl+Shift+Enter)
      // forces a break immediately after this line, regardless of how much
      // room remains on the current page/column — unlike ordinary
      // overflow-driven splitting just below, it isn't subject to
      // widow-control adjustment; the document explicitly asked for a
      // break exactly here.
      const chunkCount = forcedBreakOffset + 1
      placeLineSlice(unit, lineOffset, chunkCount, currentPage)
      const breakingLine = remainingLines[forcedBreakOffset]
      lineOffset += chunkCount
      currentPage =
        breakingLine.endsWithPageBreak === true
          ? forcePageBreak(currentPage, pages, openNewPage)
          : forceColumnBreak(currentPage, pages, openNewPage)
      continue
    }

    if (fitCount >= remainingLines.length) {
      placeLineSlice(unit, lineOffset, remainingLines.length, currentPage)
      return currentPage
    }

    // D24b/DXL-15 fix — a table row that doesn't fit is handled entirely
    // separately from a paragraph's lines below: place whatever rows DO
    // fit as-is (possibly zero — `fitCount` can be 0 even for the very
    // first candidate row, e.g. right after a preceding paragraph already
    // used up most of the page), then try splitting the very next
    // unplaced row across the page boundary using THIS page's own
    // remaining space (which is the FULL page when `fitCount` is 0)
    // before ever moving a row wholesale to a new page.
    //
    // The original implementation only ever attempted a split when
    // `fitCount` was 0 AND the page was already completely empty — i.e.
    // only for a row too tall to fit on any single page at all. That
    // covers the extreme case, but misses the ordinary, common one this
    // task actually targets: a moderately sized row that simply doesn't
    // fit in whatever space is LEFT on a page that already has other
    // content (or other rows) on it, but that would fit on a fresh page.
    // The old code just moved such a row whole to the next page (matching
    // this page's own `fitCount` bookkeeping, but not what Word itself
    // does — Word splits it there, using the leftover space, rather than
    // leaving that space blank). See `trySplitTableRowAtOffset`'s own doc
    // comment for exactly when it declines (cantSplit, a merged cell, or
    // no cell can fit even one more line) — every decline case here falls
    // through to the pre-existing move-whole-row-to-a-new-page (or, on an
    // already-empty page, force-place/clip) behavior, unchanged.
    if (unit.kind === 'table') {
      if (fitCount > 0) {
        placeLineSlice(unit, lineOffset, fitCount, currentPage)
        lineOffset += fitCount
      }

      if (lineOffset >= unit.lines.length) {
        return currentPage
      }

      if (trySplitTableRowAtOffset(unit, lineOffset, currentPage)) {
        continue
      }

      if (fitCount === 0 && isPageEmpty(currentPage)) {
        placeLineSlice(unit, lineOffset, 1, currentPage)
        lineOffset += 1
      }

      if (lineOffset < unit.lines.length && !isPageEmpty(currentPage)) {
        pages.push(finalizePage(currentPage, pages.length))
        currentPage = openNewPage()
      }

      continue
    }

    if (fitCount === 0) {
      if (isPageEmpty(currentPage)) {
        placeLineSlice(unit, lineOffset, 1, currentPage)
        lineOffset += 1
      } else {
        pages.push(finalizePage(currentPage, pages.length))
        currentPage = openNewPage()
      }

      if (lineOffset < unit.lines.length && !isPageEmpty(currentPage)) {
        pages.push(finalizePage(currentPage, pages.length))
        currentPage = openNewPage()
      }

      continue
    }

    if (unit.widowControl) {
      const adjustedFitCount = adjustSplitCount(remainingLines.length, fitCount)
      if (adjustedFitCount === 0 && !isPageEmpty(currentPage)) {
        pages.push(finalizePage(currentPage, pages.length))
        currentPage = openNewPage()
        continue
      }
      fitCount = adjustedFitCount === 0 ? fitCount : adjustedFitCount
    }

    placeLineSlice(unit, lineOffset, fitCount, currentPage)
    lineOffset += fitCount

    if (lineOffset < unit.lines.length) {
      pages.push(finalizePage(currentPage, pages.length))
      currentPage = openNewPage()
    }
  }

  return currentPage
}

/** The offset (within `lines`) of the first line ending in a manual page or column break, or `undefined` if none. */
function findForcedBreakLineOffset(lines: ReadonlyArray<LineBox>): number | undefined {
  const index = lines.findIndex((line) => line.endsWithPageBreak === true || line.endsWithColumnBreak === true)
  return index === -1 ? undefined : index
}

/**
 * D24b/DXL-15 — attempts to split the table row backing `unit.lines[rowLineIndex]`
 * across the page boundary, using `currentPage`'s CURRENT remaining space
 * (`effectiveContentHeightPt(currentPage) - column.usedHeightPt`) — whatever
 * that happens to be at the call site: the page's full height when nothing
 * has been placed on it yet, or whatever is left after other content
 * (earlier rows of this same table, or a preceding paragraph) already
 * consumed some of it. Called from `placeUnit`'s table-specific branch for
 * every row that doesn't fit as-is, not only when the row is too tall for
 * an entirely empty page. On success, mutates `unit.lines` and
 * `unit.laidOutTable.rows` in place, replacing that one entry with two
 * (first piece, continuation) so every later index in both arrays shifts by
 * one but STAYS 1:1 aligned with each other — the exact invariant
 * `attachPageTables`'s `buildPageTableRefFromSlice` already relies on to
 * map a placed synthetic line back to its row content, so nothing downstream
 * needs to know a split happened at all: the continuation is just another
 * ordinary row to that code.
 *
 * Returns `false` (no mutation) when there's nothing backing this unit to
 * split (`laidOutTable` absent — the legacy injected-`tableLayout` path),
 * the row is one of the table's repeated header rows (splitting would shift
 * every later index and corrupt `repeatHeaderRowCount`'s "first N rows"
 * assumption — a rare enough case to just keep the old force-place
 * behavior for), or `splitTableRowForPage` itself declines (see its own
 * doc comment) — in every decline case the caller falls back to moving the
 * whole row to a new page (or, on an already-empty page, force-placing it).
 */
function trySplitTableRowAtOffset(unit: TableUnit, rowLineIndex: number, currentPage: ActivePage): boolean {
  const laidOutTable = unit.laidOutTable
  if (laidOutTable === undefined || rowLineIndex < laidOutTable.repeatHeaderRowCount) {
    return false
  }

  const row = laidOutTable.rows[rowLineIndex]
  if (row === undefined) {
    return false
  }

  const column = currentPage.columns[currentPage.currentColumnIndex]
  const availableHeightPt = effectiveContentHeightPt(currentPage) - column.usedHeightPt
  const split = splitTableRowForPage(row, availableHeightPt)
  if (split === undefined) {
    return false
  }

  // Mutates the SAME `LaidOutTable` object in place (rather than replacing
  // `unit.laidOutTable` with a new object) because `paginate()`'s main loop
  // already captured a reference to this exact object in `tableMetaByPath`
  // (keyed by paragraphPath, right after `buildSectionUnits` returns —
  // before any placement/splitting happens). `attachPageTables`'s later
  // post-pass reads through that captured reference, so replacing it here
  // would leave that map pointing at a stale, unsplit table.
  const nextRows = laidOutTable.rows.slice()
  nextRows.splice(rowLineIndex, 1, split.firstPiece, split.continuation)
  laidOutTable.rows = nextRows

  const nextLines = unit.lines.slice()
  nextLines.splice(
    rowLineIndex,
    1,
    createSyntheticLineBox(split.firstPiece.heightPt),
    createSyntheticLineBox(split.continuation.heightPt),
  )
  unit.lines = nextLines

  return true
}

/**
 * D24b/DXL-15 — divides `row`'s cell content into a `firstPiece` whose
 * height fits within `availableHeightPt` and a `continuation` carrying
 * every cell's remaining content, cut only at each cell's own line
 * boundaries (never mid-line — see `splitCellLinesForHeight`). A cell
 * shorter than the split point contributes nothing to the continuation
 * (its content already fully appeared in the first piece), matching how a
 * short cell in an unsplit row simply leaves blank space below its text.
 *
 * Returns `undefined` when splitting can't help or doesn't apply:
 * `row.cantSplit` is set; any cell spans multiple source rows (`rowSpan > 1`
 * or a `vMerge` continuation — dividing a vertically-merged cell's content
 * between two physically separate table fragments is out of scope here,
 * an accepted limitation given this task's effort budget); no cell can fit
 * even one line within `availableHeightPt` (splitting would make zero
 * progress, and the caller would loop forever re-attempting it); or every
 * cell's content already fits (the row's nominal height exceeding
 * `availableHeightPt` in that case comes from a fixed `w:trHeight` taller
 * than its actual content, which trimming text can't shrink).
 */
function splitTableRowForPage(
  row: LaidOutRow,
  availableHeightPt: number,
): { firstPiece: LaidOutRow; continuation: LaidOutRow } | undefined {
  if (row.cantSplit || availableHeightPt <= 0) {
    return undefined
  }

  if (row.cells.some((cell) => cell.rowSpan > 1 || cell.vMergeContinue)) {
    return undefined
  }

  const splits = row.cells.map((cell) =>
    splitCellLinesForHeight(
      cell.contentLines,
      Math.max(0, availableHeightPt - cell.paddingPt.top - cell.paddingPt.bottom),
    ),
  )

  if (!splits.some((split) => split.continuationLines.length > 0)) {
    return undefined
  }

  const firstCells = row.cells.map((cell, index) => buildSplitCell(cell, splits[index].firstLines))
  const firstPieceHeightPt = firstCells.reduce((tallest, cell) => Math.max(tallest, cell.heightPt), 0)

  if (firstPieceHeightPt <= 0 || firstPieceHeightPt > availableHeightPt) {
    return undefined
  }

  const continuationCells = row.cells.map((cell, index) => buildSplitCell(cell, splits[index].continuationLines))
  const continuationHeightPt = continuationCells.reduce((tallest, cell) => Math.max(tallest, cell.heightPt), 0)

  return {
    firstPiece: { ...row, heightPt: firstPieceHeightPt, cells: firstCells },
    continuation: { ...row, heightPt: continuationHeightPt, cells: continuationCells },
  }
}

type CellLineSplit = {
  readonly firstLines: ReadonlyArray<LineBox>
  readonly continuationLines: ReadonlyArray<LineBox>
}

/** Where within `lines` to cut so `firstLines`'s total height fits `availableHeightPt` — always between two lines, never through one. */
function splitCellLinesForHeight(lines: ReadonlyArray<LineBox>, availableHeightPt: number): CellLineSplit {
  let usedHeightPt = 0
  let cutIndex = lines.length

  for (const [index, line] of lines.entries()) {
    if (usedHeightPt + line.lineHeight > availableHeightPt) {
      cutIndex = index
      break
    }
    usedHeightPt += line.lineHeight
  }

  return { firstLines: lines.slice(0, cutIndex), continuationLines: lines.slice(cutIndex) }
}

function buildSplitCell(cell: LaidOutCell, contentLines: ReadonlyArray<LineBox>): LaidOutCell {
  return {
    ...cell,
    contentLines,
    heightPt: sumLineHeights(contentLines) + cell.paddingPt.top + cell.paddingPt.bottom,
  }
}

function forcePageBreak(currentPage: ActivePage, pages: Page[], openNewPage: () => ActivePage): ActivePage {
  pages.push(finalizePage(currentPage, pages.length))
  return openNewPage()
}

/** Advances to the next column on the current page, or opens a new page when already on the last column. */
function forceColumnBreak(currentPage: ActivePage, pages: Page[], openNewPage: () => ActivePage): ActivePage {
  if (currentPage.currentColumnIndex < currentPage.columns.length - 1) {
    currentPage.currentColumnIndex += 1
    return currentPage
  }

  return forcePageBreak(currentPage, pages, openNewPage)
}

type SectionContinuation = 'fresh' | 'continuous' | 'nextColumn'

/**
 * D9: a `continuous` or `nextColumn` section break keeps flowing into the
 * page the previous section left in progress instead of always starting a
 * fresh page — every other section type (the vast majority: `nextPage`,
 * `evenPage`, `oddPage`, or no explicit type at all) starts fresh, exactly
 * as before this task.
 */
function resolveSectionContinuation(sectionType: SectionProps['type']): SectionContinuation {
  if (sectionType === 'continuous') {
    return 'continuous'
  }
  if (sectionType === 'nextColumn') {
    return 'nextColumn'
  }
  return 'fresh'
}

/**
 * A `continuous` section break doesn't start a new page, so the page's own
 * geometry (size, margins, header/footer — already resolved and, for
 * header/footer, already laid out — when the page was created) stays
 * exactly as it was; Word itself defers a page-level property change
 * specified on a continuous section until the next real page. Only the
 * COLUMN geometry can meaningfully change content flow without a page
 * break, so this rebuilds `columns` from the new section's `cols`
 * definition and resumes flowing from the first of them.
 *
 * Existing columns whose index still exists in the new layout keep their
 * already-placed lines (the common case: the break doesn't actually change
 * column count); a genuinely new column index starts empty. Both start at
 * the SAME resumed height — the tallest of the old columns' used height —
 * so new content never overlaps whatever was already placed, even though
 * an actual change in column count means the old and new column shapes
 * don't isometrically line up (a documented, deliberate approximation:
 * getting this pixel-perfect would need the page model to represent
 * differently-shaped column regions stacked vertically on one page, which
 * is out of scope for this task).
 */
function applyContinuousSectionGeometry(
  currentPage: ActivePage,
  sectionLayout: ResolvedSectionLayout,
  sectionIndex: number,
): ActivePage {
  const resumeUsedHeightPt = currentPage.columns.reduce(
    (tallest, column) => Math.max(tallest, column.usedHeightPt),
    0,
  )

  currentPage.sectionIndex = sectionIndex
  currentPage.vAlign = sectionLayout.vAlign
  currentPage.columns = sectionLayout.columnLeftsPt.map((leftPt, columnIndex) => ({
    widthPt: sectionLayout.columnWidthPt,
    leftPt,
    lines: currentPage.columns[columnIndex]?.lines ?? [],
    usedHeightPt: resumeUsedHeightPt,
  }))
  currentPage.currentColumnIndex = 0

  return currentPage
}

function adjustSplitCount(totalRemainingLines: number, fitCount: number): number {
  if (fitCount >= totalRemainingLines) {
    return totalRemainingLines
  }

  if (fitCount <= 1) {
    return 0
  }

  if (totalRemainingLines - fitCount === 1) {
    return fitCount - 1 >= 2 ? fitCount - 1 : 0
  }

  return fitCount
}

function placeLineSlice(
  unit: LayoutUnit,
  startLineIndex: number,
  count: number,
  currentPage: ActivePage,
): void {
  for (let index = 0; index < count; index += 1) {
    const lineIndex = startLineIndex + index
    const line = unit.lines[lineIndex]
    const leftOffsetPt = unit.kind === 'paragraph' ? computeLineLeftOffsetPt(unit, lineIndex, line) : 0
    const leadingGapPt =
      unit.kind === 'paragraph' && lineIndex === 0 ? unit.leadingGapPt : 0
    placeLine(currentPage, unit.paragraphPath, lineIndex, line, leftOffsetPt, leadingGapPt)
  }
}

/**
 * Horizontal offset from the column's left edge for one line of a
 * paragraph: base left indent, plus the first-line/hanging adjustment for
 * that specific line (see `resolveLineIndentExtraPt`), plus an alignment
 * offset for center/right-aligned text (0 for left/both/distribute, whose
 * flush-right edge for "both" comes from justification stretch instead —
 * see D5).
 */
function computeLineLeftOffsetPt(unit: ParagraphUnit, lineIndex: number, line: LineBox): number {
  return computeParagraphLineOffsetPt(unit.paraProps, unit.columnWidthPt, lineIndex, line)
}

/**
 * Same computation as `computeLineLeftOffsetPt`, but taking `paraProps`/
 * `columnWidthPt` directly instead of a `ParagraphUnit` — reused by
 * `buildBlockGroupLines` (header/footer content) and
 * `buildFootnoteContentLines` (footnote body content), neither of which
 * builds a `ParagraphUnit` or flows through `placeLineSlice`'s per-line
 * placement, but both of which still need each line's indent/alignment
 * offset (stashed on the `LineBox` itself as `leftOffsetPt` — see its own
 * doc comment) since their lines render directly rather than being placed.
 */
function computeParagraphLineOffsetPt(
  paraProps: EffectiveParaProps,
  columnWidthPt: number,
  lineIndex: number,
  line: LineBox,
): number {
  const leftIndentPt = resolveLeftIndentPt(paraProps.ind)
  const lineIndentExtraPt = resolveLineIndentExtraPt(paraProps.ind, lineIndex)
  const lineLimitPt = Math.max(0, columnWidthPt - leftIndentPt - lineIndentExtraPt)
  const alignmentOffsetPt = resolveAlignmentOffsetPt(paraProps.jc, lineLimitPt, line.width)

  return leftIndentPt + lineIndentExtraPt + alignmentOffsetPt
}

function resolveAlignmentOffsetPt(
  alignment: EffectiveParaProps['jc'],
  lineLimitPt: number,
  lineWidthPt: number,
): number {
  if (alignment === 'end') {
    return Math.max(0, lineLimitPt - lineWidthPt)
  }

  if (alignment === 'center') {
    return Math.max(0, (lineLimitPt - lineWidthPt) / 2)
  }

  return 0
}

function placeLine(
  currentPage: ActivePage,
  paragraphPath: ReadonlyArray<number>,
  lineIndex: number,
  line: LineBox,
  leftOffsetPt: number = 0,
  leadingGapPt: number = 0,
): void {
  // D11 milestone 3 — grow this page's footnote reservation BEFORE deciding
  // where this line fits, so a line that newly introduces a footnote
  // reference already sees the reduced effective content height (see
  // `effectiveContentHeightPt`). Look-ahead fit-checks elsewhere
  // (`countLinesThatFitOnPage`, called from `groupFitsOnPage`/
  // `linesFitOnPage`/`placeUnit`) mirror this same growth via
  // `simulateFootnoteReservationGrowthPt` so a batch of candidate lines is
  // fit-checked against the reservation as it will actually stand once
  // placed, including a footnote first referenced partway through that same
  // batch — not just whatever was already reserved before the fit-check
  // began.
  reserveFootnotesForLine(currentPage, line)

  let column = currentPage.columns[currentPage.currentColumnIndex]

  while (!lineFitsInColumn(line, column.usedHeightPt, effectiveContentHeightPt(currentPage))) {
    if (column.usedHeightPt <= 0 || currentPage.currentColumnIndex >= currentPage.columns.length - 1) {
      break
    }

    currentPage.currentColumnIndex += 1
    column = currentPage.columns[currentPage.currentColumnIndex]
  }

  // Word suppresses a paragraph's leading spacing when it lands at the very
  // top of a page/column (an empty column here) — only apply it when this
  // paragraph is continuing a column that already has content above it.
  const appliedGapPt = column.usedHeightPt > 0 ? leadingGapPt : 0

  column.lines.push({
    paragraphPath,
    lineIndex,
    line,
    topPt: currentPage.contentTopPt + column.usedHeightPt + appliedGapPt,
    leftPt: column.leftPt + leftOffsetPt,
  })
  column.usedHeightPt += appliedGapPt + line.lineHeight
}

function groupFitsOnPage(group: ReadonlyArray<LayoutUnit>, currentPage: ActivePage): boolean {
  let simulatedPage = clonePage(currentPage)

  for (const unit of group) {
    if (!linesFitOnPage(unit.lines, simulatedPage)) {
      return false
    }

    simulatedPage = simulatePlacedLines(unit.lines, simulatedPage)
  }

  return true
}

function linesFitOnPage(lines: ReadonlyArray<LineBox>, currentPage: ActivePage): boolean {
  return countLinesThatFitOnPage(lines, currentPage) === lines.length
}

function countLinesThatFitOnPage(lines: ReadonlyArray<LineBox>, currentPage: ActivePage): number {
  // D11 milestone 3 fix — a candidate line further down `lines` may itself
  // be the FIRST reference to a footnote, growing the page's footnote
  // reservation once it's actually placed (see `placeLine`'s call to
  // `reserveFootnotesForLine`). Simulating that growth here too — instead
  // of fit-checking every candidate line against today's fixed
  // `effectiveContentHeightPt(currentPage)` — keeps this count from
  // reporting more lines "fit" than actually will once placed: `placeLine`
  // itself never truncates on overflow (only advances columns when
  // possible), so an over-count here previously meant real content could
  // render past the page's true remaining space and visually overlap the
  // footnote area it was about to reserve.
  const reservedPtByLine = simulateFootnoteReservationGrowthPt(lines, currentPage)
  let columnIndex = currentPage.currentColumnIndex
  let usedHeightPt = currentPage.columns[columnIndex].usedHeightPt
  let count = 0

  for (const [index, line] of lines.entries()) {
    const contentHeightPt = Math.max(0, currentPage.contentHeightPt - reservedPtByLine[index])

    while (!lineFitsInColumn(line, usedHeightPt, contentHeightPt)) {
      if (line.lineHeight > contentHeightPt || columnIndex >= currentPage.columns.length - 1) {
        return count
      }

      columnIndex += 1
      usedHeightPt = currentPage.columns[columnIndex].usedHeightPt
    }

    usedHeightPt += line.lineHeight
    count += 1
  }

  return count
}

/**
 * Mirrors `reserveFootnotesForLine`'s reservation growth WITHOUT mutating
 * the page, one running total per line in `lines` (parallel array), so
 * `countLinesThatFitOnPage` can fit-check each candidate line against the
 * footnote area as it would ACTUALLY stand once every earlier line in this
 * same candidate batch has been placed — including a footnote first
 * referenced by one of those earlier candidate lines, not just whatever was
 * already reserved before this fit-check began. Must stay in lock-step with
 * `reserveFootnotesForLine`'s own gap/dedup rules (first-on-page separator
 * vs. inter-note gap; a footnote id reserved at most once per page) or the
 * two would predict and then actually reserve different amounts.
 */
function simulateFootnoteReservationGrowthPt(
  lines: ReadonlyArray<LineBox>,
  currentPage: ActivePage,
): ReadonlyArray<number> {
  const seenIds = new Set(currentPage.footnoteIds)
  let reservedPt = currentPage.footnoteAreaHeightPt
  const perLine: number[] = []

  for (const line of lines) {
    for (const item of line.items) {
      if (item.kind !== 'word' || item.noteRef === undefined || item.noteRef.kind !== 'footnote') {
        continue
      }

      const id = item.noteRef.id
      if (seenIds.has(id)) {
        continue
      }

      const content = currentPage.footnoteContentById.get(id)
      if (content === undefined) {
        continue
      }

      const gapPt = seenIds.size === 0 ? FOOTNOTE_SEPARATOR_RESERVE_PT : FOOTNOTE_INTER_NOTE_GAP_PT
      seenIds.add(id)
      reservedPt += gapPt + sumLineHeights(content)
    }

    perLine.push(reservedPt)
  }

  return perLine
}

function simulatePlacedLines(lines: ReadonlyArray<LineBox>, currentPage: ActivePage): ActivePage {
  const nextPage = clonePage(currentPage)

  for (const line of lines) {
    placeLine(nextPage, EMPTY_PARAGRAPH_PATH, 0, line)
  }

  return nextPage
}

function clonePage(currentPage: ActivePage): ActivePage {
  return {
    ...currentPage,
    columns: currentPage.columns.map((column) => ({
      ...column,
      lines: column.lines.slice(),
    })),
    footnoteIds: currentPage.footnoteIds.slice(),
  }
}

function lineFitsInColumn(line: LineBox, usedHeightPt: number, contentHeightPt: number): boolean {
  return line.lineHeight <= Math.max(contentHeightPt - usedHeightPt, 0)
}

/** `contentHeightPt` (header/footer already excluded) minus whatever footnote area this page has reserved so far — see `placeLine`'s `reserveFootnotesForLine` call. */
function effectiveContentHeightPt(currentPage: ActivePage): number {
  return Math.max(0, currentPage.contentHeightPt - currentPage.footnoteAreaHeightPt)
}

/**
 * D11 milestone 3 — grows `currentPage.footnoteAreaHeightPt` the first time
 * a given footnote id is seen on this page (via a placed line's
 * `noteRef`-tagged word item — see `itemize.ts`), reserving that footnote's
 * pre-laid-out content height (`currentPage.footnoteContentById`, built
 * once per section by `buildFootnoteContentLines`) plus either the
 * separator's fixed height (the page's first footnote) or a small
 * inter-note gap (every subsequent one). A footnote id with no precomputed
 * content (shouldn't normally happen — every id `buildSectionUnits`
 * collects gets content built before placement starts) is silently
 * skipped rather than reserving nothing but still rendering it: it simply
 * won't appear in this page's footnote area, which fails safe under an
 * inconsistency rather than corrupting layout.
 */
function reserveFootnotesForLine(currentPage: ActivePage, line: LineBox): void {
  for (const item of line.items) {
    if (item.kind !== 'word' || item.noteRef === undefined || item.noteRef.kind !== 'footnote') {
      continue
    }

    const id = item.noteRef.id
    if (currentPage.footnoteIds.includes(id)) {
      continue
    }

    const content = currentPage.footnoteContentById.get(id)
    if (content === undefined) {
      continue
    }

    const gapPt = currentPage.footnoteIds.length === 0 ? FOOTNOTE_SEPARATOR_RESERVE_PT : FOOTNOTE_INTER_NOTE_GAP_PT
    currentPage.footnoteIds.push(id)
    currentPage.footnoteAreaHeightPt += gapPt + sumLineHeights(content)
  }
}

function isPageEmpty(currentPage: ActivePage): boolean {
  return currentPage.columns.every((column) => column.lines.length === 0)
}

async function buildSectionUnits(
  input: PaginatorInput,
  section: Section,
  sectionIndex: number,
  columnWidthPt: number,
  totalBlocks: number,
  progressState: { completedBlocks: number; tick: number },
  styleCache: StyleResolutionCache,
  listCounterState: NumberingCounterState,
  noteState: NoteNumberingState,
  referencedFootnoteIds: string[],
): Promise<ReadonlyArray<LayoutUnit>> {
  const units: LayoutUnit[] = []
  let previousParagraphSpacing: PreviousParagraphSpacing | undefined

  for (const [blockIndex, block] of section.blocks.entries()) {
    if (block.kind === 'paragraph') {
      const unit = await buildParagraphUnit(
        input,
        block,
        sectionIndex,
        blockIndex,
        columnWidthPt,
        styleCache,
        previousParagraphSpacing,
        listCounterState,
        noteState,
        referencedFootnoteIds,
      )
      units.push(unit)
      previousParagraphSpacing = {
        spacingAfterPt: twipToPt(unit.paraProps.spacing?.after),
        pStyle: unit.paraProps.pStyle,
      }
    } else if (block.kind === 'table') {
      units.push(await buildTableUnit(input, block, sectionIndex, blockIndex, columnWidthPt))
      previousParagraphSpacing = undefined
    }

    progressState.completedBlocks += 1
    progressState.tick += 1
    if (input.onProgress !== undefined) {
      input.onProgress({
        phase: 'layout',
        completedBlocks: progressState.completedBlocks,
        totalBlocks,
      })
    }
    await checkpoint(input, progressState.tick)
  }

  return units
}

type PreviousParagraphSpacing = {
  readonly spacingAfterPt: number
  readonly pStyle: string | undefined
}

async function buildParagraphUnit(
  input: PaginatorInput,
  paragraph: Paragraph,
  sectionIndex: number,
  blockIndex: number,
  columnWidthPt: number,
  styleCache: StyleResolutionCache,
  previousParagraphSpacing: PreviousParagraphSpacing | undefined,
  listCounterState: NumberingCounterState,
  noteState: NoteNumberingState,
  referencedFootnoteIds: string[],
): Promise<ParagraphUnit> {
  void sectionIndex

  const document = input.document
  const resolvedParaProps = resolveEffectiveParaProps(paragraph.props, document, styleCache)
  const paraProps = applyNumberingIndentFallback(resolvedParaProps, document)
  const leadingItems = await buildMarkerLeadingItems(input, paraProps, styleCache, listCounterState)
  const tabStops = buildTabStops(paraProps, leadingItems.length > 0)

  // D11 milestones 2/5 — assign display marks for any footnote/endnote
  // reference in THIS paragraph before itemizing it, in document order, so
  // `noteState`'s maps already contain them by the time `breakLines`/
  // `itemizeRuns` looks them up. Passing the full accumulated maps (rather
  // than a per-paragraph slice) is harmless — a paragraph only ever looks
  // up its own reference ids — and avoids building a new Map per paragraph.
  for (const ref of collectNoteReferences(paragraph.children)) {
    if (ref.kind === 'footnote') {
      assignFootnoteMark(noteState, ref.id, input.footnoteNumbering?.numFmt)
      referencedFootnoteIds.push(ref.id)
    } else {
      assignEndnoteMark(noteState, ref.id, input.endnoteNumbering?.numFmt)
    }
  }

  const lines = await breakLines({
    paragraph,
    paraProps,
    runs: collectParagraphRuns(paragraph.children, paragraph.props?.pStyle, document, styleCache),
    availableWidth: columnWidthPt,
    fontResolver: input.fontResolver,
    tabStops,
    theme: input.theme,
    ...(leadingItems.length > 0 ? { leadingItems } : {}),
    noteMarks: { footnote: noteState.footnoteMarkById, endnote: noteState.endnoteMarkById },
  })

  return {
    kind: 'paragraph',
    paragraphPath: [blockIndex],
    lines,
    keepWithNext: paraProps.keepNext === true,
    keepLines: paraProps.keepLines === true,
    widowControl: paraProps.widowControl !== false,
    pageBreakBefore: paraProps.pageBreakBefore === true,
    paraProps,
    columnWidthPt,
    leadingGapPt: resolveLeadingGapPt(paraProps, previousParagraphSpacing),
  }
}

/**
 * A list paragraph's indent almost always lives on its NUMBERING LEVEL
 * (`w:lvl/w:pPr/w:ind`), not directly on the paragraph itself — `resolveParaProps`
 * has no numbering awareness, so a list paragraph with no direct `w:ind`
 * would otherwise resolve to no indent at all and lose its hanging-indent
 * marker layout entirely. Fall back to the level's indent only when the
 * paragraph doesn't already resolve one of its own (direct formatting, or a
 * paragraph style, always wins).
 */
function applyNumberingIndentFallback(paraProps: EffectiveParaProps, document: Document): EffectiveParaProps {
  if (paraProps.ind !== undefined || paraProps.numPr?.numId === undefined) {
    return paraProps
  }

  const levelDef = resolveNumberingLevelDef(document, paraProps.numPr.numId, paraProps.numPr.ilvl ?? 0)
  const levelIndent = levelDef?.paragraph?.ind
  if (levelIndent === undefined) {
    return paraProps
  }

  return { ...paraProps, ind: levelIndent }
}

function resolveNumberingLevelDef(document: Document, numId: string, ilvl: number): LvlDef | undefined {
  const def = document.numbering.get(numId)
  return def?.levelOverrides?.get(ilvl)?.levelDefinition ?? def?.levels.get(ilvl)
}

/**
 * Builds the already-itemized marker (+ trailing tab/space) for a list
 * paragraph, measured through the same font pipeline as real content so it
 * participates in line-breaking (D3). Returns an empty array for a
 * non-list paragraph, or when the numbering definition/level can't be
 * resolved.
 */
async function buildMarkerLeadingItems(
  input: PaginatorInput,
  paraProps: EffectiveParaProps,
  styleCache: StyleResolutionCache,
  listCounterState: NumberingCounterState,
): Promise<ReadonlyArray<LineItem>> {
  const marker = resolveListMarker(paraProps.numPr, input.document, listCounterState)
  if (marker === undefined) {
    return []
  }

  const markerRunProps = resolveEffectiveRunProps(marker.runProps, paraProps.pStyle, input.document, styleCache)
  const markerRun: Run = {
    kind: 'run',
    props: marker.runProps,
    children:
      marker.suffix === 'tab'
        ? [{ kind: 'text', value: marker.text }, { kind: 'tab' }]
        : marker.suffix === 'space'
          ? [{ kind: 'text', value: `${marker.text} ` }]
          : [{ kind: 'text', value: marker.text }],
  }

  const markerItems = await itemizeRuns(
    [{ run: markerRun, runProps: markerRunProps }],
    input.fontResolver,
    input.theme,
  )

  return markerItems.map((item) => ({ ...item, runIndex: MARKER_RUN_INDEX }))
}

/**
 * D11 milestones 2-4 — the leading marker (mark text + tab, forced
 * superscript) prepended to a footnote/endnote body's FIRST paragraph only,
 * built through the same itemize pipeline as real content so it
 * participates in measurement exactly like a list marker does (see
 * `buildMarkerLeadingItems`, which this mirrors). Tagged with `noteRef` on
 * the mark word itself (not the trailing tab) so `applyEachPageFootnoteRestart`
 * can find and relabel it after layout without re-deriving which line is
 * "the marker line" from scratch.
 */
async function buildNoteMarkerLeadingItems(
  input: PaginatorInput,
  kind: 'footnote' | 'endnote',
  id: string,
  mark: string,
): Promise<ReadonlyArray<LineItem>> {
  if (mark === '') {
    return []
  }

  const markerRun: Run = {
    kind: 'run',
    children: [{ kind: 'text', value: mark }, { kind: 'tab' }],
  }
  const markerRunProps: EffectiveRunProps = { vertAlign: 'superscript' }

  const markerItems = await itemizeRuns([{ run: markerRun, runProps: markerRunProps }], input.fontResolver, input.theme)

  return markerItems.map((item) =>
    item.kind === 'word'
      ? { ...item, runIndex: MARKER_RUN_INDEX, noteRef: { kind, id } }
      : { ...item, runIndex: MARKER_RUN_INDEX },
  )
}

/**
 * D11 milestone 3 — lays out one footnote's body (paragraphs only; a table
 * inside a footnote is skipped, mirroring `buildDefaultHeaderFooterLines`'s
 * header/footer scope limit) into a flat `LineBox[]`, with the resolved
 * mark prepended to the first paragraph. Footnote content is never itself
 * split across a page break in this implementation (see `paginate.ts`'s
 * module-level notes referenced from the plan) — its whole `LineBox[]` is
 * reserved as one unit by `reserveFootnotesForLine`.
 *
 * Note references NESTED inside a footnote's own text (a footnote citing
 * another footnote) are not resolved: `noteMarks` is intentionally omitted
 * from the `breakLines` call, so any such reference itemizes with an empty
 * mark (see `noteNumbering.ts`'s module doc and `itemize.ts`'s fallback).
 */
async function buildFootnoteContentLines(
  input: PaginatorInput,
  footnote: Footnote,
  contentWidthPt: number,
  styleCache: StyleResolutionCache,
  noteState: NoteNumberingState,
): Promise<ReadonlyArray<LineBox>> {
  const paragraphs = footnote.blocks.filter((block): block is Paragraph => block.kind === 'paragraph')
  const mark = noteState.footnoteMarkById.get(footnote.id) ?? ''
  const lines: LineBox[] = []

  for (const [paragraphIndex, paragraph] of paragraphs.entries()) {
    const paraProps = applyNumberingIndentFallback(
      resolveEffectiveParaProps(paragraph.props, input.document, styleCache),
      input.document,
    )
    const leadingItems =
      paragraphIndex === 0 ? await buildNoteMarkerLeadingItems(input, 'footnote', footnote.id, mark) : []
    const tabStops = buildTabStops(paraProps, leadingItems.length > 0)

    const paragraphLines = await breakLines({
      paragraph,
      paraProps,
      runs: collectParagraphRuns(paragraph.children, paragraph.props?.pStyle, input.document, styleCache),
      availableWidth: contentWidthPt,
      fontResolver: input.fontResolver,
      tabStops,
      theme: input.theme,
      ...(leadingItems.length > 0 ? { leadingItems } : {}),
    })

    // See `buildBlockGroupLines`'s doc comment on `leftOffsetPt` — footnote
    // content is rendered directly from this flat line list
    // (`buildPageFootnoteLines`), never through `placeLineSlice`, so
    // without this every footnote paragraph (including a hanging indent
    // that aligns a wrapped line's continuation under the marker's own
    // text, exactly like a list marker's hanging indent) would render
    // flush left instead.
    lines.push(...attachParagraphLineOffsets(paragraphLines, paraProps, contentWidthPt))
  }

  return lines
}

/** Builds (and caches by id) the referenced footnotes' content for one section — see `buildFootnoteContentLines`. */
async function buildSectionFootnoteContent(
  input: PaginatorInput,
  referencedFootnoteIds: ReadonlyArray<string>,
  contentWidthPt: number,
  styleCache: StyleResolutionCache,
  noteState: NoteNumberingState,
): Promise<ReadonlyMap<string, ReadonlyArray<LineBox>>> {
  const result = new Map<string, ReadonlyArray<LineBox>>()

  for (const id of referencedFootnoteIds) {
    if (result.has(id)) {
      continue
    }
    const footnote = input.document.footnotes.get(id)
    if (footnote === undefined) {
      continue
    }
    result.set(id, await buildFootnoteContentLines(input, footnote, contentWidthPt, styleCache, noteState))
  }

  return result
}

/**
 * When a marker's trailing tab needs to land at the paragraph's hanging
 * "text starts here" position, add a synthetic tab stop there — see
 * `resolveLineIndentExtraPt`'s doc comment for the hanging-vs-firstLine
 * geometry this derives from. Line 0's own left offset is
 * `leftIndent + extra` (negative `extra` for hanging); the marker's tab is
 * line-relative, so it needs to travel exactly `-extra` from line 0's own
 * origin to land back at `leftIndent`.
 */
function buildTabStops(paraProps: EffectiveParaProps, hasMarker: boolean): ReadonlyArray<TabStop> {
  const configuredStops = resolveTabStops(paraProps.tabs?.items ?? EMPTY_TABS)
  if (!hasMarker) {
    return configuredStops
  }

  const markerTabStopPt = Math.max(0, -resolveLineIndentExtraPt(paraProps.ind, 0))
  return [{ positionPt: markerTabStopPt, alignment: 'left', leader: 'none' }, ...configuredStops]
}

/**
 * The blank space reserved above a paragraph's first line: the larger of
 * this paragraph's own `spacing.before` and the immediately preceding
 * paragraph's `spacing.after` (Word collapses adjacent spacing rather than
 * summing it), unless `contextualSpacing` is set and both paragraphs share
 * the same style (Word's "don't add space between paragraphs of the same
 * style"). `previous` is `undefined` for a section's first paragraph or
 * right after a table, in which case only this paragraph's own `before`
 * applies. Suppressing the gap entirely at an actual page/column top is
 * handled separately, at placement time, since only the placer knows
 * whether a column is currently empty.
 */
function resolveLeadingGapPt(
  paraProps: EffectiveParaProps,
  previous: PreviousParagraphSpacing | undefined,
): number {
  const beforePt = twipToPt(paraProps.spacing?.before)

  if (previous === undefined) {
    return beforePt
  }

  const suppressedByContext =
    paraProps.contextualSpacing === true &&
    previous.pStyle !== undefined &&
    previous.pStyle === paraProps.pStyle

  if (suppressedByContext) {
    return 0
  }

  return Math.max(beforePt, previous.spacingAfterPt)
}

async function buildTableUnit(
  input: PaginatorInput,
  table: Table,
  sectionIndex: number,
  blockIndex: number,
  columnWidthPt: number,
): Promise<TableUnit> {
  if (input.tableLayout !== undefined) {
    // Legacy injected-callback path: callers (tests, custom layouts) provide
    // per-row heights. Preserve historical contract of ONE synthetic line
    // whose height is the sum of all row heights so the existing
    // paginate.test.ts contract holds.
    const rowHeights = await input.tableLayout({
      table,
      availableWidth: columnWidthPt,
      sectionIndex,
      blockIndex,
    })

    return {
      kind: 'table',
      paragraphPath: [blockIndex],
      lines: [createSyntheticLineBox(sumValues(rowHeights))],
      keepWithNext: false,
      keepLines: false,
      widowControl: false,
      pageBreakBefore: false,
    }
  }

  // Real path: lay out the table fully so the post-pagination pass can emit
  // a `PageTableRef` with borders, shading, and cell content. We synthesize
  // one placeholder LineBox per row; the existing line-placement algorithm
  // then naturally splits rows across columns/pages at row boundaries.
  const laidOutTable = await layoutTable({
    table,
    availableWidthPt: columnWidthPt,
    fontResolver: input.fontResolver,
    theme: input.theme,
    styles: input.document.styles,
  })

  const lines: LineBox[] =
    laidOutTable.rows.length === 0
      ? [createSyntheticLineBox(0)]
      : laidOutTable.rows.map((row) => createSyntheticLineBox(row.heightPt))

  return {
    kind: 'table',
    paragraphPath: [blockIndex],
    lines,
    keepWithNext: false,
    keepLines: false,
    widowControl: false,
    pageBreakBefore: false,
    laidOutTable,
  }
}

function createSyntheticLineBox(lineHeight: number): LineBox {
  return {
    items: [],
    width: 0,
    ascent: 0,
    descent: 0,
    lineHeight,
    isJustified: false,
    justificationStretch: 0,
  }
}

function collectParagraphRuns(
  children: ReadonlyArray<ParagraphChild>,
  paraStyleId: string | undefined,
  document: Document,
  styleCache: StyleResolutionCache,
): ReadonlyArray<{
  run: Run
  runProps: EffectiveRunProps
}> {
  const runs: Array<{
    run: Run
    runProps: EffectiveRunProps
  }> = []

  for (const child of children) {
    if (child.kind === 'run') {
      runs.push({
        run: child,
        runProps: resolveEffectiveRunProps(child.props, paraStyleId, document, styleCache),
      })
      continue
    }

    if (child.kind === 'hyperlink') {
      runs.push(...collectHyperlinkRuns(child.children, paraStyleId, document, styleCache))
      continue
    }

    if (child.kind === 'ins-revision' || child.kind === 'del-revision') {
      const tag: 'ins' | 'del' = child.kind === 'ins-revision' ? 'ins' : 'del'
      for (const run of child.children) {
        const resolved = resolveEffectiveRunProps(run.props, paraStyleId, document, styleCache)
        runs.push({
          run,
          runProps: { ...resolved, _revision: tag },
        })
      }
    }
  }

  return runs
}

function collectHyperlinkRuns(
  children: ReadonlyArray<HyperlinkChild>,
  paraStyleId: string | undefined,
  document: Document,
  styleCache: StyleResolutionCache,
): ReadonlyArray<{
  run: Run
  runProps: EffectiveRunProps
}> {
  const runs: Array<{
    run: Run
    runProps: EffectiveRunProps
  }> = []

  for (const child of children) {
    if (child.kind === 'run') {
      runs.push({
        run: child,
        runProps: resolveEffectiveRunProps(child.props, paraStyleId, document, styleCache),
      })
    }
  }

  return runs
}

function resolveTabStops(tabs: ReadonlyArray<Tab>): ReadonlyArray<TabStop> {
  return tabs.map((tab) => ({
    positionPt: twipToPt(tab.position),
    alignment: normalizeTabAlignment(tab.alignment),
    leader: normalizeTabLeader(tab.leader),
  }))
}

function normalizeTabAlignment(alignment: Tab['alignment']): TabStop['alignment'] {
  if (alignment === 'center') {
    return 'center'
  }

  if (alignment === 'right' || alignment === 'end') {
    return 'right'
  }

  if (alignment === 'decimal' || alignment === 'num') {
    return 'decimal'
  }

  return 'left'
}

function normalizeTabLeader(leader: Tab['leader']): TabStop['leader'] {
  if (leader === 'dot' || leader === 'hyphen' || leader === 'underscore') {
    return leader
  }

  return 'none'
}

/**
 * Resolving a paragraph's or run's effective properties means walking the
 * full `basedOn` style chain (cascade.ts's `resolveParaProps`/
 * `resolveRunProps`) every time it's called. A large document re-lays-out on
 * every keystroke (D23), and most paragraphs/runs in a document share a
 * small handful of distinct styleId + direct-props combinations — so cache
 * the resolved result per `paginate()` call, keyed by styleId plus a stable
 * hash of the direct-formatting object, to avoid re-walking the chain for
 * every occurrence.
 */
type StyleResolutionCache = {
  readonly paragraphs: Map<string, EffectiveParaProps>
  readonly runs: Map<string, EffectiveRunProps>
}

function createStyleResolutionCache(): StyleResolutionCache {
  return { paragraphs: new Map(), runs: new Map() }
}

function buildStyleCacheKey(styleId: string | undefined, direct: unknown): string {
  return `${styleId ?? ''}::${JSON.stringify(direct ?? null)}`
}

function resolveEffectiveParaProps(
  direct: ParaProps | undefined,
  document: Document,
  cache: StyleResolutionCache,
): EffectiveParaProps {
  const styleId = direct?.pStyle
  const key = buildStyleCacheKey(styleId, direct)
  const cached = cache.paragraphs.get(key)
  if (cached !== undefined) {
    return cached
  }

  const resolved = resolveParaProps(direct, styleId, document.styles, {
    pPr: document.defaults?.paragraph,
  })
  cache.paragraphs.set(key, resolved)
  return resolved
}

/**
 * A run's effective properties resolve against the run's own `rStyle`
 * (character style) when it has one; otherwise they fall back to the
 * enclosing paragraph's `pStyle` so plain runs in, e.g., a "Heading 1"
 * paragraph pick up that style's run-level formatting (bold/size/color) —
 * `cascade.ts`'s `resolveRunProps` already follows a paragraph-type style's
 * `linked` character style in that case. This is the DXP-02 fix: previously
 * the (dead) legacy renderer resolved every run against the paragraph's
 * style id even when the run carried its own `rStyle`.
 */
function resolveEffectiveRunProps(
  direct: RunProps | undefined,
  paraStyleId: string | undefined,
  document: Document,
  cache: StyleResolutionCache,
): EffectiveRunProps {
  const styleId = direct?.rStyle ?? paraStyleId
  const key = buildStyleCacheKey(styleId, direct)
  const cached = cache.runs.get(key)
  if (cached !== undefined) {
    return cached
  }

  const resolved = resolveRunProps(direct, styleId, document.styles, {
    rPr: document.defaults?.run,
  })
  cache.runs.set(key, resolved)
  return resolved
}

function resolveSectionLayout(section: Section): ResolvedSectionLayout {
  const sizePt = resolvePageSize(section.props)
  const marginsPt = resolveMargins(section.props)
  const columnCount = resolveColumnCount(section.props)
  const columnSpacePt = typeof section.props.cols?.space === 'number'
    ? twipToPt(section.props.cols.space)
    : DEFAULT_COLUMN_SPACE_PT
  const totalContentWidthPt = Math.max(
    0,
    sizePt.width - marginsPt.left - marginsPt.right - marginsPt.gutter - (columnCount - 1) * columnSpacePt,
  )
  const columnWidthPt = columnCount > 0 ? totalContentWidthPt / columnCount : totalContentWidthPt

  return {
    sizePt,
    marginsPt,
    columnWidthPt,
    columnLeftsPt: Array.from({ length: columnCount }, (_, columnIndex) =>
      marginsPt.left + marginsPt.gutter + columnIndex * (columnWidthPt + columnSpacePt),
    ),
    columnCount,
    sectionType: section.props.type,
    titlePage: section.props.titlePg === true,
    headerReferences: section.props.headerReference ?? EMPTY_HEADER_REFERENCES,
    footerReferences: section.props.footerReference ?? EMPTY_FOOTER_REFERENCES,
    vAlign: section.props.vAlign,
  }
}

function resolvePageSize(sectionProps: SectionProps): ResolvedPageSize {
  const width = resolveTwipWithDefault(sectionProps.pgSz?.w, DEFAULT_PAGE_WIDTH_PT)
  const height = resolveTwipWithDefault(sectionProps.pgSz?.h, DEFAULT_PAGE_HEIGHT_PT)

  if (sectionProps.pgSz?.orient === 'landscape' && height > width) {
    return {
      width: height,
      height: width,
    }
  }

  return { width, height }
}

function resolveMargins(sectionProps: SectionProps): ResolvedPageMargins {
  return {
    top: resolveTwipWithDefault(sectionProps.pgMar?.top, DEFAULT_MARGIN_PT),
    right: resolveTwipWithDefault(sectionProps.pgMar?.right, DEFAULT_MARGIN_PT),
    bottom: resolveTwipWithDefault(sectionProps.pgMar?.bottom, DEFAULT_MARGIN_PT),
    left: resolveTwipWithDefault(sectionProps.pgMar?.left, DEFAULT_MARGIN_PT),
    header: resolveTwipWithDefault(sectionProps.pgMar?.header, DEFAULT_HEADER_FOOTER_DISTANCE_PT),
    footer: resolveTwipWithDefault(sectionProps.pgMar?.footer, DEFAULT_HEADER_FOOTER_DISTANCE_PT),
    gutter: twipToPt(sectionProps.pgMar?.gutter),
  }
}

function resolveColumnCount(sectionProps: SectionProps): number {
  const count = sectionProps.cols?.num ?? sectionProps.cols?.col.length ?? 1
  return Math.max(count, 1)
}

function createActivePage(
  sectionLayout: ResolvedSectionLayout,
  sectionIndex: number,
  pageNumberInSection: number,
  physicalPageNumber: number,
  headerFooterLines: ReadonlyMap<string, ReadonlyArray<LineBox>>,
  evenAndOddHeaders: boolean,
  footnoteContentById: ReadonlyMap<string, ReadonlyArray<LineBox>>,
): ActivePage {
  const contentWidthPt = resolveFullContentWidthPt(sectionLayout)
  const headerLines = resolveHeaderFooterLines(
    sectionLayout.headerReferences,
    sectionLayout.titlePage,
    pageNumberInSection,
    physicalPageNumber,
    headerFooterLines,
    evenAndOddHeaders,
    contentWidthPt,
  )
  const footerLines = resolveHeaderFooterLines(
    sectionLayout.footerReferences,
    sectionLayout.titlePage,
    pageNumberInSection,
    physicalPageNumber,
    headerFooterLines,
    evenAndOddHeaders,
    contentWidthPt,
  )
  const headerReservedPt = sumLineHeights(headerLines)
  const footerReservedPt = sumLineHeights(footerLines)

  return {
    sectionIndex,
    sizePt: sectionLayout.sizePt,
    marginsPt: sectionLayout.marginsPt,
    headerLines,
    footerLines,
    columns: sectionLayout.columnLeftsPt.map((leftPt) => ({
      widthPt: sectionLayout.columnWidthPt,
      leftPt,
      lines: [],
      usedHeightPt: 0,
    })),
    contentTopPt: sectionLayout.marginsPt.top + headerReservedPt,
    contentHeightPt: Math.max(
      0,
      sectionLayout.sizePt.height -
        sectionLayout.marginsPt.top -
        sectionLayout.marginsPt.bottom -
        headerReservedPt -
        footerReservedPt,
    ),
    currentColumnIndex: 0,
    vAlign: sectionLayout.vAlign,
    footnoteAreaHeightPt: 0,
    footnoteIds: [],
    footnoteContentById,
  }
}

// TODO(docx-drawings): this is the natural hook point for floating-image
// placement — once `./floats.ts` exists, a finalized `Page` here is exactly
// where an `applyPageFloats(page, ctx)` pass would slot in (after body/
// footnote content is placed, before the page is pushed). None of this
// wave's own milestones (headers/footers/footnotes/endnotes/vAlign/tab
// stops/table row-splitting) needed to reserve space for a float, so no
// call site was added speculatively — left as a comment per this task's
// own instructions rather than importing a module that doesn't exist yet.
function finalizePage(currentPage: ActivePage, pageIndex: number): Page {
  const verticalOffsetPt = resolveVerticalAlignOffsetPt(currentPage)
  const footnoteLines = buildPageFootnoteLines(currentPage)

  return {
    sectionIndex: currentPage.sectionIndex,
    pageIndex,
    sizePt: currentPage.sizePt,
    marginsPt: currentPage.marginsPt,
    columns: currentPage.columns.map<ColumnBox>((column) => ({
      widthPt: column.widthPt,
      leftPt: column.leftPt,
      lines:
        verticalOffsetPt > 0
          ? column.lines.map((lineRef) => ({ ...lineRef, topPt: lineRef.topPt + verticalOffsetPt }))
          : column.lines,
      tables: [],
    })),
    headerLines: currentPage.headerLines,
    footerLines: currentPage.footerLines,
    footnoteLines,
    hasFootnoteSeparator: footnoteLines.length > 0,
  }
}

/**
 * D11 milestone 3 — flattens this page's reserved footnotes (in the order
 * `reserveFootnotesForLine` recorded them) into one list of lines, each
 * positioned relative to the footnote area's own top edge. A gap
 * (`FOOTNOTE_SEPARATOR_RESERVE_PT` before the first note, matching the
 * separator `reserveFootnotesForLine` already reserved space for;
 * `FOOTNOTE_INTER_NOTE_GAP_PT` between subsequent ones) precedes every
 * note's own lines, mirroring exactly what was reserved so the rendered
 * content never exceeds the space `placeLine` accounted for.
 */
function buildPageFootnoteLines(currentPage: ActivePage): ReadonlyArray<PageFootnoteLineRef> {
  const result: PageFootnoteLineRef[] = []
  let topPt = 0

  currentPage.footnoteIds.forEach((noteId, noteIndex) => {
    const content = currentPage.footnoteContentById.get(noteId)
    if (content === undefined) {
      return
    }

    topPt += noteIndex === 0 ? FOOTNOTE_SEPARATOR_RESERVE_PT : FOOTNOTE_INTER_NOTE_GAP_PT

    for (const line of content) {
      result.push({ noteId, line, topPt, leftPt: line.leftOffsetPt ?? 0 })
      topPt += line.lineHeight
    }
  })

  return result
}

/**
 * D24/DXL-17 — `w:vAlign` offsets the page's body content by the leftover
 * ("slack") space between what was actually placed and the page's full
 * content height. Naturally a no-op on any page that's already full (slack
 * <= 0), which is what makes this correct for a multi-page section too:
 * only the section's last (partially filled) page ever has slack to
 * distribute; every earlier page is full and gets offset 0.
 *
 * `both` (OOXML's vertical "justify", which is meant to stretch inter-
 * paragraph spacing so content spans the full height) is approximated as
 * `center` — implementing genuine space redistribution would mean re-
 * deriving every placed line's `leadingGapPt` after the fact, which this
 * single top-down placement pass doesn't support. Documented limitation.
 */
function resolveVerticalAlignOffsetPt(currentPage: ActivePage): number {
  if (currentPage.vAlign === undefined || currentPage.vAlign === 'top') {
    return 0
  }

  const maxUsedHeightPt = currentPage.columns.reduce(
    (tallest, column) => Math.max(tallest, column.usedHeightPt),
    0,
  )
  const slackPt = currentPage.contentHeightPt - currentPage.footnoteAreaHeightPt - maxUsedHeightPt

  if (slackPt <= 0) {
    return 0
  }

  return currentPage.vAlign === 'bottom' ? slackPt : slackPt / 2
}

function resolveHeaderFooterLines(
  references: ReadonlyArray<HeaderReference | FooterReference>,
  titlePage: boolean,
  pageNumberInSection: number,
  physicalPageNumber: number,
  headerFooterLines: ReadonlyMap<string, ReadonlyArray<LineBox>>,
  evenAndOddHeaders: boolean,
  contentWidthPt: number,
): ReadonlyArray<LineBox> {
  if (references.length === 0) {
    return EMPTY_LINES
  }

  const preferredReference =
    (titlePage && pageNumberInSection === 1 ? references.find((reference) => reference.type === 'first') : undefined) ??
    // D11 milestone 1/DXL-09: an `even`-typed reference is only honored when
    // the document actually turned on `w:evenAndOddHeaders` — otherwise
    // Word ignores it and every page (odd or even) uses `default`, even if
    // the source document happens to still define an `even` part (common
    // after converting a doc that once had the setting on).
    (evenAndOddHeaders && physicalPageNumber % 2 === 0
      ? references.find((reference) => reference.type === 'even')
      : undefined) ??
    references.find((reference) => reference.type === 'default') ??
    references[0]

  if (preferredReference === undefined) {
    return EMPTY_LINES
  }

  // Prefer the width-specific entry `buildDefaultHeaderFooterLines` builds
  // (see its doc comment on why the SAME id can have different content per
  // section width); fall back to a plain-id entry for `PaginatorInput.
  // headerFooterLines`'s test-seam overrides, which aren't width-aware.
  return (
    headerFooterLines.get(headerFooterContentKey(preferredReference.id, contentWidthPt)) ??
    headerFooterLines.get(preferredReference.id) ??
    EMPTY_LINES
  )
}

function sumLineHeights(lines: ReadonlyArray<LineBox>): number {
  return lines.reduce((total, line) => total + line.lineHeight, 0)
}

function sumValues(values: ReadonlyArray<number>): number {
  return values.reduce((total, value) => total + value, 0)
}

function twipToPt(value: number | undefined): number {
  return typeof value === 'number' ? value / 20 : 0
}

function resolveTwipWithDefault(value: number | undefined, fallback: number): number {
  return typeof value === 'number' ? value / 20 : fallback
}
