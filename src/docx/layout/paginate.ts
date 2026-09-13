import type {
  FooterReference,
  HeaderReference,
  HyperlinkChild,
  ParaProps,
  Paragraph,
  ParagraphChild,
  Run,
  RunProps,
  Section,
  SectionProps,
  Table,
  Tab,
} from '../model'

import { breakLines } from './breakLines'
import { layoutTable } from './layoutTable'
import type {
  ColumnBox,
  Page,
  PageLineRef,
  PageTableRef,
  PageTableRowRef,
  PaginatorInput,
} from './pageTypes'
import { PaginationCancelledError } from './pageTypes'
import type { LaidOutTable } from './tableTypes'
import type { EffectiveParaProps, EffectiveRunProps, LineBox, TabStop } from './types'

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
}

type ParagraphUnit = {
  kind: 'paragraph'
  paragraphPath: ReadonlyArray<number>
  lines: ReadonlyArray<LineBox>
  keepWithNext: boolean
  keepLines: boolean
  widowControl: boolean
  pageBreakBefore: boolean
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
}

export async function paginate(input: PaginatorInput): Promise<ReadonlyArray<Page>> {
  const pages: Page[] = []
  const headerFooterLines = input.headerFooterLines ?? new Map<string, ReadonlyArray<LineBox>>()
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

  for (const [sectionIndex, section] of input.document.sections.entries()) {
    const sectionLayout = resolveSectionLayout(section)
    const units = await buildSectionUnits(
      input,
      section,
      sectionIndex,
      sectionLayout.columnWidthPt,
      totalBlocks,
      progressState,
    )

    for (const unit of units) {
      if (unit.kind === 'table' && unit.laidOutTable !== undefined) {
        tableMetaByPath.set(paragraphPathKey(unit.paragraphPath), unit.laidOutTable)
      }
    }

    let pageNumberInSection = 0

    if (sectionIndex > 0) {
      pageNumberInSection = addParityPaddingPages(
        pages,
        sectionLayout,
        sectionIndex,
        pageNumberInSection,
        headerFooterLines,
      )
    }

    let currentPage = createActivePage(
      sectionLayout,
      sectionIndex,
      ++pageNumberInSection,
      pages.length + 1,
      headerFooterLines,
    )

    if (units.length === 0) {
      pages.push(finalizePage(currentPage, pages.length))
      continue
    }

    let pendingGroup: LayoutUnit[] = []

    for (const unit of units) {
      pendingGroup.push(unit)

      if (!unit.keepWithNext) {
        currentPage = flushPendingGroup(pendingGroup, currentPage, pages, () => {
          pageNumberInSection += 1
          return createActivePage(
            sectionLayout,
            sectionIndex,
            pageNumberInSection,
            pages.length + 1,
            headerFooterLines,
          )
        })
        pendingGroup = []
      }
    }

    if (pendingGroup.length > 0) {
      currentPage = flushPendingGroup(pendingGroup, currentPage, pages, () => {
        pageNumberInSection += 1
        return createActivePage(
          sectionLayout,
          sectionIndex,
          pageNumberInSection,
          pages.length + 1,
          headerFooterLines,
        )
      })
    }

    pages.push(finalizePage(currentPage, pages.length))
  }

  return attachPageTables(pages, tableMetaByPath)
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
): number {
  if (sectionLayout.sectionType !== 'evenPage' && sectionLayout.sectionType !== 'oddPage') {
    return pageNumberInSection
  }

  const nextPhysicalPageNumber = pages.length + 1
  const targetParity = sectionLayout.sectionType === 'evenPage' ? 0 : 1

  if (nextPhysicalPageNumber % 2 === targetParity) {
    return pageNumberInSection
  }

  const paddingPage = createActivePage(
    sectionLayout,
    sectionIndex,
    pageNumberInSection + 1,
    nextPhysicalPageNumber,
    headerFooterLines,
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
    let fitCount = countLinesThatFitOnPage(remainingLines, currentPage)

    if (fitCount >= remainingLines.length) {
      placeLineSlice(unit, lineOffset, remainingLines.length, currentPage)
      return currentPage
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
    placeLine(currentPage, unit.paragraphPath, startLineIndex + index, unit.lines[startLineIndex + index])
  }
}

function placeLine(
  currentPage: ActivePage,
  paragraphPath: ReadonlyArray<number>,
  lineIndex: number,
  line: LineBox,
): void {
  let column = currentPage.columns[currentPage.currentColumnIndex]

  while (!lineFitsInColumn(line, column.usedHeightPt, currentPage.contentHeightPt)) {
    if (column.usedHeightPt <= 0 || currentPage.currentColumnIndex >= currentPage.columns.length - 1) {
      break
    }

    currentPage.currentColumnIndex += 1
    column = currentPage.columns[currentPage.currentColumnIndex]
  }

  column.lines.push({
    paragraphPath,
    lineIndex,
    line,
    topPt: currentPage.contentTopPt + column.usedHeightPt,
    leftPt: column.leftPt,
  })
  column.usedHeightPt += line.lineHeight
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
  let columnIndex = currentPage.currentColumnIndex
  let usedHeightPt = currentPage.columns[columnIndex].usedHeightPt
  let count = 0

  for (const line of lines) {
    while (!lineFitsInColumn(line, usedHeightPt, currentPage.contentHeightPt)) {
      if (line.lineHeight > currentPage.contentHeightPt || columnIndex >= currentPage.columns.length - 1) {
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
  }
}

function lineFitsInColumn(line: LineBox, usedHeightPt: number, contentHeightPt: number): boolean {
  return line.lineHeight <= Math.max(contentHeightPt - usedHeightPt, 0)
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
): Promise<ReadonlyArray<LayoutUnit>> {
  const units: LayoutUnit[] = []

  for (const [blockIndex, block] of section.blocks.entries()) {
    if (block.kind === 'paragraph') {
      units.push(
        await buildParagraphUnit(
          input,
          block,
          sectionIndex,
          blockIndex,
          columnWidthPt,
          input.document.defaults?.paragraph,
          input.document.defaults?.run,
        ),
      )
    } else if (block.kind === 'table') {
      units.push(await buildTableUnit(input, block, sectionIndex, blockIndex, columnWidthPt))
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

async function buildParagraphUnit(
  input: PaginatorInput,
  paragraph: Paragraph,
  sectionIndex: number,
  blockIndex: number,
  columnWidthPt: number,
  defaultParaProps: ParaProps | undefined,
  defaultRunProps: RunProps | undefined,
): Promise<ParagraphUnit> {
  void sectionIndex

  const paraProps = mergeParaProps(defaultParaProps, paragraph.props)
  const lines = await breakLines({
    paragraph,
    paraProps,
    runs: collectParagraphRuns(paragraph.children, defaultRunProps),
    availableWidth: columnWidthPt,
    fontResolver: input.fontResolver,
    tabStops: resolveTabStops(paraProps.tabs?.items ?? EMPTY_TABS),
    theme: input.theme,
  })

  return {
    kind: 'paragraph',
    paragraphPath: [blockIndex],
    lines,
    keepWithNext: paraProps.keepNext === true,
    keepLines: paraProps.keepLines === true,
    widowControl: paraProps.widowControl !== false,
    pageBreakBefore: paraProps.pageBreakBefore === true,
  }
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
  defaultRunProps: RunProps | undefined,
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
        runProps: mergeRunProps(defaultRunProps, child.props),
      })
      continue
    }

    if (child.kind === 'hyperlink') {
      runs.push(...collectHyperlinkRuns(child.children, defaultRunProps))
      continue
    }

    if (child.kind === 'ins-revision' || child.kind === 'del-revision') {
      const tag: 'ins' | 'del' = child.kind === 'ins-revision' ? 'ins' : 'del'
      for (const run of child.children) {
        const merged = mergeRunProps(defaultRunProps, run.props)
        runs.push({
          run,
          runProps: { ...merged, _revision: tag },
        })
      }
    }
  }

  return runs
}

function collectHyperlinkRuns(
  children: ReadonlyArray<HyperlinkChild>,
  defaultRunProps: RunProps | undefined,
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
        runProps: mergeRunProps(defaultRunProps, child.props),
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

function mergeParaProps(
  base: ParaProps | undefined,
  override: ParaProps | undefined,
): EffectiveParaProps {
  return {
    ...(base ?? {}),
    ...(override ?? {}),
    ...(base?.spacing || override?.spacing
      ? {
          spacing: {
            ...(base?.spacing ?? {}),
            ...(override?.spacing ?? {}),
          },
        }
      : {}),
    ...(base?.ind || override?.ind
      ? {
          ind: {
            ...(base?.ind ?? {}),
            ...(override?.ind ?? {}),
          },
        }
      : {}),
    ...(base?.tabs || override?.tabs
      ? {
          tabs: {
            ...(base?.tabs ?? {}),
            ...(override?.tabs ?? {}),
            items: override?.tabs?.items ?? base?.tabs?.items ?? EMPTY_TABS,
          },
        }
      : {}),
    ...(base?.numPr || override?.numPr
      ? {
          numPr: {
            ...(base?.numPr ?? {}),
            ...(override?.numPr ?? {}),
          },
        }
      : {}),
    ...(base?.framePr || override?.framePr
      ? {
          framePr: {
            ...(base?.framePr ?? {}),
            ...(override?.framePr ?? {}),
          },
        }
      : {}),
  }
}

function mergeRunProps(base: RunProps | undefined, override: RunProps | undefined): EffectiveRunProps {
  return {
    ...(base ?? {}),
    ...(override ?? {}),
    ...(base?.rFonts || override?.rFonts
      ? {
          rFonts: {
            ...(base?.rFonts ?? {}),
            ...(override?.rFonts ?? {}),
          },
        }
      : {}),
    ...(base?.lang || override?.lang
      ? {
          lang: {
            ...(base?.lang ?? {}),
            ...(override?.lang ?? {}),
          },
        }
      : {}),
  }
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
): ActivePage {
  const headerLines = resolveHeaderFooterLines(
    sectionLayout.headerReferences,
    sectionLayout.titlePage,
    pageNumberInSection,
    physicalPageNumber,
    headerFooterLines,
  )
  const footerLines = resolveHeaderFooterLines(
    sectionLayout.footerReferences,
    sectionLayout.titlePage,
    pageNumberInSection,
    physicalPageNumber,
    headerFooterLines,
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
  }
}

function finalizePage(currentPage: ActivePage, pageIndex: number): Page {
  return {
    sectionIndex: currentPage.sectionIndex,
    pageIndex,
    sizePt: currentPage.sizePt,
    marginsPt: currentPage.marginsPt,
    columns: currentPage.columns.map<ColumnBox>((column) => ({
      widthPt: column.widthPt,
      leftPt: column.leftPt,
      lines: column.lines,
      tables: [],
    })),
    headerLines: currentPage.headerLines,
    footerLines: currentPage.footerLines,
  }
}

function resolveHeaderFooterLines(
  references: ReadonlyArray<HeaderReference | FooterReference>,
  titlePage: boolean,
  pageNumberInSection: number,
  physicalPageNumber: number,
  headerFooterLines: ReadonlyMap<string, ReadonlyArray<LineBox>>,
): ReadonlyArray<LineBox> {
  if (references.length === 0) {
    return EMPTY_LINES
  }

  const preferredReference =
    (titlePage && pageNumberInSection === 1 ? references.find((reference) => reference.type === 'first') : undefined) ??
    (physicalPageNumber % 2 === 0 ? references.find((reference) => reference.type === 'even') : undefined) ??
    references.find((reference) => reference.type === 'default') ??
    references[0]

  return preferredReference ? headerFooterLines.get(preferredReference.id) ?? EMPTY_LINES : EMPTY_LINES
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
