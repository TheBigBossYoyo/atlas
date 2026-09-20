import type { Position } from './Selection'
import type { Document as DocxDocument, Table, TableCell, TableRow } from '../model'

type NodePoint = Readonly<{
  node: Node
  offset: number
}>

type MappedTarget = Readonly<{
  point: NodePoint
  runElement: HTMLElement
  paragraphPath: ReadonlyArray<number>
  runIndex: number
}>

type CharMetrics = Readonly<{
  start: number | null
  length: number
}>

type CaretPositionLike = Readonly<{
  offsetNode: Node
  offset: number
}>

type CaretDocument = Document & {
  caretRangeFromPoint?: (x: number, y: number) => globalThis.Range | null
  caretPositionFromPoint?: (x: number, y: number) => CaretPositionLike | null
}

function parseInteger(value: string | undefined): number | null {
  if (value === undefined) {
    return null
  }

  const parsed = Number.parseInt(value, 10)
  return Number.isNaN(parsed) ? null : parsed
}

function parseParagraphPath(value: string | undefined): ReadonlyArray<number> | null {
  if (value === undefined) {
    return null
  }

  const matches = value.match(/-?\d+/g)
  if (!matches || matches.length === 0) {
    return null
  }

  const path = matches.map((match) => Number.parseInt(match, 10))
  if (path.some((segment) => Number.isNaN(segment))) {
    return null
  }

  return path
}

function paragraphPathsEqual(
  a: ReadonlyArray<number>,
  b: ReadonlyArray<number>,
): boolean {
  if (a.length !== b.length) {
    return false
  }

  for (let index = 0; index < a.length; index += 1) {
    if (a[index] !== b[index]) {
      return false
    }
  }

  return true
}

function maxDomOffset(node: Node): number {
  if (node.nodeType === Node.TEXT_NODE) {
    return node.textContent?.length ?? 0
  }

  return node.childNodes.length
}

function clampDomOffset(node: Node, offset: number): number {
  return Math.max(0, Math.min(offset, maxDomOffset(node)))
}

function pushPoint(points: Array<NodePoint>, point: NodePoint): void {
  const exists = points.some((candidate) => candidate.node === point.node && candidate.offset === point.offset)
  if (!exists) {
    points.push(point)
  }
}

function getCandidatePoints(node: Node, offset: number): ReadonlyArray<NodePoint> {
  const points: Array<NodePoint> = []
  pushPoint(points, {
    node,
    offset: clampDomOffset(node, offset),
  })

  if (node.nodeType !== Node.ELEMENT_NODE) {
    return points
  }

  const element = node as Element
  const childCount = element.childNodes.length
  if (childCount === 0) {
    return points
  }

  const safeOffset = Math.max(0, Math.min(offset, childCount))
  const nextChild = safeOffset < childCount ? element.childNodes[safeOffset] : null
  const previousChild = safeOffset > 0 ? element.childNodes[safeOffset - 1] : null

  if (nextChild) {
    pushPoint(points, { node: nextChild, offset: 0 })
  }

  if (previousChild) {
    pushPoint(points, {
      node: previousChild,
      offset: maxDomOffset(previousChild),
    })
  }

  return points
}

function getElementFromNode(node: Node): HTMLElement | null {
  if (node.nodeType === Node.ELEMENT_NODE) {
    return node as HTMLElement
  }

  return node.parentElement
}

function findClosestRunElement(node: Node): HTMLElement | null {
  let current: HTMLElement | null = getElementFromNode(node)

  while (current) {
    if (current.dataset.runIndex !== undefined) {
      return current
    }

    current = current.parentElement
  }

  return null
}

function findClosestParagraphPath(node: Node): ReadonlyArray<number> | null {
  let current: HTMLElement | null = getElementFromNode(node)

  while (current) {
    const paragraphPath = parseParagraphPath(current.dataset.paragraphPath)
    if (paragraphPath) {
      return paragraphPath
    }

    current = current.parentElement
  }

  return null
}

function getCharMetrics(element: HTMLElement): CharMetrics {
  const start = parseInteger(element.dataset.charStart)
  const end = parseInteger(element.dataset.charEnd)
  const textLength = element.textContent?.length ?? 0

  if (start === null) {
    return {
      start: null,
      length: textLength,
    }
  }

  if (end !== null && end >= start) {
    return {
      start,
      length: end - start,
    }
  }

  return {
    start,
    length: textLength,
  }
}

function getTextOffsetWithin(
  runElement: HTMLElement,
  point: NodePoint,
  document: Document,
): number | null {
  if (runElement.ownerDocument !== document || point.node.ownerDocument !== document) {
    return null
  }

  if (runElement !== point.node && !runElement.contains(point.node)) {
    return null
  }

  try {
    const range = document.createRange()
    range.setStart(runElement, 0)
    range.setEnd(point.node, clampDomOffset(point.node, point.offset))
    return range.toString().length
  } catch {
    return null
  }
}

function collectRunElements(root: ParentNode): ReadonlyArray<HTMLElement> {
  const runElements: Array<HTMLElement> = []

  if (root instanceof HTMLElement && root.dataset.runIndex !== undefined) {
    runElements.push(root)
  }

  runElements.push(...Array.from(root.querySelectorAll<HTMLElement>('[data-run-index]')))

  return runElements
}

function getRunFragments(
  root: ParentNode,
  paragraphPath: ReadonlyArray<number>,
  runIndex: number,
  docModel?: DocxDocument,
): ReadonlyArray<HTMLElement> {
  return collectRunElements(root).filter((element) => {
    const elementRunIndex = parseInteger(element.dataset.runIndex)
    if (elementRunIndex !== runIndex) {
      return false
    }

    const elementParagraphPath = resolveElementParagraphPath(element, docModel)
    return elementParagraphPath !== null && paragraphPathsEqual(elementParagraphPath, paragraphPath)
  })
}

function sumPreviousFragmentLengths(
  fragments: ReadonlyArray<HTMLElement>,
  target: HTMLElement,
): number {
  let total = 0

  for (const fragment of fragments) {
    if (fragment === target) {
      break
    }

    total += getCharMetrics(fragment).length
  }

  return total
}

function resolveMappedTarget(point: NodePoint, docModel: DocxDocument | undefined): MappedTarget | null {
  const runElement = findClosestRunElement(point.node)
  if (!runElement) {
    return null
  }

  const paragraphPath =
    resolveElementParagraphPath(runElement, docModel) ?? resolveElementParagraphPath(point.node, docModel)
  if (!paragraphPath) {
    return null
  }

  const runIndex = parseInteger(runElement.dataset.runIndex)
  if (runIndex === null) {
    return null
  }

  return {
    point,
    runElement,
    paragraphPath,
    runIndex,
  }
}

function resolveDomPointWithin(
  element: HTMLElement,
  textOffset: number,
): Readonly<{
  node: Node
  offset: number
}> {
  const safeOffset = Math.max(0, textOffset)
  const walker = element.ownerDocument.createTreeWalker(element, NodeFilter.SHOW_TEXT)

  let remaining = safeOffset
  let currentNode = walker.nextNode()
  let lastTextNode: Text | null = null

  while (currentNode) {
    const textNode = currentNode as Text
    const textLength = textNode.data.length
    lastTextNode = textNode

    if (remaining <= textLength) {
      return {
        node: textNode,
        offset: remaining,
      }
    }

    remaining -= textLength
    currentNode = walker.nextNode()
  }

  if (lastTextNode) {
    return {
      node: lastTextNode,
      offset: lastTextNode.data.length,
    }
  }

  return {
    node: element,
    offset: Math.min(safeOffset, element.childNodes.length),
  }
}

export function domPointToPosition(
  node: Node,
  offset: number,
  document: Document,
  docModel?: DocxDocument,
): Position | null {
  if (node.ownerDocument !== document) {
    return null
  }

  for (const point of getCandidatePoints(node, offset)) {
    const target = resolveMappedTarget(point, docModel)
    if (!target) {
      continue
    }

    const localOffset = getTextOffsetWithin(target.runElement, target.point, document)
    if (localOffset === null) {
      continue
    }

    const metrics = getCharMetrics(target.runElement)
    const clampedOffset = Math.min(localOffset, metrics.length)

    const charOffset = metrics.start !== null
      ? metrics.start + clampedOffset
      : sumPreviousFragmentLengths(
          getRunFragments(document, target.paragraphPath, target.runIndex, docModel),
          target.runElement,
        ) + clampedOffset

    return {
      paragraphPath: target.paragraphPath,
      runIndex: target.runIndex,
      charOffset,
    }
  }

  return null
}

/**
 * USR-06 — model positions address a first-section paragraph either as
 * `[blockIndex]` (what the rendered DOM carries) or as `[0, blockIndex]`
 * (what Find, the toolbar and several commands produce). Try the position as
 * given, then without the leading section 0, so both forms map to the DOM.
 */
export function positionToDomRange(
  position: Position,
  root: HTMLElement,
  docModel?: DocxDocument,
): Readonly<{
  node: Node
  offset: number
}> | null {
  const direct = positionToDomRangeForPath(position, root, docModel)
  if (direct !== null || position.paragraphPath.length < 2 || position.paragraphPath[0] !== 0) {
    return direct
  }
  return positionToDomRangeForPath({ ...position, paragraphPath: position.paragraphPath.slice(1) }, root, docModel)
}

function positionToDomRangeForPath(
  position: Position,
  root: HTMLElement,
  docModel?: DocxDocument,
): Readonly<{
  node: Node
  offset: number
}> | null {
  const fragments = getRunFragments(root, position.paragraphPath, position.runIndex, docModel).map((element) => ({
    element,
    metrics: getCharMetrics(element),
  }))

  if (fragments.length === 0) {
    // An empty paragraph renders a line with no run spans: anchor the DOM
    // point on the line element itself so caret/selection sync still works.
    const emptyLine = findLineElements(root, docModel).find((line) =>
      paragraphPathsEqual(line.paragraphPath, position.paragraphPath),
    )
    return emptyLine === undefined ? null : { node: emptyLine.element, offset: 0 }
  }

  const explicitFragments = fragments.filter((fragment) => fragment.metrics.start !== null)
  if (explicitFragments.length > 0) {
    for (const fragment of explicitFragments) {
      const start = fragment.metrics.start ?? 0
      const end = start + fragment.metrics.length

      if (position.charOffset >= start && position.charOffset <= end) {
        return resolveDomPointWithin(fragment.element, position.charOffset - start)
      }
    }

    const lastFragment = explicitFragments[explicitFragments.length - 1]
    const start = lastFragment.metrics.start ?? 0
    return resolveDomPointWithin(
      lastFragment.element,
      Math.min(Math.max(position.charOffset - start, 0), lastFragment.metrics.length),
    )
  }

  let remaining = position.charOffset

  for (const fragment of fragments) {
    if (remaining <= fragment.metrics.length) {
      return resolveDomPointWithin(fragment.element, remaining)
    }

    remaining -= fragment.metrics.length
  }

  const lastFragment = fragments[fragments.length - 1]
  return resolveDomPointWithin(lastFragment.element, lastFragment.metrics.length)
}

// ─── Table-cell paragraph-path resolution (DOCX-17) ───────────────────────────
//
// Table cell content is rendered without a `data-paragraph-path` on its
// lines: `PageView.tsx`'s table row renderer replays each cell's laid-out
// lines through the same `renderLine` ordinary paragraphs use, but never
// threads a `paragraphPath` through for them, because the layout pipeline's
// per-cell line list (`LaidOutCell.contentLines`) is a flat array with no
// paragraph identity attached to it in the first place. That leaves every
// caret/selection helper below — which all key off `data-paragraph-path` —
// unable to resolve ANY position inside a table: a click lands on whatever
// non-table line happens to be nearest, and even Input.ts's already-correct
// Tab/Shift+Tab cell navigation (which builds a proper table-cell `Position`
// straight from the model) can't get the native caret to actually appear
// there, because `positionToDomRange` has nothing in the DOM to match it
// against.
//
// Fixing that at the source means teaching the layout/render layers to tag
// cell lines with their paragraph's real path — a
// `src/docx/layout/layoutTable.ts` / `PageView.tsx` change outside this
// file's ownership (see the DOCX-17 handoff notes). Instead, the functions
// below resolve a table paragraph's path independently, purely from data
// already in the rendered DOM (`.docx-page__table-wrapper[data-block-path]`,
// `tr[data-source-row]`, `td[data-col]`) plus the document model, replaying
// the SAME row/column accounting `layoutTable.ts` uses (only `'table-row'`/
// `'table-cell'` children count; columns accumulate `gridSpan`) so the two
// stay in step without sharing code.
//
// Known gap: a cell paragraph's line count isn't tracked either (the same
// flattening above), so when a cell holds more than one paragraph AND at
// least one of them wraps onto more than one line, a click on one of the
// extra wrapped lines can resolve to the wrong paragraph within that cell
// (clamped to the last one). A single paragraph per cell — by far the
// common case, and what every table fixture in this corpus uses — is always
// resolved correctly regardless of wrapping.

function gridSpanOf(cell: TableCell): number {
  const span = cell.props?.gridSpan
  return typeof span === 'number' && Number.isFinite(span) ? Math.max(1, Math.floor(span)) : 1
}

function resolveTableAtPath(doc: DocxDocument, tablePath: ReadonlyArray<number>): Table | null {
  if (tablePath.length === 0) {
    return null
  }

  let sectionIndex = 0
  let blockPath = tablePath
  if (tablePath.length > 1 && doc.sections[tablePath[0]] !== undefined) {
    sectionIndex = tablePath[0]
    blockPath = tablePath.slice(1)
  }

  const section = doc.sections[sectionIndex]
  if (section === undefined || blockPath.length !== 1) {
    return null
  }

  const block = section.blocks[blockPath[0]]
  return block !== undefined && block.kind === 'table' ? block : null
}

/**
 * The raw `table.rows` index of the `laidOutIndex`-th actual table row
 * (`layoutTable.ts` filters out any non-`'table-row'` child before counting
 * rows for its own `rowIndex`/`sourceRowIndex`, so this replays that same
 * count to invert it).
 */
function tableRowIndexAtLaidOutIndex(table: Table, laidOutIndex: number): number | null {
  let count = 0
  for (let index = 0; index < table.rows.length; index += 1) {
    if (table.rows[index].kind !== 'table-row') {
      continue
    }
    if (count === laidOutIndex) {
      return index
    }
    count += 1
  }
  return null
}

/** The raw `row.cells` index of the table cell whose grid position starts at
 * `columnStart` (a `<td>`'s own `data-col`), replaying `layoutTable.ts`'s
 * cumulative-`gridSpan` column accounting to invert it. */
function tableCellIndexAtColumn(row: TableRow, columnStart: number): number | null {
  let columnIndex = 0
  for (let index = 0; index < row.cells.length; index += 1) {
    const cell = row.cells[index]
    if (cell.kind !== 'table-cell') {
      continue
    }
    if (columnIndex === columnStart) {
      return index
    }
    columnIndex += gridSpanOf(cell)
  }
  return null
}

/**
 * The model `paragraphPath` for a `.docx-page__line` rendered inside a table
 * cell (`null` when `lineElement` isn't inside one, or the DOM/model don't
 * agree closely enough to resolve it — a merge-continuation cell that
 * renders no `<td>` at all, for instance). See this section's module
 * comment for the multi-paragraph-cell caveat.
 */
function resolveTableCellParagraphPath(
  docModel: DocxDocument,
  lineElement: HTMLElement,
): ReadonlyArray<number> | null {
  const cellElement = lineElement.closest<HTMLElement>('td')
  const rowElement = cellElement?.closest<HTMLElement>('tr[data-source-row]') ?? null
  const wrapperElement = lineElement.closest<HTMLElement>('.docx-page__table-wrapper[data-block-path]')
  if (cellElement === null || rowElement === null || wrapperElement === null) {
    return null
  }

  const tablePath = parseParagraphPath(wrapperElement.dataset.blockPath)
  const laidOutRowIndex = parseInteger(rowElement.dataset.sourceRow)
  const columnStart = parseInteger(cellElement.dataset.col)
  if (tablePath === null || laidOutRowIndex === null || columnStart === null) {
    return null
  }

  const table = resolveTableAtPath(docModel, tablePath)
  if (table === null) {
    return null
  }

  const rowIndex = tableRowIndexAtLaidOutIndex(table, laidOutRowIndex)
  const row = rowIndex === null ? undefined : table.rows[rowIndex]
  if (rowIndex === null || row === undefined || row.kind !== 'table-row') {
    return null
  }

  const cellIndex = tableCellIndexAtColumn(row, columnStart)
  const cell = cellIndex === null ? undefined : row.cells[cellIndex]
  if (cellIndex === null || cell === undefined || cell.kind !== 'table-cell') {
    return null
  }

  const paragraphCount = cell.blocks.filter((block) => block.kind === 'paragraph').length
  if (paragraphCount === 0) {
    return null
  }

  const linesInCell = Array.from(cellElement.querySelectorAll<HTMLElement>('.docx-page__line'))
  const lineIndex = linesInCell.indexOf(lineElement)
  const paragraphIndex = Math.min(Math.max(lineIndex, 0), paragraphCount - 1)

  return [...tablePath, rowIndex, cellIndex, paragraphIndex]
}

/**
 * `findClosestParagraphPath`, extended to resolve a table cell's paragraph
 * path (via `resolveTableCellParagraphPath`) when `node` has no
 * `data-paragraph-path` ancestor of its own — i.e. it's inside a table.
 * `docModel` is optional so every caller that doesn't have a document model
 * handy (or doesn't care about table cells) keeps the old, DOM-only
 * behavior exactly as before.
 */
function resolveElementParagraphPath(
  node: Node,
  docModel: DocxDocument | undefined,
): ReadonlyArray<number> | null {
  const direct = findClosestParagraphPath(node)
  if (direct !== null || docModel === undefined) {
    return direct
  }

  const line = getElementFromNode(node)?.closest<HTMLElement>('.docx-page__line') ?? null
  return line === null ? null : resolveTableCellParagraphPath(docModel, line)
}

// ─── Pointer hit-testing (USR-04 / USR-05) ────────────────────────────────────

export type PointerRange = Readonly<{ anchor: Position; focus: Position }>

type LineHit = Readonly<{
  line: HTMLElement
  paragraphPath: ReadonlyArray<number>
  spans: ReadonlyArray<HTMLElement>
}>

type ResolvedLine = Readonly<{ element: HTMLElement; paragraphPath: ReadonlyArray<number> }>

/**
 * Every rendered line with a resolvable paragraph path — ordinary
 * paragraphs via their own `data-paragraph-path`, table cell lines (which
 * carry no such attribute — see the DOCX-17 section above) via
 * `resolveTableCellParagraphPath` when `docModel` is supplied. A line whose
 * path can't be resolved either way (e.g. table lines when no `docModel` is
 * given) is left out, matching the old `[data-paragraph-path]` selector's
 * behavior for every caller that doesn't pass one.
 */
function findLineElements(root: ParentNode, docModel?: DocxDocument): ReadonlyArray<ResolvedLine> {
  const resolved: Array<ResolvedLine> = []
  for (const element of Array.from(root.querySelectorAll<HTMLElement>('.docx-page__line'))) {
    const direct = parseParagraphPath(element.dataset.paragraphPath)
    const paragraphPath = direct ?? (docModel === undefined ? null : resolveTableCellParagraphPath(docModel, element))
    if (paragraphPath !== null) {
      resolved.push({ element, paragraphPath })
    }
  }
  return resolved
}

function spanPosition(span: HTMLElement, paragraphPath: ReadonlyArray<number>, atEnd: boolean): Position | null {
  const runIndex = parseInteger(span.dataset.runIndex)
  const metrics = getCharMetrics(span)
  if (runIndex === null || metrics.start === null) {
    return null
  }
  return { paragraphPath, runIndex, charOffset: atEnd ? metrics.start + metrics.length : metrics.start }
}

/**
 * The rendered line closest to a client point: the line whose vertical band
 * contains `y` (nearest horizontally when several do, e.g. multi-column
 * pages), otherwise the line with the smallest vertical distance.
 */
function hitLine(x: number, y: number, root: HTMLElement, docModel?: DocxDocument): LineHit | null {
  let best: { line: ResolvedLine; score: number } | null = null
  for (const line of findLineElements(root, docModel)) {
    const rect = line.element.getBoundingClientRect()
    if (rect.height === 0 && rect.width === 0) {
      continue
    }
    const dy = y < rect.top ? rect.top - y : y > rect.bottom ? y - rect.bottom : 0
    const dx = x < rect.left ? rect.left - x : x > rect.right ? x - rect.right : 0
    const score = dy * 1000 + dx
    if (best === null || score < best.score) {
      best = { line, score }
    }
  }
  if (best === null) {
    return null
  }
  const spans = Array.from(best.line.element.querySelectorAll<HTMLElement>('[data-run-index][data-char-start]'))
  return { line: best.line.element, paragraphPath: best.line.paragraphPath, spans }
}

function offsetWithinSpan(span: HTMLElement, x: number): number {
  const metrics = getCharMetrics(span)
  const text = span.firstChild
  if (text === null || text.nodeType !== Node.TEXT_NODE || metrics.length === 0) {
    const rect = span.getBoundingClientRect()
    return x > rect.left + rect.width / 2 ? metrics.length : 0
  }
  const ownerDocument = span.ownerDocument
  const range = ownerDocument.createRange()
  const length = Math.min(metrics.length, text.textContent?.length ?? 0)
  for (let index = 0; index < length; index += 1) {
    range.setStart(text, index)
    range.setEnd(text, index + 1)
    const rect = range.getBoundingClientRect()
    if (x < rect.left + rect.width / 2) {
      return index
    }
  }
  return length
}

/**
 * Model position for a click at client coordinates. Unlike the browser's own
 * caret placement on this absolutely-positioned layout, a click to the right
 * of a line lands at the end of THAT line, to the left at its start, and a
 * click between lines on the nearest line.
 */
export function positionFromClientPoint(x: number, y: number, root: HTMLElement, docModel?: DocxDocument): Position | null {
  const hit = hitLine(x, y, root, docModel)
  if (hit === null) {
    return null
  }
  if (hit.spans.length === 0) {
    return { paragraphPath: hit.paragraphPath, runIndex: 0, charOffset: 0 }
  }

  const first = hit.spans[0]
  const last = hit.spans[hit.spans.length - 1]
  if (x <= first.getBoundingClientRect().left) {
    return spanPosition(first, hit.paragraphPath, false)
  }
  if (x >= last.getBoundingClientRect().right) {
    return spanPosition(last, hit.paragraphPath, true)
  }

  for (const span of hit.spans) {
    const rect = span.getBoundingClientRect()
    if (x <= rect.right) {
      const start = spanPosition(span, hit.paragraphPath, false)
      if (start === null) {
        return null
      }
      return { ...start, charOffset: start.charOffset + offsetWithinSpan(span, x) }
    }
  }
  return spanPosition(last, hit.paragraphPath, true)
}

/** The word (or whitespace run) under a client point, for double-click. */
export function wordRangeFromClientPoint(
  x: number,
  y: number,
  root: HTMLElement,
  docModel?: DocxDocument,
): PointerRange | null {
  const hit = hitLine(x, y, root, docModel)
  if (hit === null || hit.spans.length === 0) {
    const caret = positionFromClientPoint(x, y, root, docModel)
    return caret === null ? null : { anchor: caret, focus: caret }
  }
  const target =
    hit.spans.find((span) => {
      const rect = span.getBoundingClientRect()
      return x >= rect.left && x <= rect.right
    }) ?? (x < hit.spans[0].getBoundingClientRect().left ? hit.spans[0] : hit.spans[hit.spans.length - 1])
  const anchor = spanPosition(target, hit.paragraphPath, false)
  const focus = spanPosition(target, hit.paragraphPath, true)
  return anchor === null || focus === null ? null : { anchor, focus }
}

/** The whole paragraph under a client point (all of its lines), for triple-click. */
export function paragraphRangeFromClientPoint(
  x: number,
  y: number,
  root: HTMLElement,
  docModel?: DocxDocument,
): PointerRange | null {
  const hit = hitLine(x, y, root, docModel)
  if (hit === null) {
    return null
  }
  const spans = findLineElements(root, docModel)
    .filter((line) => paragraphPathsEqual(line.paragraphPath, hit.paragraphPath))
    .flatMap((line) => Array.from(line.element.querySelectorAll<HTMLElement>('[data-run-index][data-char-start]')))
  if (spans.length === 0) {
    const empty = { paragraphPath: hit.paragraphPath, runIndex: 0, charOffset: 0 }
    return { anchor: empty, focus: empty }
  }
  const anchor = spanPosition(spans[0], hit.paragraphPath, false)
  const focus = spanPosition(spans[spans.length - 1], hit.paragraphPath, true)
  return anchor === null || focus === null ? null : { anchor, focus }
}

export function findPositionAtClientPoint(
  x: number,
  y: number,
  document: Document,
  root: HTMLElement,
): Position | null {
  if (root.ownerDocument !== document) {
    return null
  }

  const caretDocument = document as CaretDocument
  const range = caretDocument.caretRangeFromPoint?.(x, y)
  if (range && root.contains(range.startContainer)) {
    return domPointToPosition(range.startContainer, range.startOffset, document)
  }

  const caretPosition = caretDocument.caretPositionFromPoint?.(x, y)
  if (caretPosition && root.contains(caretPosition.offsetNode)) {
    return domPointToPosition(caretPosition.offsetNode, caretPosition.offset, document)
  }

  return null
}
