import type { Position } from './Selection'

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
): ReadonlyArray<HTMLElement> {
  return collectRunElements(root).filter((element) => {
    const elementRunIndex = parseInteger(element.dataset.runIndex)
    if (elementRunIndex !== runIndex) {
      return false
    }

    const elementParagraphPath = findClosestParagraphPath(element)
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

function resolveMappedTarget(point: NodePoint): MappedTarget | null {
  const runElement = findClosestRunElement(point.node)
  if (!runElement) {
    return null
  }

  const paragraphPath = findClosestParagraphPath(runElement) ?? findClosestParagraphPath(point.node)
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
): Position | null {
  if (node.ownerDocument !== document) {
    return null
  }

  for (const point of getCandidatePoints(node, offset)) {
    const target = resolveMappedTarget(point)
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
          getRunFragments(document, target.paragraphPath, target.runIndex),
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
): Readonly<{
  node: Node
  offset: number
}> | null {
  const direct = positionToDomRangeForPath(position, root)
  if (direct !== null || position.paragraphPath.length < 2 || position.paragraphPath[0] !== 0) {
    return direct
  }
  return positionToDomRangeForPath({ ...position, paragraphPath: position.paragraphPath.slice(1) }, root)
}

function positionToDomRangeForPath(
  position: Position,
  root: HTMLElement,
): Readonly<{
  node: Node
  offset: number
}> | null {
  const fragments = getRunFragments(root, position.paragraphPath, position.runIndex).map((element) => ({
    element,
    metrics: getCharMetrics(element),
  }))

  if (fragments.length === 0) {
    // An empty paragraph renders a line with no run spans: anchor the DOM
    // point on the line element itself so caret/selection sync still works.
    const emptyLine = findLineElements(root).find((line) => {
      const linePath = parseParagraphPath(line.dataset.paragraphPath)
      return linePath !== null && paragraphPathsEqual(linePath, position.paragraphPath)
    })
    return emptyLine === undefined ? null : { node: emptyLine, offset: 0 }
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

// ─── Pointer hit-testing (USR-04 / USR-05) ────────────────────────────────────

export type PointerRange = Readonly<{ anchor: Position; focus: Position }>

type LineHit = Readonly<{
  line: HTMLElement
  paragraphPath: ReadonlyArray<number>
  spans: ReadonlyArray<HTMLElement>
}>

function findLineElements(root: ParentNode): ReadonlyArray<HTMLElement> {
  return Array.from(root.querySelectorAll<HTMLElement>('.docx-page__line[data-paragraph-path]'))
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
function hitLine(x: number, y: number, root: HTMLElement): LineHit | null {
  let best: { line: HTMLElement; score: number } | null = null
  for (const line of findLineElements(root)) {
    const rect = line.getBoundingClientRect()
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
  const paragraphPath = parseParagraphPath(best.line.dataset.paragraphPath)
  if (paragraphPath === null) {
    return null
  }
  const spans = Array.from(best.line.querySelectorAll<HTMLElement>('[data-run-index][data-char-start]'))
  return { line: best.line, paragraphPath, spans }
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
export function positionFromClientPoint(x: number, y: number, root: HTMLElement): Position | null {
  const hit = hitLine(x, y, root)
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
export function wordRangeFromClientPoint(x: number, y: number, root: HTMLElement): PointerRange | null {
  const hit = hitLine(x, y, root)
  if (hit === null || hit.spans.length === 0) {
    const caret = positionFromClientPoint(x, y, root)
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
export function paragraphRangeFromClientPoint(x: number, y: number, root: HTMLElement): PointerRange | null {
  const hit = hitLine(x, y, root)
  if (hit === null) {
    return null
  }
  const spans = findLineElements(root)
    .filter((line) => {
      const linePath = parseParagraphPath(line.dataset.paragraphPath)
      return linePath !== null && paragraphPathsEqual(linePath, hit.paragraphPath)
    })
    .flatMap((line) => Array.from(line.querySelectorAll<HTMLElement>('[data-run-index][data-char-start]')))
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
