/**
 * S7/S8/S19 — ODP shape transforms (`draw:g`/`draw:transform` group
 * composition), fill/border/geometry, and alt text.
 *
 * Unlike OOXML, ODF child shapes already carry page-absolute `svg:x`/`svg:y`
 * — a `draw:g` only contributes an additional `draw:transform` translate on
 * top, so composition here is a running (dx, dy) offset rather than a
 * chOff/chExt rescale.
 */

import { getDirectChildren, getFirstByLocalName } from '../shared/xmlUtils'
import { parseOdfLengthToPixels } from '../shared/units'
import type { SlideBorder, SlideFill, SlideGeometry, SlideTransform } from '../../shared/SlideDeck.types'
import { findInheritedProperties, resolveStyleChain, type OdpStyleIndex } from './styles'

const LEAF_TAGS = new Set(['frame', 'rect', 'ellipse', 'circle', 'custom-shape', 'line', 'connector-shape', 'path'])

type GroupOffset = { readonly dx: number; readonly dy: number }

function parseTranslateOffset(transformAttribute: string | null): GroupOffset {
  const match = transformAttribute?.match(/translate\(\s*([^\s,)]+)[\s,]+([^\s,)]+)\s*\)/)
  if (!match) {
    return { dx: 0, dy: 0 }
  }

  return {
    dx: parseOdfLengthToPixels(match[1]) ?? 0,
    dy: parseOdfLengthToPixels(match[2]) ?? 0,
  }
}

function readOwnBox(shape: Element): { x: number; y: number; w: number; h: number } | null {
  if (shape.localName === 'line' || shape.localName === 'connector-shape') {
    const x1 = parseOdfLengthToPixels(shape.getAttribute('svg:x1'))
    const y1 = parseOdfLengthToPixels(shape.getAttribute('svg:y1'))
    const x2 = parseOdfLengthToPixels(shape.getAttribute('svg:x2'))
    const y2 = parseOdfLengthToPixels(shape.getAttribute('svg:y2'))
    if (x1 === undefined || y1 === undefined || x2 === undefined || y2 === undefined) {
      return null
    }

    return { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1) || 1, h: Math.abs(y2 - y1) || 1 }
  }

  const x = parseOdfLengthToPixels(shape.getAttribute('svg:x')) ?? 0
  const y = parseOdfLengthToPixels(shape.getAttribute('svg:y')) ?? 0
  const w = parseOdfLengthToPixels(shape.getAttribute('svg:width')) ?? 0
  const h = parseOdfLengthToPixels(shape.getAttribute('svg:height')) ?? 0
  return { x, y, w, h }
}

export type OdpPositionedShape = {
  readonly element: Element
  readonly transform: SlideTransform
  /** USR-16 - nested in a draw:g, so its box is group-relative and it is not moved directly. */
  readonly inGroup: boolean
}

/** Depth-first walk of a `draw:page` (or, recursively, a `draw:g`) accumulating group translate offsets. */
export function traverseOdpShapes(
  container: Element,
  offset: GroupOffset = { dx: 0, dy: 0 },
  inGroup: boolean = false,
): OdpPositionedShape[] {
  const results: OdpPositionedShape[] = []

  for (const child of getDirectChildren(container)) {
    try {
      if (child.localName === 'g') {
        const own = parseTranslateOffset(child.getAttribute('draw:transform'))
        results.push(...traverseOdpShapes(child, { dx: offset.dx + own.dx, dy: offset.dy + own.dy }, true))
      } else if (LEAF_TAGS.has(child.localName)) {
        const box = readOwnBox(child)
        if (!box) {
          continue
        }

        results.push({
          element: child,
          transform: { x: box.x + offset.dx, y: box.y + offset.dy, w: box.w, h: box.h },
          inGroup,
        })
      }
    } catch {
      continue
    }
  }

  return results
}

const GEOMETRY_HINTS: ReadonlyArray<readonly [RegExp, SlideGeometry]> = [
  [/round/i, 'roundRect'],
  [/ellipse|circle/i, 'ellipse'],
  [/triangle/i, 'triangle'],
  [/right-arrow/i, 'rightArrow'],
  [/left-arrow/i, 'leftArrow'],
  [/up-arrow/i, 'upArrow'],
  [/down-arrow/i, 'downArrow'],
]

/** Maps a shape element to the shared `SlideGeometry` union from its tag name / `draw:enhanced-geometry@draw:type`. */
export function resolveOdpGeometry(shape: Element): SlideGeometry | undefined {
  if (shape.localName === 'rect') {
    return 'rect'
  }

  if (shape.localName === 'ellipse' || shape.localName === 'circle') {
    return 'ellipse'
  }

  if (shape.localName === 'line' || shape.localName === 'connector-shape') {
    return 'line'
  }

  if (shape.localName === 'custom-shape') {
    const geometry = getFirstByLocalName(shape, 'enhanced-geometry')
    const type = geometry?.getAttribute('draw:type') ?? ''
    const hint = GEOMETRY_HINTS.find(([pattern]) => pattern.test(type))
    return hint ? hint[1] : 'other'
  }

  return undefined
}

/** Resolves a shape's `draw:style-name` -> `style:graphic-properties` fill (S8). */
export function resolveOdpFill(styleName: string | null, index: OdpStyleIndex): SlideFill | undefined {
  const chain = resolveStyleChain(styleName, index)
  const properties = findInheritedProperties(chain, 'graphic-properties')
  if (!properties) {
    return undefined
  }

  const fillKind = properties.getAttribute('draw:fill')
  if (fillKind === 'none') {
    return { kind: 'none' }
  }

  const color = properties.getAttribute('draw:fill-color')
  if (fillKind === 'solid' && color) {
    return { kind: 'solid', color }
  }

  if (fillKind === 'gradient' && color) {
    return { kind: 'gradient', colors: [color] }
  }

  return undefined
}

/** Resolves a shape's `draw:style-name` -> `style:graphic-properties` stroke (S8). */
export function resolveOdpBorder(styleName: string | null, index: OdpStyleIndex): SlideBorder | undefined {
  const chain = resolveStyleChain(styleName, index)
  const properties = findInheritedProperties(chain, 'graphic-properties')
  if (!properties || properties.getAttribute('draw:stroke') === 'none') {
    return undefined
  }

  const color = properties.getAttribute('svg:stroke-color')
  if (!color) {
    return undefined
  }

  const widthPx = parseOdfLengthToPixels(properties.getAttribute('svg:stroke-width')) ?? 1
  return { color, widthPx }
}

/** `svg:title`/`svg:desc` direct children of a `draw:frame` (S19). */
export function resolveOdpAltText(frame: Element): string {
  const title = getFirstByLocalName(frame, 'title')
  const desc = getFirstByLocalName(frame, 'desc')
  return (title?.textContent ?? desc?.textContent ?? '').trim()
}
