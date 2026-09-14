/**
 * S1 (leaf placeholder fallback) / S7 (group transform composition) / S10
 * (crop, rotation, flip).
 *
 * `p:spTree` is walked depth-first so each shape's absolute position composes
 * every ancestor `p:grpSp`'s off/ext -> chOff/chExt translate+scale with the
 * shape's own local `xfrm`, per the OOXML group-transform algorithm — instead
 * of the previous flat, ancestry-blind scan that left grouped content bunched
 * near the slide origin.
 */

import { getAttributeValue, getDirectChildren, getFirstByLocalName } from '../shared/xmlUtils'
import { emuToPx, ooxmlAngleToDegrees } from '../shared/units'
import type { SlideTransform } from '../../shared/SlideDeck.types'
import { getPlaceholderRef, resolvePlaceholderPosition, type LayoutChain } from './layout'

const LEAF_TAGS = new Set(['sp', 'pic', 'graphicFrame', 'cxnSp'])

type RawBox = {
  readonly x: number
  readonly y: number
  readonly w: number
  readonly h: number
  readonly rotationDeg?: number
  readonly flipH?: boolean
  readonly flipV?: boolean
}

type RawGroupBox = RawBox & {
  readonly chOffX: number
  readonly chOffY: number
  readonly chExtW: number
  readonly chExtH: number
}

function readXfrmAttributes(xfrm: Element): RawBox {
  const offset = getFirstByLocalName(xfrm, 'off')
  const extent = getFirstByLocalName(xfrm, 'ext')

  return {
    x: emuToPx(getAttributeValue(offset ?? xfrm, ['x'])) ?? 0,
    y: emuToPx(getAttributeValue(offset ?? xfrm, ['y'])) ?? 0,
    w: emuToPx(getAttributeValue(extent ?? xfrm, ['cx'])) ?? 0,
    h: emuToPx(getAttributeValue(extent ?? xfrm, ['cy'])) ?? 0,
    rotationDeg: ooxmlAngleToDegrees(xfrm.getAttribute('rot')),
    flipH: xfrm.getAttribute('flipH') === '1',
    flipV: xfrm.getAttribute('flipV') === '1',
  }
}

/** A group's own `p:grpSpPr/a:xfrm`, including its `chOff`/`chExt` child coordinate space. */
function readGroupXfrm(group: Element): RawGroupBox | null {
  const grpSpPr = getFirstByLocalName(group, 'grpSpPr')
  const xfrm = grpSpPr ? getFirstByLocalName(grpSpPr, 'xfrm') : null
  if (!xfrm) {
    return null
  }

  const chOff = getFirstByLocalName(xfrm, 'chOff')
  const chExt = getFirstByLocalName(xfrm, 'chExt')

  return {
    ...readXfrmAttributes(xfrm),
    chOffX: emuToPx(getAttributeValue(chOff ?? xfrm, ['x'])) ?? 0,
    chOffY: emuToPx(getAttributeValue(chOff ?? xfrm, ['y'])) ?? 0,
    chExtW: emuToPx(getAttributeValue(chExt ?? xfrm, ['cx'])) ?? 0,
    chExtH: emuToPx(getAttributeValue(chExt ?? xfrm, ['cy'])) ?? 0,
  }
}

/** A leaf shape's own `spPr/a:xfrm` (`pic`/`sp`) or direct `p:xfrm` (`graphicFrame`/`cxnSp`). */
function readLeafOwnXfrm(shape: Element): RawBox | null {
  const spPr = getFirstByLocalName(shape, 'spPr')
  const xfrm = (spPr ? getFirstByLocalName(spPr, 'xfrm') : null) ?? getFirstByLocalName(shape, 'xfrm')
  return xfrm ? readXfrmAttributes(xfrm) : null
}

/**
 * A leaf shape's box in ITS OWN parent coordinate space (slide-absolute at
 * the root, or the enclosing group's child space when nested) — before any
 * group scaling is applied. Falls back to the layout/master placeholder
 * position (S1) when the shape carries no `xfrm` of its own.
 */
function resolveLeafOwnBox(shape: Element, chain: LayoutChain): RawBox | null {
  const own = readLeafOwnXfrm(shape)
  if (own && own.w > 0 && own.h > 0) {
    return own
  }

  if (shape.localName !== 'sp') {
    return own
  }

  const ref = getPlaceholderRef(shape)
  if (!ref) {
    return own
  }

  const inherited = resolvePlaceholderPosition(shape, ref, chain)
  if (inherited.x === undefined || inherited.y === undefined || inherited.w === undefined || inherited.h === undefined) {
    return own
  }

  return { ...inherited, rotationDeg: own?.rotationDeg, flipH: own?.flipH, flipV: own?.flipV } as RawBox
}

function composeWithParent(parent: SlideTransform, parentGroup: RawGroupBox, own: RawBox): SlideTransform {
  const scaleX = parentGroup.chExtW !== 0 ? parent.w / parentGroup.chExtW : 1
  const scaleY = parentGroup.chExtH !== 0 ? parent.h / parentGroup.chExtH : 1

  return {
    x: parent.x + (own.x - parentGroup.chOffX) * scaleX,
    y: parent.y + (own.y - parentGroup.chOffY) * scaleY,
    w: own.w * scaleX,
    h: own.h * scaleY,
    rotationDeg: own.rotationDeg,
    flipH: own.flipH,
    flipV: own.flipV,
  }
}

export type PositionedShape = {
  readonly element: Element
  readonly transform: SlideTransform
}

/**
 * Depth-first walk of a `p:spTree` (or, recursively, a `p:grpSp`), returning
 * every leaf shape (`sp`/`pic`/`graphicFrame`/`cxnSp`) with its transform
 * fully composed through every ancestor group. Malformed/unreadable shapes
 * are skipped rather than aborting the whole slide (S9 isolates failures at
 * the slide level; this isolates them at the shape level).
 */
export function traverseShapeTree(
  container: Element,
  chain: LayoutChain,
  parent: { readonly transform: SlideTransform; readonly group: RawGroupBox } | null = null,
): PositionedShape[] {
  const results: PositionedShape[] = []

  for (const child of getDirectChildren(container)) {
    try {
      if (child.localName === 'grpSp') {
        const raw = readGroupXfrm(child)
        if (!raw) {
          continue
        }

        const absolute = parent ? composeWithParent(parent.transform, parent.group, raw) : raw
        results.push(...traverseShapeTree(child, chain, { transform: absolute, group: raw }))
      } else if (LEAF_TAGS.has(child.localName)) {
        const own = resolveLeafOwnBox(child, chain)
        if (!own) {
          continue
        }

        const absolute = parent ? composeWithParent(parent.transform, parent.group, own) : own
        results.push({ element: child, transform: absolute })
      }
    } catch {
      // One malformed shape's transform must not drop every sibling.
      continue
    }
  }

  return results
}

export type ImageCrop = {
  readonly top: number
  readonly right: number
  readonly bottom: number
  readonly left: number
}

/** `a:srcRect`'s `l`/`t`/`r`/`b` crop percentages (1000ths of a percent) -> 0..1 fractions. */
export function resolveCrop(blipFill: Element | null): ImageCrop | undefined {
  const srcRect = blipFill ? getFirstByLocalName(blipFill, 'srcRect') : null
  if (!srcRect) {
    return undefined
  }

  const read = (name: string) => {
    const raw = srcRect.getAttribute(name)
    const parsed = raw !== null ? Number(raw) : 0
    return Number.isFinite(parsed) ? parsed / 100000 : 0
  }

  const crop = { top: read('t'), right: read('r'), bottom: read('b'), left: read('l') }
  return crop.top || crop.right || crop.bottom || crop.left ? crop : undefined
}
