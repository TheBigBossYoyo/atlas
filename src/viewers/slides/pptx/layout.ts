/**
 * S1 — slide -> layout -> master resolution and OOXML placeholder matching.
 *
 * A slide's `p:ph` placeholders inherit position/size and default text
 * formatting from the matching placeholder in its `slideLayout`, falling
 * back to its `slideMaster` when the layout itself has no match — the
 * standard OOXML placeholder-matching algorithm (match by `type` + `idx`).
 */

import type { CancelSignal, ZipArchive } from '../shared/xmlUtils'
import {
  getAttributeValue,
  getElementsByLocalName,
  getFirstByLocalName,
  parseXml,
  readZipText,
} from '../shared/xmlUtils'
import { emuToPx } from '../shared/units'
import { parseClrMap, parseThemeColors, type ColorMap, type ThemeColors } from './theme'
import { findRelationshipTarget, parseRelationships } from './relationships'

export type LayoutChain = {
  readonly layoutDocument: XMLDocument | null
  readonly masterDocument: XMLDocument | null
  readonly themeColors: ThemeColors
  readonly clrMap: ColorMap
}

const EMPTY_CHAIN: LayoutChain = {
  layoutDocument: null,
  masterDocument: null,
  themeColors: new Map(),
  clrMap: new Map(),
}

async function loadXmlPart(zip: ZipArchive, path: string, signal: CancelSignal): Promise<XMLDocument | null> {
  const xml = await readZipText(zip, path, signal)
  return xml ? parseXml(xml) : null
}

export async function resolveLayoutChain(
  zip: ZipArchive,
  slidePath: string,
  signal: CancelSignal,
): Promise<LayoutChain> {
  const slideRels = await parseRelationships(zip, slidePath, signal)
  const layoutPath = findRelationshipTarget(slideRels, '/slideLayout')
  if (!layoutPath) {
    return EMPTY_CHAIN
  }

  const layoutDocument = await loadXmlPart(zip, layoutPath, signal)
  if (signal.cancelled || !layoutDocument) {
    return EMPTY_CHAIN
  }

  const layoutRels = await parseRelationships(zip, layoutPath, signal)
  const masterPath = findRelationshipTarget(layoutRels, '/slideMaster')
  const masterDocument = masterPath ? await loadXmlPart(zip, masterPath, signal) : null

  let themeColors: ThemeColors = new Map()
  if (masterPath) {
    const masterRels = await parseRelationships(zip, masterPath, signal)
    const themePath = findRelationshipTarget(masterRels, '/theme')
    const themeDocument = themePath ? await loadXmlPart(zip, themePath, signal) : null
    themeColors = parseThemeColors(themeDocument)
  }

  return {
    layoutDocument,
    masterDocument,
    themeColors,
    clrMap: parseClrMap(masterDocument),
  }
}

export type PlaceholderRef = {
  readonly type: string | null
  readonly idx: string | null
}

/** Reads a shape's own `<p:nvPr><p:ph type=".." idx=".."/></p:nvPr>`, if any. */
export function getPlaceholderRef(shape: Element): PlaceholderRef | null {
  const ph = getFirstByLocalName(shape, 'ph')
  if (!ph) {
    return null
  }

  return { type: ph.getAttribute('type'), idx: ph.getAttribute('idx') }
}

/** Finds the shape in `container` whose own `p:ph` matches `ref` (type first, else idx, else any of the same family). */
export function findMatchingPlaceholder(container: XMLDocument | Element | null, ref: PlaceholderRef): Element | null {
  if (!container) {
    return null
  }

  const candidates = getElementsByLocalName(container, 'sp').filter(shape => getFirstByLocalName(shape, 'ph') !== null)

  if (ref.idx !== null) {
    const byIdx = candidates.find(shape => getFirstByLocalName(shape, 'ph')?.getAttribute('idx') === ref.idx)
    if (byIdx) {
      return byIdx
    }
  }

  if (ref.type !== null) {
    const byType = candidates.find(shape => getFirstByLocalName(shape, 'ph')?.getAttribute('type') === ref.type)
    if (byType) {
      return byType
    }
  }

  // A body placeholder with no explicit type defaults to "body"; match the layout's
  // first untyped/"body" placeholder as a last resort so position still inherits.
  if (ref.type === null) {
    return candidates.find(shape => {
      const type = getFirstByLocalName(shape, 'ph')?.getAttribute('type')
      return type === null || type === 'body'
    }) ?? null
  }

  return null
}

export type ResolvedPosition = {
  readonly x?: number
  readonly y?: number
  readonly w?: number
  readonly h?: number
}

function readXfrmBox(shape: Element): ResolvedPosition {
  const spPr = getFirstByLocalName(shape, 'spPr')
  const transform = spPr ? getFirstByLocalName(spPr, 'xfrm') : null
  if (!transform) {
    return {}
  }

  const offset = getFirstByLocalName(transform, 'off')
  const extent = getFirstByLocalName(transform, 'ext')

  return {
    x: emuToPx(getAttributeValue(offset ?? transform, ['x'])),
    y: emuToPx(getAttributeValue(offset ?? transform, ['y'])),
    w: emuToPx(getAttributeValue(extent ?? transform, ['cx'])),
    h: emuToPx(getAttributeValue(extent ?? transform, ['cy'])),
  }
}

/**
 * Resolves a placeholder's position/size by walking slide -> layout -> master,
 * taking the first fully-specified box found (OOXML never partially merges
 * x/y/w/h across levels for a single placeholder in practice).
 */
export function resolvePlaceholderPosition(
  slideShape: Element,
  ref: PlaceholderRef,
  chain: LayoutChain,
): ResolvedPosition {
  const ownBox = readXfrmBox(slideShape)
  if (ownBox.x !== undefined && ownBox.y !== undefined && ownBox.w !== undefined && ownBox.h !== undefined) {
    return ownBox
  }

  const layoutMatch = findMatchingPlaceholder(chain.layoutDocument, ref)
  if (layoutMatch) {
    const layoutBox = readXfrmBox(layoutMatch)
    if (layoutBox.x !== undefined) {
      return { ...layoutBox, ...ownBox }
    }
  }

  const masterMatch = findMatchingPlaceholder(chain.masterDocument, ref)
  if (masterMatch) {
    const masterBox = readXfrmBox(masterMatch)
    return { ...masterBox, ...ownBox }
  }

  return ownBox
}
