/** S8 — shape fill, border, and preset-geometry resolution, plus background inheritance. */

import { getDirectChildren, getFirstByLocalName } from '../shared/xmlUtils'
import { emuToPx } from '../shared/units'
import type { SlideBorder, SlideFill, SlideGeometry } from '../../shared/SlideDeck.types'
import { findColorElement, resolveColorElement, type ColorMap, type ThemeColors } from './theme'
import type { LayoutChain } from './layout'

const GEOMETRY_MAP: Record<string, SlideGeometry> = {
  rect: 'rect',
  roundRect: 'roundRect',
  round2SameRect: 'roundRect',
  round2DiagRect: 'roundRect',
  snip2SameRect: 'roundRect',
  ellipse: 'ellipse',
  triangle: 'triangle',
  rightArrow: 'rightArrow',
  leftArrow: 'leftArrow',
  upArrow: 'upArrow',
  downArrow: 'downArrow',
  line: 'line',
  straightConnector1: 'line',
  bentConnector2: 'line',
  bentConnector3: 'line',
}

/** `a:prstGeom@prst`, mapped to the shared `SlideGeometry` union (unknown presets -> `'other'`). */
export function resolveGeometry(spPr: Element | null): SlideGeometry | undefined {
  const prstGeom = spPr ? getFirstByLocalName(spPr, 'prstGeom') : null
  const prst = prstGeom?.getAttribute('prst')
  if (!prst) {
    return undefined
  }

  return GEOMETRY_MAP[prst] ?? 'other'
}

/**
 * `a:solidFill`/`a:gradFill`/`a:noFill` on `spPr` (or any container element that
 * holds one directly). Deliberately scoped to `container`'s DIRECT children —
 * `spPr` also carries an `a:ln` sibling (the shape's border/line), which can
 * carry its own `a:noFill`/`a:solidFill` for the *border*. A subtree search
 * (`getFirstByLocalName`) would find those too and mis-resolve the shape's own
 * fill from its border's fill (e.g. a solid-filled shape with "no line" would
 * wrongly resolve as `{ kind: 'none' }`, since `a:ln`'s `a:noFill` matches
 * first regardless of the shape's own `a:solidFill`).
 */
export function resolveFill(
  container: Element | null,
  themeColors: ThemeColors,
  clrMap: ColorMap,
): SlideFill | undefined {
  if (!container) {
    return undefined
  }

  const children = getDirectChildren(container)
  const findChild = (localName: string) => children.find(child => child.localName === localName) ?? null

  if (findChild('noFill')) {
    return { kind: 'none' }
  }

  const solidFill = findChild('solidFill')
  if (solidFill) {
    const colorElement = findColorElement(solidFill)
    if (colorElement) {
      return { kind: 'solid', color: resolveColorElement(colorElement, themeColors, clrMap) }
    }
  }

  const gradFill = findChild('gradFill')
  if (gradFill) {
    const gsLst = getFirstByLocalName(gradFill, 'gsLst')
    const stops = gsLst ? Array.from(gsLst.children).filter(child => child.localName === 'gs') : []
    const colors = stops
      .map(stop => {
        const colorElement = findColorElement(stop)
        return colorElement ? resolveColorElement(colorElement, themeColors, clrMap) : null
      })
      .filter((color): color is string => color !== null)

    if (colors.length > 0) {
      const lin = getFirstByLocalName(gradFill, 'lin')
      const angleRaw = lin?.getAttribute('ang')
      const angleDeg = angleRaw ? Number(angleRaw) / 60000 : undefined
      return { kind: 'gradient', colors, angleDeg }
    }
  }

  return undefined
}

/** `a:ln`'s width (`w`, EMU) and solid-fill color. */
export function resolveBorder(
  spPr: Element | null,
  themeColors: ThemeColors,
  clrMap: ColorMap,
): SlideBorder | undefined {
  const ln = spPr ? getFirstByLocalName(spPr, 'ln') : null
  if (!ln || getFirstByLocalName(ln, 'noFill')) {
    return undefined
  }

  const solidFill = getFirstByLocalName(ln, 'solidFill')
  const colorElement = solidFill ? findColorElement(solidFill) : null
  if (!colorElement) {
    return undefined
  }

  const widthPx = emuToPx(ln.getAttribute('w')) ?? 1
  return { color: resolveColorElement(colorElement, themeColors, clrMap), widthPx }
}

/**
 * Resolves a slide's background, walking slide `p:cSld/p:bg` -> layout ->
 * master (S8's background-inheritance acceptance criterion).
 */
export function resolveSlideBackground(
  slideDocument: XMLDocument,
  chain: LayoutChain,
): SlideFill | undefined {
  const slideBg = getFirstByLocalName(slideDocument, 'bg')
  const layoutBg = chain.layoutDocument ? getFirstByLocalName(chain.layoutDocument, 'bg') : null
  const masterBg = chain.masterDocument ? getFirstByLocalName(chain.masterDocument, 'bg') : null

  for (const bg of [slideBg, layoutBg, masterBg]) {
    if (!bg) {
      continue
    }

    const bgPr = getFirstByLocalName(bg, 'bgPr')
    const fill = resolveFill(bgPr ?? bg, chain.themeColors, chain.clrMap)
    if (fill) {
      return fill
    }
  }

  return undefined
}
