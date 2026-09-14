/** S3/S4/S8 — builds the `SlideShape` for a `p:sp`/`p:cxnSp`: styled text, or a plain fill/border/geometry shape. */

import { getElementsByLocalName, getFirstByLocalName } from '../shared/xmlUtils'
import type { SlideShape, SlideTransform } from '../../shared/SlideDeck.types'
import { resolveBorder, resolveFill, resolveGeometry } from './fills'
import { getPlaceholderRef, type LayoutChain } from './layout'
import { parseTextBody, resolveNormAutofitScale } from './text'

/** Builds a text/shape `SlideShape` for `sp`/`cxnSp`, or `null` for a shape with neither visible fill nor text. */
export function buildShapeElement(
  shape: Element,
  id: string,
  transform: SlideTransform,
  chain: LayoutChain,
  consumedTextNodes: Set<Element>,
): SlideShape | null {
  for (const textNode of getElementsByLocalName(shape, 't')) {
    consumedTextNodes.add(textNode)
  }

  const spPr = getFirstByLocalName(shape, 'spPr')
  const fill = resolveFill(spPr, chain.themeColors, chain.clrMap)
  const border = resolveBorder(spPr, chain.themeColors, chain.clrMap)
  const geometry = resolveGeometry(spPr)

  const ref = shape.localName === 'sp' ? getPlaceholderRef(shape) : null
  const { paragraphs, text, fontScale } = parseTextBody(shape, ref, chain)

  if (text.length > 0) {
    const bodyPr = getFirstByLocalName(getFirstByLocalName(shape, 'txBody') ?? shape, 'bodyPr')
    return {
      kind: 'text',
      id,
      transform,
      paragraphs,
      text,
      fill,
      border,
      geometry,
      fontScale: fontScale ?? resolveNormAutofitScale(bodyPr),
      placeholderType: ref?.type ?? undefined,
    }
  }

  const hasVisibleFill = fill !== undefined && fill.kind !== 'none'
  if (hasVisibleFill || border || (geometry && geometry !== 'rect')) {
    return { kind: 'shape', id, transform, fill, border, geometry }
  }

  return null
}

/** Extracts a slide's title-placeholder text for nav labels (S17), or `undefined` when none is found. */
export function extractSlideTitle(slideDocument: XMLDocument): string | undefined {
  for (const shape of getElementsByLocalName(slideDocument, 'sp')) {
    const ph = getFirstByLocalName(shape, 'ph')
    const type = ph?.getAttribute('type')
    if (type !== 'title' && type !== 'ctrTitle') {
      continue
    }

    const text = getElementsByLocalName(shape, 't')
      .map(node => node.textContent ?? '')
      .join(' ')
      .trim()

    if (text) {
      return text
    }
  }

  return undefined
}
