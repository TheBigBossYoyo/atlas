/** S3/S4/S8 — builds the `SlideShape` for a `p:sp`/`p:cxnSp`: styled text, or a plain fill/border/geometry shape. */

import { getElementsByLocalName, getFirstByLocalName } from '../shared/xmlUtils'
import { emuToPx } from '../shared/units'
import type { SlideShape, SlideTextAnchor, SlideTextBody, SlideTransform } from '../../shared/SlideDeck.types'
import { resolveBorder, resolveFill, resolveGeometry } from './fills'
import { findMatchingPlaceholder, getPlaceholderRef, type LayoutChain } from './layout'
import { parseTextBody, resolveNormAutofitScale } from './text'

/** OOXML `a:bodyPr` default insets: 91440 EMU (0.1in) left/right, 45720 EMU (0.05in) top/bottom. */
const DEFAULT_INSET_X_PX = 9.6
const DEFAULT_INSET_Y_PX = 4.8

const ANCHORS: Readonly<Record<string, SlideTextAnchor>> = { t: 'top', ctr: 'middle', b: 'bottom' }

/** USR-16 — empty text placeholders stay on the slide (editable), showing PowerPoint's prompt only while editing. */
const PLACEHOLDER_PROMPTS: Readonly<Record<string, string>> = {
  title: 'Click to add title',
  ctrTitle: 'Click to add title',
  subTitle: 'Click to add subtitle',
  body: 'Click to add text',
  obj: 'Click to add text',
}

/**
 * USR-15 — resolves text-box layout from `a:bodyPr` elements ordered slide ->
 * layout -> master: the first element that sets an attribute wins. Titles
 * with no explicit anchor anywhere default to centered, as in PowerPoint's
 * built-in master.
 */
export function resolveTextBody(bodyPrs: ReadonlyArray<Element | null>, placeholderType: string | null): SlideTextBody {
  const read = (name: string): string | null => {
    for (const bodyPr of bodyPrs) {
      const value = bodyPr?.getAttribute(name)
      if (value !== null && value !== undefined) return value
    }
    return null
  }
  const inset = (name: string, fallback: number): number => emuToPx(read(name)) ?? fallback
  const isTitle = placeholderType === 'title' || placeholderType === 'ctrTitle'
  return {
    insets: {
      top: inset('tIns', DEFAULT_INSET_Y_PX),
      right: inset('rIns', DEFAULT_INSET_X_PX),
      bottom: inset('bIns', DEFAULT_INSET_Y_PX),
      left: inset('lIns', DEFAULT_INSET_X_PX),
    },
    anchor: ANCHORS[read('anchor') ?? ''] ?? (isTitle ? 'middle' : 'top'),
    wrap: read('wrap') !== 'none',
  }
}

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

  const placeholderPrompt = text.length === 0 && ref ? PLACEHOLDER_PROMPTS[ref.type ?? 'body'] : undefined

  // A manually inserted text box (`p:cNvSpPr txBox="1"`, as opposed to a
  // layout placeholder) is a real, user-placed object the moment it's
  // created — selectable and editable — even before it has any text, the
  // same way an empty title/body placeholder stays on the slide via
  // `placeholderPrompt` above. Without this, "Insert Text Box" followed by
  // typing immediately had nothing to type into: the brand-new shape starts
  // out with no text and (by design) `<a:noFill/>`/no border, so it fell
  // through to the `hasVisibleFill || border || ...` check below, which
  // returned `null` for it — an invisible, unselectable shape that a
  // pending caret could never resolve against.
  const isManualTextBox = shape.localName === 'sp' && getFirstByLocalName(shape, 'cNvSpPr')?.getAttribute('txBox') === '1'

  if (text.length > 0 || placeholderPrompt !== undefined || isManualTextBox) {
    const bodyPr = getFirstByLocalName(getFirstByLocalName(shape, 'txBody') ?? shape, 'bodyPr')
    const inheritedBodyPrs = ref
      ? [findMatchingPlaceholder(chain.layoutDocument, ref), findMatchingPlaceholder(chain.masterDocument, ref)].map(
          (placeholder) => (placeholder ? getFirstByLocalName(placeholder, 'bodyPr') : null),
        )
      : []
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
      body: resolveTextBody([bodyPr, ...inheritedBodyPrs], ref?.type ?? null),
      ...(placeholderPrompt !== undefined ? { placeholderPrompt } : {}),
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
