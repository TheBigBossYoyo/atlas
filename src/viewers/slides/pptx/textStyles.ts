/**
 * S1 (continued) / S3 — run-level text formatting inheritance.
 *
 * OOXML resolves each run's effective formatting by merging, attribute by
 * attribute, from the most specific source down to the least specific:
 *   run `a:rPr` -> paragraph `a:pPr/a:defRPr` -> shape `a:lstStyle` level ->
 *   matching layout placeholder -> matching master placeholder ->
 *   master `p:txStyles` (title/body/other) level -> hardcoded default.
 * Only attributes missing from every higher-priority source are filled in by
 * a lower one — a paragraph is never wholesale replaced by an ancestor.
 */

import { getAttributeValue, getDirectChildren, getFirstByLocalName } from '../shared/xmlUtils'
import { hundredthsOfPointToPx } from '../shared/units'
import { resolveSolidFillColor, type ColorMap, type ThemeColors } from './theme'
import { findMatchingPlaceholder, type LayoutChain, type PlaceholderRef } from './layout'

export type RunFormatting = {
  readonly bold?: boolean
  readonly italic?: boolean
  readonly underline?: boolean
  readonly strikethrough?: boolean
  readonly color?: string
  readonly fontSizePx?: number
  readonly fontFamily?: string
}

const DEFAULT_BODY_FONT_SIZE_PX = 24

function readBooleanAttribute(element: Element, name: string, falseValue: string): boolean | undefined {
  const raw = element.getAttribute(name)
  if (raw === null) {
    return undefined
  }

  return raw !== falseValue && raw !== '0' && raw !== 'false'
}

/** Reads the direct formatting attributes on one `a:rPr`/`a:defRPr`-shaped element (no inheritance). */
function readOwnRunProps(
  element: Element | null,
  themeColors: ThemeColors,
  clrMap: ColorMap,
): RunFormatting {
  if (!element) {
    return {}
  }

  const props: RunFormatting = {}
  const bold = readBooleanAttribute(element, 'b', '0')
  if (bold !== undefined) {
    (props as { bold?: boolean }).bold = bold
  }

  const italic = readBooleanAttribute(element, 'i', '0')
  if (italic !== undefined) {
    (props as { italic?: boolean }).italic = italic
  }

  const underline = element.getAttribute('u')
  if (underline !== null) {
    (props as { underline?: boolean }).underline = underline !== 'none'
  }

  const strike = element.getAttribute('strike')
  if (strike !== null) {
    (props as { strikethrough?: boolean }).strikethrough = strike !== 'noStrike' && strike !== ''
  }

  const sizeAttribute = element.getAttribute('sz')
  const fontSizePx = hundredthsOfPointToPx(sizeAttribute)
  if (fontSizePx !== undefined) {
    (props as { fontSizePx?: number }).fontSizePx = fontSizePx
  }

  const color = resolveSolidFillColor(element, themeColors, clrMap)
  if (color) {
    (props as { color?: string }).color = color
  }

  const latin = getFirstByLocalName(element, 'latin')
  const typeface = latin?.getAttribute('typeface')
  if (typeface) {
    (props as { fontFamily?: string }).fontFamily = typeface
  }

  return props
}

/** Fills gaps in `base` from `fallback`, without ever overwriting an attribute `base` already set. */
function mergeRunProps(base: RunFormatting, fallback: RunFormatting): RunFormatting {
  return {
    bold: base.bold ?? fallback.bold,
    italic: base.italic ?? fallback.italic,
    underline: base.underline ?? fallback.underline,
    strikethrough: base.strikethrough ?? fallback.strikethrough,
    color: base.color ?? fallback.color,
    fontSizePx: base.fontSizePx ?? fallback.fontSizePx,
    fontFamily: base.fontFamily ?? fallback.fontFamily,
  }
}

/** `a:lstStyle/a:lvl{N+1}pPr/a:defRPr` for a 0-based paragraph level. */
function findLstStyleDefRPr(shape: Element, level: number): Element | null {
  const txBody = getFirstByLocalName(shape, 'txBody')
  const lstStyle = txBody ? getFirstByLocalName(txBody, 'lstStyle') : null
  if (!lstStyle) {
    return null
  }

  const lvlNode = getDirectChildren(lstStyle).find(child => child.localName === `lvl${level + 1}pPr`)
  return lvlNode ? getFirstByLocalName(lvlNode, 'defRPr') : null
}

/** A placeholder shape's first paragraph's `a:pPr/a:defRPr` — a stand-in for its authored sample formatting. */
function findFirstParagraphDefRPr(shape: Element): Element | null {
  const txBody = getFirstByLocalName(shape, 'txBody')
  if (!txBody) {
    return null
  }

  const paragraph = getDirectChildren(txBody, 'p')[0]
  if (!paragraph) {
    return null
  }

  const pPr = getFirstByLocalName(paragraph, 'pPr')
  return pPr ? getFirstByLocalName(pPr, 'defRPr') : null
}

type PlaceholderStyleFamily = 'title' | 'body' | 'other'

function classifyPlaceholderStyle(ref: PlaceholderRef): PlaceholderStyleFamily {
  switch (ref.type) {
    case 'title':
    case 'ctrTitle':
      return 'title'
    case 'body':
    case 'subTitle':
    case null:
      return 'body'
    default:
      return 'other'
  }
}

const TX_STYLE_TAG: Record<PlaceholderStyleFamily, string> = {
  title: 'titleStyle',
  body: 'bodyStyle',
  other: 'otherStyle',
}

/** `p:txStyles/{title|body|other}Style/a:lvl{N+1}pPr/a:defRPr` on the slide master. */
function findMasterTxStyleDefRPr(
  masterDocument: XMLDocument | null,
  family: PlaceholderStyleFamily,
  level: number,
): Element | null {
  if (!masterDocument) {
    return null
  }

  const txStyles = getFirstByLocalName(masterDocument, 'txStyles')
  const styleRoot = txStyles ? getFirstByLocalName(txStyles, TX_STYLE_TAG[family]) : null
  if (!styleRoot) {
    return null
  }

  const lvlNode = getDirectChildren(styleRoot).find(child => child.localName === `lvl${level + 1}pPr`)
  return lvlNode ? getFirstByLocalName(lvlNode, 'defRPr') : null
}

function defaultTextColor(themeColors: ThemeColors, clrMap: ColorMap): string {
  const slot = clrMap.get('tx1') ?? 'tx1'
  return themeColors.get(slot) ?? themeColors.get('dk1') ?? '#000000'
}

/**
 * Resolves one run's effective formatting, walking the full slide -> layout
 * -> master -> txStyles inheritance chain. `ownShape` is the run's own slide
 * shape (for its `a:lstStyle`); `ref` is that shape's placeholder reference,
 * or `null` for a non-placeholder shape (skips straight to hardcoded defaults
 * once the shape's own formatting is exhausted).
 */
export function resolveRunFormatting(
  runProps: Element | null,
  paragraphDefRPr: Element | null,
  ownShape: Element,
  ref: PlaceholderRef | null,
  level: number,
  chain: LayoutChain,
): RunFormatting {
  const { themeColors, clrMap } = chain
  const read = (element: Element | null) => readOwnRunProps(element, themeColors, clrMap)

  let resolved = mergeRunProps(read(runProps), read(paragraphDefRPr))
  resolved = mergeRunProps(resolved, read(findLstStyleDefRPr(ownShape, level)))

  if (ref) {
    const layoutMatch = findMatchingPlaceholder(chain.layoutDocument, ref)
    if (layoutMatch) {
      resolved = mergeRunProps(resolved, read(findLstStyleDefRPr(layoutMatch, level)))
      resolved = mergeRunProps(resolved, read(findFirstParagraphDefRPr(layoutMatch)))
    }

    const masterMatch = findMatchingPlaceholder(chain.masterDocument, ref)
    if (masterMatch) {
      resolved = mergeRunProps(resolved, read(findLstStyleDefRPr(masterMatch, level)))
      resolved = mergeRunProps(resolved, read(findFirstParagraphDefRPr(masterMatch)))
    }

    const family = classifyPlaceholderStyle(ref)
    resolved = mergeRunProps(resolved, read(findMasterTxStyleDefRPr(chain.masterDocument, family, level)))
  } else {
    resolved = mergeRunProps(resolved, read(findMasterTxStyleDefRPr(chain.masterDocument, 'other', level)))
  }

  return mergeRunProps(resolved, {
    color: defaultTextColor(themeColors, clrMap),
    fontSizePx: DEFAULT_BODY_FONT_SIZE_PX,
  })
}

/** 0-based `a:pPr@lvl` (OOXML defaults to level 0 when omitted). */
export function getParagraphLevel(paragraph: Element): number {
  const pPr = getFirstByLocalName(paragraph, 'pPr')
  const level = pPr ? getAttributeValue(pPr, ['lvl']) : null
  const parsed = level !== null ? Number(level) : 0
  return Number.isFinite(parsed) && parsed >= 0 ? parsed : 0
}
