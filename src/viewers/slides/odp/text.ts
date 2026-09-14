/**
 * S3/S4/S5/S21 — ODP paragraph/run extraction: styled `text:span` runs with
 * resolved formatting, `text:list`/`text:list-item` nesting for bullets, and
 * `text:line-break` for soft breaks (walked in document order, mirroring the
 * PPTX `<a:br>` fix, instead of a flat `.join('')` over every descendant).
 */

import { getDirectChildren, getFirstByLocalName } from '../shared/xmlUtils'
import { parseOdfLengthToPixels } from '../shared/units'
import type { SlideBullet, SlideParagraph, SlideParagraphAlign, SlideTextRun } from '../../shared/SlideDeck.types'
import { findInheritedProperties, resolveStyleChain, type OdpStyleIndex } from './styles'

type RunFormatting = {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strikethrough?: boolean
  color?: string
  fontSizePx?: number
  fontFamily?: string
}

function mergeGaps(base: RunFormatting, fallback: RunFormatting): RunFormatting {
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

/** Reads one `style:text-properties` element's own attributes (no inheritance). */
function readOwnTextProperties(properties: Element | null): RunFormatting {
  if (!properties) {
    return {}
  }

  const result: RunFormatting = {}
  const weight = properties.getAttribute('fo:font-weight')
  if (weight) {
    result.bold = weight === 'bold' || Number(weight) >= 700
  }

  const style = properties.getAttribute('fo:font-style')
  if (style) {
    result.italic = style === 'italic' || style === 'oblique'
  }

  const underline = properties.getAttribute('style:text-underline-style')
  if (underline) {
    result.underline = underline !== 'none'
  }

  const strike = properties.getAttribute('style:text-line-through-style')
  if (strike) {
    result.strikethrough = strike !== 'none'
  }

  const color = properties.getAttribute('fo:color')
  if (color) {
    result.color = color
  }

  const fontSize = properties.getAttribute('fo:font-size')
  if (fontSize) {
    result.fontSizePx = parseOdfLengthToPixels(fontSize)
  }

  const fontFamily = properties.getAttribute('style:font-name') ?? properties.getAttribute('fo:font-family')
  if (fontFamily) {
    result.fontFamily = fontFamily
  }

  return result
}

/** Resolves a style name's effective text formatting, walking its `style:parent-style-name` chain. */
function resolveTextFormatting(styleName: string | null, index: OdpStyleIndex): RunFormatting {
  const chain = resolveStyleChain(styleName, index)
  let resolved: RunFormatting = {}

  for (const style of chain) {
    resolved = mergeGaps(resolved, readOwnTextProperties(getFirstByLocalName(style, 'text-properties')))
  }

  return resolved
}

const ALIGN_MAP: Record<string, SlideParagraphAlign> = {
  start: 'left',
  left: 'left',
  center: 'center',
  end: 'right',
  right: 'right',
  justify: 'justify',
}

function resolveParagraphStyleProps(styleName: string | null, index: OdpStyleIndex) {
  const chain = resolveStyleChain(styleName, index)
  const paragraphProperties = findInheritedProperties(chain, 'paragraph-properties')
  const align = paragraphProperties?.getAttribute('fo:text-align')
  const listStyleName = chain[0]?.getAttribute('style:list-style-name') ?? null

  return {
    align: align ? ALIGN_MAP[align] : undefined,
    spaceBeforePx: parseOdfLengthToPixels(paragraphProperties?.getAttribute('fo:margin-top') ?? null),
    spaceAfterPx: parseOdfLengthToPixels(paragraphProperties?.getAttribute('fo:margin-bottom') ?? null),
    defaultTextFormatting: resolveTextFormatting(styleName, index),
    listStyleName,
  }
}

function resolveBullet(listStyleName: string | null, level: number, index: OdpStyleIndex): SlideBullet | undefined {
  const listStyle = listStyleName ? index.listStyles.get(listStyleName) : undefined
  if (!listStyle) {
    return undefined
  }

  const levelElement = getDirectChildren(listStyle).find(child => child.getAttribute('text:level') === String(level))
  if (!levelElement) {
    return undefined
  }

  if (levelElement.localName === 'list-level-style-bullet') {
    return { level: level - 1, char: levelElement.getAttribute('text:bullet-char') ?? undefined }
  }

  if (levelElement.localName === 'list-level-style-number') {
    return { level: level - 1, numbered: true }
  }

  return undefined
}

function parseParagraphRuns(paragraph: Element, defaultFormatting: RunFormatting, index: OdpStyleIndex): {
  readonly runs: ReadonlyArray<SlideTextRun>
  readonly text: string
} {
  const runs: SlideTextRun[] = []
  let text = ''

  const appendText = (value: string, styleName: string | null) => {
    if (!value) {
      return
    }

    const formatting = styleName ? mergeGaps(resolveTextFormatting(styleName, index), defaultFormatting) : defaultFormatting
    runs.push({ text: value, ...formatting })
    text += value
  }

  const walk = (node: Element) => {
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType === Node.TEXT_NODE) {
        appendText(child.textContent ?? '', null)
      } else if (child.nodeType === Node.ELEMENT_NODE) {
        const element = child as Element
        if (element.localName === 'line-break') {
          runs.push({ text: '\n' })
          text += '\n'
        } else if (element.localName === 'span') {
          const styleName = element.getAttribute('text:style-name')
          const before = text
          const nestedRuns: SlideTextRun[] = []
          const nestedFormatting = styleName ? mergeGaps(resolveTextFormatting(styleName, index), defaultFormatting) : defaultFormatting
          for (const spanChild of Array.from(element.childNodes)) {
            if (spanChild.nodeType === Node.TEXT_NODE) {
              const value = spanChild.textContent ?? ''
              if (value) {
                nestedRuns.push({ text: value, ...nestedFormatting })
              }
            } else if (spanChild.nodeType === Node.ELEMENT_NODE && (spanChild as Element).localName === 'line-break') {
              nestedRuns.push({ text: '\n' })
            }
          }
          runs.push(...nestedRuns)
          text = before + nestedRuns.map(run => run.text).join('')
        } else if (element.localName === 's' || element.localName === 'tab') {
          appendText(element.localName === 'tab' ? '\t' : ' ', null)
        }
      }
    }
  }

  walk(paragraph)
  return { runs, text }
}

export type ParsedOdpTextBody = {
  readonly paragraphs: ReadonlyArray<SlideParagraph>
  readonly text: string
}

/**
 * Parses a text container (`draw:text-box`, `table:table-cell`, ...) into
 * paragraphs, recursing through `text:list`/`text:list-item` nesting for
 * bullet levels (S4) and resolving each paragraph's own runs (S3) and soft
 * line breaks (S5).
 */
export function parseOdpTextBody(
  container: Element,
  index: OdpStyleIndex,
  ownListStyleName: string | null = null,
  level = 0,
): ParsedOdpTextBody {
  const paragraphs: SlideParagraph[] = []
  const textParts: string[] = []

  for (const child of getDirectChildren(container)) {
    if (child.localName === 'p' || child.localName === 'h') {
      const styleName = child.getAttribute('text:style-name')
      const styleProps = resolveParagraphStyleProps(styleName, index)
      const { runs, text } = parseParagraphRuns(child, styleProps.defaultTextFormatting, index)
      const effectiveListStyleName = styleProps.listStyleName ?? ownListStyleName
      const outlineLevelAttr = child.getAttribute('text:outline-level')
      // `level` already counts the enclosing text:list nesting (incremented on recursion
      // below), so a paragraph directly inside the first list is already at level 1 —
      // only a paragraph with no enclosing list at all (level 0) needs bumping to 1.
      const effectiveLevel = outlineLevelAttr ? Number(outlineLevelAttr) : Math.max(level, 1)

      paragraphs.push({
        runs,
        level: Math.max(0, effectiveLevel - 1),
        bullet: resolveBullet(effectiveListStyleName, effectiveLevel, index),
        align: styleProps.align,
        spaceBeforePx: styleProps.spaceBeforePx,
        spaceAfterPx: styleProps.spaceAfterPx,
      })
      textParts.push(text)
    } else if (child.localName === 'list') {
      const listStyleName = child.getAttribute('text:style-name') ?? ownListStyleName
      for (const item of getDirectChildren(child, 'list-item')) {
        const nested = parseOdpTextBody(item, index, listStyleName, level + 1)
        paragraphs.push(...nested.paragraphs)
        textParts.push(nested.text)
      }
    }
  }

  return { paragraphs, text: textParts.join('\n').trim() }
}
