/**
 * S3/S4/S5/S20/S21 — paragraph/run extraction for a shape's `p:txBody`.
 *
 * Walks each paragraph's children in document order (not just `<a:t>`
 * descendants) so `<a:br>` soft line breaks land between the correct runs
 * instead of being lost by a flat `.join('')` (S5's empirically-confirmed
 * "QuarterlyReportQ3 2024" bug). Each run carries fully-resolved formatting
 * (S3) and each paragraph its bullet (S4) and spacing (S21).
 */

import { getDirectChildren, getFirstByLocalName, parseNumber } from '../shared/xmlUtils'
import { hundredthsOfPointToPx } from '../shared/units'
import type { SlideParagraph, SlideParagraphAlign, SlideTextRun } from '../../shared/SlideDeck.types'
import { resolveBullet } from './bullets'
import type { LayoutChain, PlaceholderRef } from './layout'
import { getParagraphLevel, resolveRunFormatting } from './textStyles'

export type ParsedTextBody = {
  readonly paragraphs: ReadonlyArray<SlideParagraph>
  /** Flattened plain text (paragraphs joined by `\n`) for search/nav/tests. */
  readonly text: string
  /** `a:normAutofit`'s `fontScale`, as a 0..1 multiplier — undefined when absent. */
  readonly fontScale?: number
}

const ALIGN_MAP: Record<string, SlideParagraphAlign> = {
  l: 'left',
  ctr: 'center',
  r: 'right',
  just: 'justify',
  justLow: 'justify',
}

function resolveAlign(pPr: Element | null): SlideParagraphAlign | undefined {
  const raw = pPr?.getAttribute('algn')
  return raw ? ALIGN_MAP[raw] : undefined
}

/** `a:spcBef`/`a:spcAft`'s `a:spcPts@val` (hundredths of a point) -> px. Percent-based spacing is left unresolved. */
function resolveSpacingPx(pPr: Element | null, tag: 'spcBef' | 'spcAft'): number | undefined {
  const container = pPr ? getFirstByLocalName(pPr, tag) : null
  const spcPts = container ? getFirstByLocalName(container, 'spcPts') : null
  return spcPts ? hundredthsOfPointToPx(spcPts.getAttribute('val')) : undefined
}

/** `a:normAutofit@fontScale` is in thousandths of a percent (92500 -> 0.925). */
export function resolveNormAutofitScale(bodyPr: Element | null): number | undefined {
  const normAutofit = bodyPr ? getFirstByLocalName(bodyPr, 'normAutofit') : null
  const raw = normAutofit?.getAttribute('fontScale')
  const parsed = raw !== undefined && raw !== null ? parseNumber(raw) : undefined
  return parsed === undefined ? undefined : parsed / 100000
}

function parseRun(
  run: Element,
  paragraphDefRPr: Element | null,
  ownShape: Element,
  ref: PlaceholderRef | null,
  level: number,
  chain: LayoutChain,
): SlideTextRun | null {
  const textNode = getFirstByLocalName(run, 't')
  const text = textNode?.textContent ?? ''
  if (!text) {
    return null
  }

  const rPr = getFirstByLocalName(run, 'rPr')
  const formatting = resolveRunFormatting(rPr, paragraphDefRPr, ownShape, ref, level, chain)
  return { text, ...formatting }
}

function parseParagraph(
  paragraph: Element,
  ownShape: Element,
  ref: PlaceholderRef | null,
  chain: LayoutChain,
): { readonly paragraph: SlideParagraph; readonly text: string } {
  const level = getParagraphLevel(paragraph)
  const pPr = getFirstByLocalName(paragraph, 'pPr')
  const paragraphDefRPr = pPr ? getFirstByLocalName(pPr, 'defRPr') : null

  const runs: SlideTextRun[] = []
  let text = ''

  for (const child of getDirectChildren(paragraph)) {
    if (child.localName === 'r' || child.localName === 'fld') {
      const run = parseRun(child, paragraphDefRPr, ownShape, ref, level, chain)
      if (run) {
        runs.push(run)
        text += run.text
      }
    } else if (child.localName === 'br') {
      runs.push({ text: '\n' })
      text += '\n'
    }
  }

  return {
    paragraph: {
      runs,
      level,
      bullet: resolveBullet(pPr, level, ownShape, ref, chain),
      align: resolveAlign(pPr),
      spaceBeforePx: resolveSpacingPx(pPr, 'spcBef'),
      spaceAfterPx: resolveSpacingPx(pPr, 'spcAft'),
    },
    text,
  }
}

/** Parses a shape's `p:txBody` into paragraphs with fully-resolved run formatting. */
export function parseTextBody(
  shape: Element,
  ref: PlaceholderRef | null,
  chain: LayoutChain,
): ParsedTextBody {
  const txBody = getFirstByLocalName(shape, 'txBody')
  if (!txBody) {
    return { paragraphs: [], text: '' }
  }

  const bodyPr = getFirstByLocalName(txBody, 'bodyPr')
  const fontScale = resolveNormAutofitScale(bodyPr)

  const paragraphs: SlideParagraph[] = []
  const flatParts: string[] = []

  for (const paragraphElement of getDirectChildren(txBody, 'p')) {
    const { paragraph, text } = parseParagraph(paragraphElement, shape, ref, chain)
    paragraphs.push(paragraph)
    flatParts.push(text)
  }

  return { paragraphs, text: flatParts.join('\n').trim(), fontScale }
}
