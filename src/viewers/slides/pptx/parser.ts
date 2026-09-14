/**
 * S1/S6/S7/S8/S9/S11/S12/S14/S17/S18 — PPTX slide parsing orchestrator.
 *
 * Produces the shared `SlideData[]` model consumed by `SlideDeck`. Each
 * slide's parse is isolated in its own try/catch (S9 — a malformed slide no
 * longer discards every already-parsed slide) and independent slides parse
 * in parallel (S18).
 */

import type { SlideData, SlideShape } from '../../shared/SlideDeck.types'
import type { CancelSignal, ZipArchive } from '../shared/xmlUtils'
import {
  getAttributeValue,
  getElementsByLocalName,
  getFirstByLocalName,
  getRequiredZipEntry,
  parseXml,
} from '../shared/xmlUtils'
import { emuToPx } from '../shared/units'
import { resolveSlideBackground } from './fills'
import { buildGraphicFrameShape } from './graphicFrame'
import { buildImageShape } from './image'
import { resolveLayoutChain, type LayoutChain } from './layout'
import { resolveSlideNotes } from './notes'
import { parseRelationships, relationshipsById } from './relationships'
import { buildShapeElement, extractSlideTitle } from './shape'
import { traverseShapeTree } from './transforms'

const DEFAULT_SLIDE_WIDTH = 1280
const DEFAULT_SLIDE_HEIGHT = 720

type SlideSize = { readonly width: number; readonly height: number }

async function resolveSlidePaths(
  zip: ZipArchive,
  signal: CancelSignal,
): Promise<{ readonly slidePaths: ReadonlyArray<string>; readonly size: SlideSize }> {
  const empty = { slidePaths: [], size: { width: DEFAULT_SLIDE_WIDTH, height: DEFAULT_SLIDE_HEIGHT } }
  const presentationXml = await getRequiredZipEntry(zip, 'ppt/presentation.xml').async('string')
  if (signal.cancelled) {
    return empty
  }

  const presentationDocument = parseXml(presentationXml)
  const relationshipTargets = relationshipsById(await parseRelationships(zip, 'ppt/presentation.xml', signal))

  const slideSize = getFirstByLocalName(presentationDocument, 'sldSz')
  const width = emuToPx(getAttributeValue(slideSize ?? presentationDocument.documentElement, ['cx']))
    ?? DEFAULT_SLIDE_WIDTH
  const height = emuToPx(getAttributeValue(slideSize ?? presentationDocument.documentElement, ['cy']))
    ?? DEFAULT_SLIDE_HEIGHT

  const slidePaths = getElementsByLocalName(presentationDocument, 'sldId')
    .map(slideId => getAttributeValue(slideId, ['r:id', 'id']))
    .flatMap(relationshipId => {
      if (!relationshipId) {
        return []
      }

      const path = relationshipTargets.get(relationshipId)
      return path ? [path] : []
    })

  return { slidePaths, size: { width, height } }
}

/** `p:sld@show="0"` marks a slide hidden (S14) — absent/any other value means visible. */
function isSlideHidden(slideDocument: XMLDocument): boolean {
  return slideDocument.documentElement.getAttribute('show') === '0'
}

/** Any `<a:t>` left over after every recognized shape/table has claimed its own — kept visible rather than silently dropped. */
function buildStrayTextShapes(
  slideDocument: XMLDocument,
  slidePath: string,
  size: SlideSize,
  consumedTextNodes: ReadonlySet<Element>,
): SlideShape[] {
  const strays: SlideShape[] = []

  for (const textNode of getElementsByLocalName(slideDocument, 't')) {
    if (consumedTextNodes.has(textNode)) {
      continue
    }

    const text = (textNode.textContent ?? '').trim()
    if (!text) {
      continue
    }

    const left = 24
    const top = 24 + strays.length * 28
    strays.push({
      kind: 'text',
      id: `${slidePath}-stray-${strays.length}`,
      transform: { x: left, y: top, w: Math.max(size.width - left - 24, 120), h: 24 },
      paragraphs: [{ runs: [{ text }], level: 0 }],
      text,
    })
  }

  return strays
}

async function buildSlideShapes(
  zip: ZipArchive,
  slideDocument: XMLDocument,
  slidePath: string,
  size: SlideSize,
  chain: LayoutChain,
  signal: CancelSignal,
): Promise<SlideShape[]> {
  const spTree = getFirstByLocalName(slideDocument, 'spTree')
  if (!spTree) {
    return []
  }

  const relationships = relationshipsById(await parseRelationships(zip, slidePath, signal))
  const positioned = traverseShapeTree(spTree, chain)
  const consumedTextNodes = new Set<Element>()
  const shapes: SlideShape[] = []

  for (const { element, transform } of positioned) {
    if (signal.cancelled) {
      break
    }

    const id = `${slidePath}-shape-${shapes.length}`

    try {
      if (element.localName === 'sp' || element.localName === 'cxnSp') {
        const shape = buildShapeElement(element, id, transform, chain, consumedTextNodes)
        if (shape) {
          shapes.push(shape)
        }
      } else if (element.localName === 'pic') {
        const image = await buildImageShape(zip, element, id, transform, relationships, signal)
        if (image) {
          shapes.push(image)
        }
      } else if (element.localName === 'graphicFrame') {
        const graphicShape = buildGraphicFrameShape(element, id, transform, chain, consumedTextNodes)
        if (graphicShape) {
          shapes.push(graphicShape)
        }
      }
    } catch {
      // One malformed shape must not drop the rest of the slide (S9's spirit, applied per-shape too).
      continue
    }
  }

  return [...shapes, ...buildStrayTextShapes(slideDocument, slidePath, size, consumedTextNodes)]
}

async function parseOneSlide(
  zip: ZipArchive,
  slidePath: string,
  index: number,
  size: SlideSize,
  signal: CancelSignal,
): Promise<SlideData> {
  try {
    const slideXml = await getRequiredZipEntry(zip, slidePath).async('string')
    if (signal.cancelled) {
      throw new Error('cancelled')
    }

    const slideDocument = parseXml(slideXml)
    const chain = await resolveLayoutChain(zip, slidePath, signal)
    const shapes = await buildSlideShapes(zip, slideDocument, slidePath, size, chain, signal)
    const notes = await resolveSlideNotes(zip, slidePath, signal)

    return {
      id: `slide-${index + 1}`,
      index,
      title: extractSlideTitle(slideDocument),
      hidden: isSlideHidden(slideDocument),
      notes,
      width: size.width,
      height: size.height,
      background: resolveSlideBackground(slideDocument, chain),
      shapes,
    }
  } catch (err: unknown) {
    // S9 — one corrupted slide becomes a placeholder, not a discarded deck.
    return {
      id: `slide-${index + 1}`,
      index,
      width: size.width,
      height: size.height,
      shapes: [],
      error: err instanceof Error ? err.message : 'This slide failed to load.',
    }
  }
}

/** Parses every slide in a PPTX archive into the shared `SlideData[]` model. */
export async function parsePptxSlides(
  zip: ZipArchive,
  signal: CancelSignal,
): Promise<ReadonlyArray<SlideData>> {
  const { slidePaths, size } = await resolveSlidePaths(zip, signal)
  if (signal.cancelled) {
    return []
  }

  // S18 — independent slides parse in parallel.
  const slides = await Promise.all(
    slidePaths.map((slidePath, index) => parseOneSlide(zip, slidePath, index, size, signal)),
  )

  return signal.cancelled ? [] : slides
}
