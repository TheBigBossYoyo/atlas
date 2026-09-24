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
  getDirectChildren,
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

/** Local names `traverseShapeTree` treats as addressable leaf shapes (mirrors `transforms.ts`'s own `LEAF_TAGS`). */
const ADDRESSABLE_SHAPE_TAGS = new Set(['sp', 'pic', 'graphicFrame', 'cxnSp'])

/**
 * SHELL-3 / USR-16 fallback identity: every `sp`/`pic`/`graphicFrame`/`cxnSp`
 * element under `spTree`, depth-first, entering `p:grpSp` containers — used
 * to address a shape for editing when its own `cNvPr` has no `id` (real
 * PowerPoint always writes one, but hand-built/converted files may not), or
 * when its `id` is shared with another shape (both make plain id-matching
 * ambiguous or impossible). A position is always available even when an id
 * isn't. `pptxEdits.ts`'s `collectShapeElementsInOrder` MUST walk in this
 * same order — the two are the read and write sides of one addressing
 * scheme, and edits land on the wrong shape (or none) if they diverge.
 */
function collectShapeElementsInOrder(container: Element): Element[] {
  const results: Element[] = []
  for (const child of getDirectChildren(container)) {
    if (child.localName === 'grpSp') {
      results.push(...collectShapeElementsInOrder(child))
    } else if (ADDRESSABLE_SHAPE_TAGS.has(child.localName)) {
      results.push(child)
    }
  }
  return results
}

export type SlideSize = { readonly width: number; readonly height: number }

export async function resolveSlidePaths(
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
  const addressable = collectShapeElementsInOrder(spTree)
  // SHELL-3 — id -> occurrence count across the WHOLE slide, so a duplicated
  // id (same bug class as a missing one: plain id-matching can no longer
  // pick out the right shape) also falls back to positional addressing.
  const idCounts = new Map<string, number>()
  for (const cNvPr of getElementsByLocalName(slideDocument, 'cNvPr')) {
    const id = cNvPr.getAttribute('id')
    if (id) idCounts.set(id, (idCounts.get(id) ?? 0) + 1)
  }
  const consumedTextNodes = new Set<Element>()
  const shapes: SlideShape[] = []
  let genericStackIndex = 0

  for (const { element, transform: resolvedTransform, inGroup } of positioned) {
    if (signal.cancelled) {
      break
    }

    const id = `${slidePath}-shape-${shapes.length}`
    // S1 — a shape with no xfrm and no placeholder to inherit from still gets a
    // place in the deck, generically stacked, instead of being dropped.
    const transform = resolvedTransform ?? {
      x: 24,
      y: 24 + genericStackIndex * 28,
      w: Math.max(size.width - 48, 120),
      h: 24,
    }
    if (!resolvedTransform) {
      genericStackIndex += 1
    }

    // USR-16 — address for editing: the shape's own cNvPr id when it's present and unique;
    // SHELL-3 — otherwise its structural position (see `collectShapeElementsInOrder` above).
    // Only top-level shapes move directly.
    const rawId = getFirstByLocalName(element, 'cNvPr')?.getAttribute('id') ?? null
    const hasUniqueId = rawId !== null && rawId !== '' && idCounts.get(rawId) === 1
    const ordinal = addressable.indexOf(element)
    const sourceId = hasUniqueId ? rawId : ordinal >= 0 ? `@${ordinal}` : undefined
    const source = { sourceId, movable: !inGroup }

    try {
      if (element.localName === 'sp' || element.localName === 'cxnSp') {
        const shape = buildShapeElement(element, id, transform, chain, consumedTextNodes)
        if (shape) {
          shapes.push({ ...shape, ...source })
        }
      } else if (element.localName === 'pic') {
        const image = await buildImageShape(zip, element, id, transform, relationships, signal)
        if (image) {
          shapes.push({ ...image, ...source })
        }
      } else if (element.localName === 'graphicFrame') {
        const graphicShape = buildGraphicFrameShape(element, id, transform, chain, consumedTextNodes)
        if (graphicShape) {
          shapes.push({ ...graphicShape, ...source })
        }
      }
    } catch {
      // One malformed shape must not drop the rest of the slide (S9's spirit, applied per-shape too).
      continue
    }
  }

  return [...shapes, ...buildStrayTextShapes(slideDocument, slidePath, size, consumedTextNodes)]
}

/** Parses one slide part (USR-16 re-parses only the slides an edit touched). */
export async function parseOneSlide(
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
      partPath: slidePath,
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
      partPath: slidePath,
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
