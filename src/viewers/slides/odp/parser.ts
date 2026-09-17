/**
 * S1(ODP)/S2/S6/S7/S8/S9/S11/S12/S14/S17/S18 — ODP slide parsing orchestrator.
 *
 * Mirrors the PPTX orchestrator: each slide (`draw:page`) is parsed in its
 * own try/catch (S9) and independent slides parse in parallel (S18).
 */

import type { SlideData, SlideFill, SlideShape } from '../../shared/SlideDeck.types'
import type { CancelSignal, ZipArchive } from '../shared/xmlUtils'
import { getElementsByLocalName, getFirstByLocalName, readZipText } from '../shared/xmlUtils'
import { buildOdpShape } from './frame'
import { resolveOdpNotes } from './notes'
import { findInheritedProperties, resolveStyleChain, resolveMasterPage, resolveSlideSize, buildStyleIndex, loadOdpDocuments, type OdpStyleIndex, type SlidePageSize } from './styles'
import { traverseOdpShapes } from './shapes'
import { parseOdpTextBody } from './text'

const TITLE_CLASSES = new Set(['title'])

function resolveBackgroundFromStyleName(styleName: string | null, index: OdpStyleIndex): SlideFill | undefined {
  const chain = resolveStyleChain(styleName, index)
  const properties = findInheritedProperties(chain, 'drawing-page-properties')
  if (!properties) {
    return undefined
  }

  if (properties.getAttribute('draw:fill') === 'none') {
    return { kind: 'none' }
  }

  const color = properties.getAttribute('draw:fill-color')
  return color ? { kind: 'solid', color } : undefined
}

/** Slide own background, falling back to its master page's (S8's background-inheritance acceptance criterion). */
function resolveSlideBackground(page: Element, index: OdpStyleIndex): SlideFill | undefined {
  const ownFill = resolveBackgroundFromStyleName(page.getAttribute('draw:style-name'), index)
  if (ownFill) {
    return ownFill
  }

  const masterPage = resolveMasterPage(page, index)
  return masterPage
    ? resolveBackgroundFromStyleName(masterPage.getAttribute('draw:style-name'), index)
    : undefined
}

/** `presentation:visibility="hidden"` marks a slide hidden (S14). */
function isPageHidden(page: Element): boolean {
  return page.getAttribute('presentation:visibility') === 'hidden'
}

/** A `presentation:class="title"` frame's text, for nav labels (S17). */
function extractPageTitle(page: Element, index: OdpStyleIndex): string | undefined {
  for (const frame of getElementsByLocalName(page, 'frame')) {
    const presentationClass = frame.getAttribute('presentation:class')
    if (!presentationClass || !TITLE_CLASSES.has(presentationClass)) {
      continue
    }

    const textBox = getFirstByLocalName(frame, 'text-box')
    const text = textBox ? parseOdpTextBody(textBox, index).text : ''
    if (text) {
      return text
    }
  }

  return undefined
}

async function buildPageShapes(
  zip: ZipArchive,
  page: Element,
  index: OdpStyleIndex,
  signal: CancelSignal,
): Promise<SlideShape[]> {
  const positioned = traverseOdpShapes(page)
  const shapes: SlideShape[] = []

  for (const [shapeIndex, { element, transform, inGroup }] of positioned.entries()) {
    if (signal.cancelled) {
      break
    }

    try {
      const id = `${page.getAttribute('draw:name') ?? 'page'}-shape-${shapeIndex}`
      const shape = await buildOdpShape(zip, element, id, transform, index, signal)
      if (shape) {
        // USR-16 — a shape is addressed for editing by its position in this
        // same traversal order (see odp/editing/odpEdits.ts).
        shapes.push({ ...shape, sourceId: String(shapeIndex), movable: !inGroup })
      }
    } catch {
      continue
    }
  }

  return shapes
}

async function parseOnePage(
  zip: ZipArchive,
  page: Element,
  pageIndex: number,
  size: SlidePageSize,
  styleIndex: OdpStyleIndex,
  signal: CancelSignal,
): Promise<SlideData> {
  try {
    const shapes = await buildPageShapes(zip, page, styleIndex, signal)

    return {
      id: page.getAttribute('draw:name') ?? `slide-${pageIndex + 1}`,
      index: pageIndex,
      title: extractPageTitle(page, styleIndex),
      hidden: isPageHidden(page),
      notes: resolveOdpNotes(page, styleIndex),
      width: size.width,
      height: size.height,
      background: resolveSlideBackground(page, styleIndex),
      shapes,
    }
  } catch (err: unknown) {
    // S9 — one corrupted page becomes a placeholder, not a discarded deck.
    return {
      id: `slide-${pageIndex + 1}`,
      index: pageIndex,
      width: size.width,
      height: size.height,
      shapes: [],
      error: err instanceof Error ? err.message : 'This slide failed to load.',
    }
  }
}

/** Parses every slide (`draw:page`) in an ODP archive into the shared `SlideData[]` model. */
export async function parseOdpSlides(
  zip: ZipArchive,
  signal: CancelSignal,
): Promise<ReadonlyArray<SlideData>> {
  const docs = await loadOdpDocuments(path => readZipText(zip, path, signal))
  if (signal.cancelled) {
    return []
  }

  const styleIndex = buildStyleIndex(docs)
  const pages = getElementsByLocalName(docs.contentDocument, 'page')
  const size = resolveSlideSize(pages[0] ?? null, styleIndex, docs.contentDocument)

  // S18 — independent slides parse in parallel.
  const slides = await Promise.all(
    pages.map((page, pageIndex) => parseOnePage(zip, page, pageIndex, size, styleIndex, signal)),
  )

  return signal.cancelled ? [] : slides
}
