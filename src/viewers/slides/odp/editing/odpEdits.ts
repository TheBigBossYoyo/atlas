/**
 * USR-16 (ODP) — editing an OpenDocument presentation.
 *
 * ODF keeps every slide in one part, `content.xml`, so these edits rewrite
 * that single part and leave the rest of the package (styles, master pages,
 * pictures, settings) byte-for-byte untouched — the same passthrough approach
 * the PPTX editor uses for its own parts.
 *
 * Shapes are addressed by their position in the page's shape order, which is
 * exactly the order the parser walks (`traverseOdpShapes`), so the ids the
 * renderer hands back always point at the same element the user clicked.
 */
import {
  childElements,
  descendantElements,
  firstChildElement,
  parseXmlPart,
  serializeXmlPart,
} from '../../../../office/ooxmlDom'
import { readPart, withParts, type OfficePackage } from '../../../../office/officePackage'
import { CENTIMETERS_TO_PIXELS } from '../../shared/units'
import { traverseOdpShapes } from '../shapes'

export const CONTENT_PART = 'content.xml'

export type ShapeBox = { readonly x: number; readonly y: number; readonly w: number; readonly h: number }

const NS = {
  draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
  presentation: 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0',
} as const

/** ODF lengths carry a unit; centimetres are what every producer writes for slide geometry. */
function toCm(px: number): string {
  return `${(px / CENTIMETERS_TO_PIXELS).toFixed(3)}cm`
}

function pagesOf(doc: XMLDocument): Element[] {
  return descendantElements(doc, 'page').filter((page) => page.namespaceURI === NS.draw)
}

function shapeAt(page: Element, sourceId: string): Element | null {
  const index = Number(sourceId)
  if (!Number.isInteger(index) || index < 0) return null
  return traverseOdpShapes(page)[index]?.element ?? null
}

/** Applies `mutate` to the content part; returns the same package when the target is missing. */
function editContent(pkg: OfficePackage, mutate: (doc: XMLDocument) => boolean): OfficePackage {
  const xml = readPart(pkg, CONTENT_PART)
  if (xml === null) return pkg
  const doc = parseXmlPart(xml)
  if (!mutate(doc)) return pkg
  return withParts(pkg, { [CONTENT_PART]: serializeXmlPart(doc) })
}

function editShape(
  pkg: OfficePackage,
  pageIndex: number,
  sourceId: string,
  mutate: (shape: Element, doc: XMLDocument) => void,
): OfficePackage {
  return editContent(pkg, (doc) => {
    const page = pagesOf(doc)[pageIndex]
    const shape = page ? shapeAt(page, sourceId) : null
    if (!shape) return false
    mutate(shape, doc)
    return true
  })
}

/**
 * Replaces a shape's paragraphs with `text` (one `text:p` per line), keeping
 * the first paragraph's style so the new text still looks like the old text.
 */
function writeParagraphs(doc: XMLDocument, container: Element, text: string): void {
  const existing = childElements(container, 'p')
  const styleName = existing[0]?.getAttributeNS(NS.text, 'style-name') ?? existing[0]?.getAttribute('text:style-name')
  for (const paragraph of existing) container.removeChild(paragraph)

  for (const line of text.replace(/\r\n?/g, '\n').split('\n')) {
    const paragraph = doc.createElementNS(NS.text, 'text:p')
    if (styleName) paragraph.setAttribute('text:style-name', styleName)
    if (line !== '') paragraph.textContent = line
    container.appendChild(paragraph)
  }
}

/** A shape's text lives in its `draw:text-box` when it has one, and directly on the shape otherwise. */
function textContainer(doc: XMLDocument, shape: Element): Element {
  const textBox = firstChildElement(shape, 'text-box')
  if (textBox) return textBox
  if (shape.localName === 'frame') {
    const created = doc.createElementNS(NS.draw, 'draw:text-box')
    shape.appendChild(created)
    return created
  }
  return shape
}

export function setShapeText(pkg: OfficePackage, pageIndex: number, sourceId: string, text: string): OfficePackage {
  return editShape(pkg, pageIndex, sourceId, (shape, doc) => {
    writeParagraphs(doc, textContainer(doc, shape), text)
  })
}

export function setShapeBox(pkg: OfficePackage, pageIndex: number, sourceId: string, box: ShapeBox): OfficePackage {
  return editShape(pkg, pageIndex, sourceId, (shape) => {
    shape.setAttribute('svg:x', toCm(box.x))
    shape.setAttribute('svg:y', toCm(box.y))
    shape.setAttribute('svg:width', toCm(Math.max(box.w, 1)))
    shape.setAttribute('svg:height', toCm(Math.max(box.h, 1)))
  })
}

export function deleteShape(pkg: OfficePackage, pageIndex: number, sourceId: string): OfficePackage {
  return editShape(pkg, pageIndex, sourceId, (shape) => {
    shape.parentNode?.removeChild(shape)
  })
}

/** Adds a plain text frame; its id is its position in the page's shape order. */
export function insertTextBox(
  pkg: OfficePackage,
  pageIndex: number,
  box: ShapeBox,
  text: string,
): { readonly pkg: OfficePackage; readonly sourceId: string | null } {
  let sourceId: string | null = null
  const next = editContent(pkg, (doc) => {
    const page = pagesOf(doc)[pageIndex]
    if (!page) return false

    const frame = doc.createElementNS(NS.draw, 'draw:frame')
    frame.setAttribute('draw:layer', 'layout')
    frame.setAttribute('svg:x', toCm(box.x))
    frame.setAttribute('svg:y', toCm(box.y))
    frame.setAttribute('svg:width', toCm(box.w))
    frame.setAttribute('svg:height', toCm(box.h))
    const textBox = doc.createElementNS(NS.draw, 'draw:text-box')
    frame.appendChild(textBox)
    writeParagraphs(doc, textBox, text)
    // Before any presentation:notes, which must stay the page's last child.
    page.insertBefore(frame, firstChildElement(page, 'notes'))

    sourceId = String(traverseOdpShapes(page).findIndex((shape) => shape.element === frame))
    return true
  })
  return { pkg: next, sourceId: sourceId === '-1' ? null : sourceId }
}

// ---------------------------------------------------------------------------
// Speaker notes
// ---------------------------------------------------------------------------

export function setSlideNotes(pkg: OfficePackage, pageIndex: number, text: string): OfficePackage {
  return editContent(pkg, (doc) => {
    const page = pagesOf(doc)[pageIndex]
    if (!page) return false

    let notes = firstChildElement(page, 'notes')
    if (!notes) {
      if (text.trim() === '') return false
      notes = doc.createElementNS(NS.presentation, 'presentation:notes')
      page.appendChild(notes)
    }

    let frame = firstChildElement(notes, 'frame')
    if (!frame) {
      frame = doc.createElementNS(NS.draw, 'draw:frame')
      frame.setAttribute('presentation:class', 'notes')
      frame.setAttribute('svg:x', '1cm')
      frame.setAttribute('svg:y', '13cm')
      frame.setAttribute('svg:width', '16cm')
      frame.setAttribute('svg:height', '10cm')
      notes.appendChild(frame)
    }

    writeParagraphs(doc, textContainer(doc, frame), text)
    return true
  })
}

// ---------------------------------------------------------------------------
// Slides
// ---------------------------------------------------------------------------

function uniquePageName(doc: XMLDocument): string {
  const used = new Set(pagesOf(doc).map((page) => page.getAttribute('draw:name')))
  let index = used.size + 1
  while (used.has(`page${index}`)) index += 1
  return `page${index}`
}

/** Inserts an empty slide after `afterIndex` (-1 for the start), on the reference slide's master page. */
export function addSlide(pkg: OfficePackage, afterIndex: number): { readonly pkg: OfficePackage; readonly index: number } {
  let index = afterIndex
  const next = editContent(pkg, (doc) => {
    const pages = pagesOf(doc)
    const reference = pages[Math.min(Math.max(afterIndex, 0), pages.length - 1)]
    if (!reference?.parentNode) return false

    const page = doc.createElementNS(NS.draw, 'draw:page')
    page.setAttribute('draw:name', uniquePageName(doc))
    for (const attribute of ['draw:master-page-name', 'presentation:presentation-page-layout-name', 'draw:style-name']) {
      const value = reference.getAttribute(attribute)
      if (value !== null) page.setAttribute(attribute, value)
    }
    reference.parentNode.insertBefore(page, afterIndex < 0 ? reference : reference.nextSibling)
    index = afterIndex < 0 ? 0 : afterIndex + 1
    return true
  })
  return { pkg: next, index }
}

export function duplicateSlide(pkg: OfficePackage, slideIndex: number): OfficePackage {
  return editContent(pkg, (doc) => {
    const page = pagesOf(doc)[slideIndex]
    if (!page?.parentNode) return false
    const copy = page.cloneNode(true) as Element
    copy.setAttribute('draw:name', uniquePageName(doc))
    page.parentNode.insertBefore(copy, page.nextSibling)
    return true
  })
}

export function deleteSlide(pkg: OfficePackage, slideIndex: number): OfficePackage {
  return editContent(pkg, (doc) => {
    const pages = pagesOf(doc)
    // A presentation must keep at least one slide.
    if (pages.length <= 1) return false
    const page = pages[slideIndex]
    if (!page) return false
    page.parentNode?.removeChild(page)
    return true
  })
}

export function moveSlide(pkg: OfficePackage, from: number, to: number): OfficePackage {
  return editContent(pkg, (doc) => {
    const pages = pagesOf(doc)
    const moved = pages[from]
    const target = pages[to]
    if (from === to || !moved || !target || !moved.parentNode) return false
    moved.parentNode.insertBefore(moved, from < to ? target.nextSibling : target)
    return true
  })
}
