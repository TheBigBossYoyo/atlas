/** S3/S6/S8/S11/S19 — builds the `SlideShape` for one positioned ODP drawing element. */

import type { CancelSignal, ZipArchive } from '../shared/xmlUtils'
import { getFirstByLocalName, resolvePartPath, toDataUrl } from '../shared/xmlUtils'
import type { SlideShape, SlideTransform } from '../../shared/SlideDeck.types'
import { resolveOdpAltText, resolveOdpBorder, resolveOdpFill, resolveOdpGeometry } from './shapes'
import type { OdpStyleIndex } from './styles'
import { parseOdpTable } from './table'
import { parseOdpTextBody } from './text'

async function buildImageShape(
  zip: ZipArchive,
  image: Element,
  id: string,
  transform: SlideTransform,
  alt: string,
  signal: CancelSignal,
): Promise<SlideShape | null> {
  const href = image.getAttribute('xlink:href') ?? image.getAttribute('href')
  if (!href) {
    return null
  }

  const src = await toDataUrl(zip, resolvePartPath('content.xml', href), signal)
  return src ? { kind: 'image', id, transform, src, alt } : null
}

function buildTextShape(
  container: Element,
  id: string,
  transform: SlideTransform,
  styleName: string | null,
  index: OdpStyleIndex,
): SlideShape | null {
  const { paragraphs, text } = parseOdpTextBody(container, index)
  if (!text) {
    return null
  }

  return {
    kind: 'text',
    id,
    transform,
    paragraphs,
    text,
    fill: resolveOdpFill(styleName, index),
    border: resolveOdpBorder(styleName, index),
    geometry: resolveOdpGeometry(container),
  }
}

/**
 * Builds the `SlideShape` for one positioned drawing element — a `draw:frame`
 * (dispatching on its text-box/image/table/object content), a bare shape
 * (`draw:rect`/`draw:custom-shape`/...) with its own inline text, or a
 * top-level `table:table`.
 */
export async function buildOdpShape(
  zip: ZipArchive,
  shape: Element,
  id: string,
  transform: SlideTransform,
  index: OdpStyleIndex,
  signal: CancelSignal,
): Promise<SlideShape | null> {
  const styleName = shape.getAttribute('draw:style-name')

  if (shape.localName === 'table') {
    return { kind: 'table', id, transform, rows: parseOdpTable(shape, index) }
  }

  if (shape.localName !== 'frame') {
    const fill = resolveOdpFill(styleName, index)
    const border = resolveOdpBorder(styleName, index)
    const geometry = resolveOdpGeometry(shape)
    const text = buildTextShape(shape, id, transform, styleName, index)
    if (text) {
      return text
    }

    const hasVisibleFill = fill !== undefined && fill.kind !== 'none'
    return hasVisibleFill || border || (geometry && geometry !== 'rect')
      ? { kind: 'shape', id, transform, fill, border, geometry }
      : null
  }

  const table = getFirstByLocalName(shape, 'table')
  if (table) {
    return { kind: 'table', id, transform, rows: parseOdpTable(table, index) }
  }

  const image = getFirstByLocalName(shape, 'image')
  if (image) {
    return buildImageShape(zip, image, id, transform, resolveOdpAltText(shape), signal)
  }

  if (getFirstByLocalName(shape, 'object') || getFirstByLocalName(shape, 'object-ole')) {
    return { kind: 'unsupported', id, transform, label: 'Embedded object not supported' }
  }

  const textBox = getFirstByLocalName(shape, 'text-box')
  if (textBox) {
    return buildTextShape(textBox, id, transform, styleName, index)
  }

  return null
}
