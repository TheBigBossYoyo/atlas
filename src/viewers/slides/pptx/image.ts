/** S10/S19 — builds an image `SlideShape` for `p:pic`: source data URL, crop, and alt text. */

import type { CancelSignal, ZipArchive } from '../shared/xmlUtils'
import { getAttributeValue, getFirstByLocalName, toDataUrl } from '../shared/xmlUtils'
import type { SlideImage, SlideTransform } from '../../shared/SlideDeck.types'
import { resolveCrop } from './transforms'

/** Resolves an image relationship id (`a:blip@r:embed`) to a slide-relative part path. */
function resolveEmbedPath(picture: Element, relationships: ReadonlyMap<string, string>): string | undefined {
  const blip = getFirstByLocalName(picture, 'blip')
  const relationshipId = blip ? getAttributeValue(blip, ['r:embed', 'embed']) : null
  return relationshipId ? relationships.get(relationshipId) : undefined
}

function resolveAltText(picture: Element): string {
  const cNvPr = getFirstByLocalName(picture, 'cNvPr')
  return cNvPr?.getAttribute('descr') ?? ''
}

/** Builds a `SlideImage` for `p:pic`, or `null` when its embedded relationship can't be resolved to bytes. */
export async function buildImageShape(
  zip: ZipArchive,
  picture: Element,
  id: string,
  transform: SlideTransform,
  relationships: ReadonlyMap<string, string>,
  signal: CancelSignal,
): Promise<SlideImage | null> {
  const imagePath = resolveEmbedPath(picture, relationships)
  if (!imagePath) {
    return null
  }

  const src = await toDataUrl(zip, imagePath, signal)
  if (!src) {
    return null
  }

  const blipFill = getFirstByLocalName(picture, 'blipFill')
  const crop = resolveCrop(blipFill)

  return { kind: 'image', id, transform, src, alt: resolveAltText(picture), crop }
}
