/**
 * Atlas — image insertion helper (Wave E.4, hardened for D14/DXE-18)
 *
 * Given a DocxBundle and a target Position, returns a NEW bundle + document +
 * range with:
 *   - a fresh image media part inserted into rawArchive
 *   - a fresh relationship pointing to the image part
 *   - a fresh `Drawing` inline inserted at the caret's actual position within
 *     the target paragraph (splitting the run there if the caret sits
 *     mid-run), routed through the editor's Command/History pipeline so the
 *     insertion is undoable
 *   - cursor positioned just after the inserted drawing
 *
 * MVP scope: PNG / JPEG / GIF inputs only.
 * 1 EMU = 1/914400 inch; 1 pt = 1/72 inch ⇒ 1 pt = 12700 EMU.
 */

import type { DocxBundle } from '../index'
import type { Document, Drawing } from '../model/document'
import type { Relationship } from '../parser/relationships'
import type { ContentTypesPart } from '../serializer/contentTypesWriter'
import { ensureMediaContentType } from '../serializer/contentTypesWriter'

import { applyCommand } from './commands'
import type { Command, Position, Range } from './commandTypes'

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

export type ImageMimeType = 'image/png' | 'image/jpeg' | 'image/gif'

export interface InsertImageInput {
  readonly bytes: Uint8Array
  readonly mime: ImageMimeType
  readonly suggestedName: string
  /** Width in points. */
  readonly widthPt: number
  /** Height in points. */
  readonly heightPt: number
  /** Optional alt text. */
  readonly description?: string
}

export interface InsertImageResult {
  readonly bundle: DocxBundle
  readonly document: Document
  readonly range: Range
  /** Pushed onto the editor's undo stack so Ctrl+Z removes the image again. */
  readonly inverse: Command
}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const EMU_PER_POINT = 12700
const IMAGE_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

const MIME_EXTENSION: Record<ImageMimeType, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/gif': 'gif',
}

/** Standard screen-image assumption: 96px per inch, 72pt per inch. */
const POINTS_PER_PIXEL = 72 / 96
/** Default Letter/A4 content width (6.5in) — a decoded image wider than this
 * is scaled down (preserving aspect ratio) so it doesn't overflow the page. */
const MAX_IMAGE_WIDTH_PT = 468
const FALLBACK_SIZE_PT = Object.freeze({ widthPt: 200, heightPt: 150 })

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Decodes an image's natural pixel dimensions and converts them to points at
 * 96dpi, scaling down (preserving aspect ratio) if wider than a page's usual
 * content width. Falls back to a fixed 200x150pt box (the old hardcoded
 * default) when decoding isn't available or fails — e.g. in a test
 * environment without `createImageBitmap`, or a corrupt/unsupported image.
 */
export async function decodeImageNaturalSizePt(
  bytes: Uint8Array,
  mime: ImageMimeType,
): Promise<{ readonly widthPt: number; readonly heightPt: number }> {
  if (typeof createImageBitmap !== 'function') {
    return FALLBACK_SIZE_PT
  }

  try {
    const blob = new Blob([new Uint8Array(bytes)], { type: mime })
    const bitmap = await createImageBitmap(blob)

    try {
      const widthPx = bitmap.width
      const heightPx = bitmap.height
      if (!(widthPx > 0) || !(heightPx > 0)) {
        return FALLBACK_SIZE_PT
      }

      const naturalWidthPt = widthPx * POINTS_PER_PIXEL
      const naturalHeightPt = heightPx * POINTS_PER_PIXEL
      if (naturalWidthPt <= MAX_IMAGE_WIDTH_PT) {
        return { widthPt: naturalWidthPt, heightPt: naturalHeightPt }
      }

      const scale = MAX_IMAGE_WIDTH_PT / naturalWidthPt
      return { widthPt: naturalWidthPt * scale, heightPt: naturalHeightPt * scale }
    } finally {
      bitmap.close?.()
    }
  } catch {
    return FALLBACK_SIZE_PT
  }
}

export function insertImageIntoBundle(
  bundle: DocxBundle,
  position: Position,
  image: InsertImageInput,
): InsertImageResult {
  const document = bundle.document
  const relationships = bundle.relationships ?? []
  const rawArchive = bundle.rawArchive ?? new Map<string, Uint8Array>()

  const relationshipId = allocateRelationshipId(relationships)
  const mediaPath = allocateMediaPath(rawArchive, image)
  const extension = MIME_EXTENSION[image.mime]

  const nextRelationships: ReadonlyArray<Relationship> = [
    ...relationships,
    {
      id: relationshipId,
      type: IMAGE_REL_TYPE,
      target: mediaPath.replace(/^word\//, ''),
    },
  ]

  const nextRawArchive = new Map(rawArchive)
  nextRawArchive.set(mediaPath, image.bytes)

  // DXS-07: an image-free document has no `Default` content-type entry for
  // this extension, so without this the saved package references media Word
  // cannot resolve a MIME type for. Registering it here (once, on the first
  // image of a given extension) keeps every subsequent insert a no-op.
  const contentTypesBase: ContentTypesPart = bundle.contentTypes ?? { defaults: [], overrides: [] }
  const nextContentTypes = ensureMediaContentType(contentTypesBase, extension)

  const drawing: Drawing = {
    kind: 'drawing',
    layout: 'inline',
    relationshipId,
    name: image.suggestedName,
    description: image.description,
    extent: {
      cx: Math.round(image.widthPt * EMU_PER_POINT),
      cy: Math.round(image.heightPt * EMU_PER_POINT),
    },
  }

  // DXE-18 — routed through the same Command/History pipeline as every other
  // edit: this both places the image at the caret's real position (splitting
  // the run there if needed) instead of always appending at the paragraph's
  // end, and gives the caller an exact inverse to push onto the undo stack.
  const result = applyCommand(document, { kind: 'insert-inline', at: position, child: drawing })
  // `insert-inline` always computes a resulting range; the fallback below
  // only exists to satisfy applyCommand's general (range is optional for
  // most command kinds) return type.
  const range = result.range ?? { anchor: position, focus: position }

  return {
    bundle: {
      ...bundle,
      document: result.document,
      relationships: nextRelationships,
      rawArchive: nextRawArchive,
      contentTypes: nextContentTypes,
    },
    document: result.document,
    range,
    inverse: result.inverse,
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function allocateRelationshipId(existing: ReadonlyArray<Relationship>): string {
  let max = 0
  for (const rel of existing) {
    const match = /^rId(\d+)$/.exec(rel.id)
    if (match !== null) {
      const value = Number.parseInt(match[1] ?? '0', 10)
      if (Number.isFinite(value) && value > max) {
        max = value
      }
    }
  }
  return `rId${max + 1}`
}

function allocateMediaPath(
  archive: ReadonlyMap<string, Uint8Array>,
  image: InsertImageInput,
): string {
  const extension = MIME_EXTENSION[image.mime]
  let counter = 1
  while (archive.has(`word/media/image${counter}.${extension}`)) {
    counter += 1
  }
  return `word/media/image${counter}.${extension}`
}
