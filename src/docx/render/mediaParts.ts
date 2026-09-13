/**
 * Atlas — DOCX media part helpers
 *
 * Pure helpers used to turn `word/_rels/document.xml.rels` image
 * relationships into archive paths and browser-paintable MIME types.
 */

import type { Relationship } from '../parser/relationships'

export const IMAGE_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

// Relationship targets in document.xml.rels are relative to the `word/` part folder.
const DOCUMENT_PART_FOLDER = 'word'

const PAINTABLE_IMAGE_TYPES: ReadonlyMap<string, string> = new Map([
  ['png', 'image/png'],
  ['jpg', 'image/jpeg'],
  ['jpeg', 'image/jpeg'],
  ['jfif', 'image/jpeg'],
  ['gif', 'image/gif'],
  ['bmp', 'image/bmp'],
  ['webp', 'image/webp'],
  ['svg', 'image/svg+xml'],
])

/**
 * Returns the archive path of an embedded image relationship, or `null` for
 * non-image or externally linked targets (which are never read from the zip).
 */
export function resolveImagePartPath(relationship: Relationship): string | null {
  if (relationship.type !== IMAGE_RELATIONSHIP_TYPE || relationship.targetMode === 'External') {
    return null
  }

  const target = relationship.target
  if (target.startsWith('/')) {
    return normalizeSegments(target.slice(1).split('/'))
  }
  if (target.startsWith(`${DOCUMENT_PART_FOLDER}/`)) {
    return normalizeSegments(target.split('/'))
  }

  return normalizeSegments([DOCUMENT_PART_FOLDER, ...target.split('/')])
}

/**
 * MIME type for an image part the browser can decode, or `null` for formats
 * it cannot paint (EMF, WMF, TIFF, unknown), which render as a placeholder.
 */
export function imageMimeType(path: string): string | null {
  const fileName = path.split('/').pop() ?? ''
  const dot = fileName.lastIndexOf('.')
  if (dot < 0) {
    return null
  }

  return PAINTABLE_IMAGE_TYPES.get(fileName.slice(dot + 1).toLowerCase()) ?? null
}

function normalizeSegments(segments: ReadonlyArray<string>): string {
  const resolved = segments.reduce<ReadonlyArray<string>>((acc, segment) => {
    if (segment === '' || segment === '.') {
      return acc
    }
    if (segment === '..') {
      return acc.slice(0, -1)
    }
    return [...acc, segment]
  }, [])

  return resolved.join('/')
}
