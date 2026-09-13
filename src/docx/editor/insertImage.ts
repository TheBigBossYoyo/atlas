/**
 * Atlas — image insertion helper (Wave E.4)
 *
 * Pure function that, given a DocxBundle and a target Position, returns a
 * NEW bundle + document + range with:
 *   - a fresh image media part inserted into rawArchive
 *   - a fresh relationship pointing to the image part
 *   - a fresh `Drawing` inline appended at the target paragraph (after the
 *     run that contains the position)
 *   - cursor positioned just after the inserted drawing
 *
 * MVP scope:
 *   - Only PNG / JPEG / GIF inputs.
 *   - Image is inserted as a new run with a single Drawing child at the END
 *     of the target paragraph (the simplest deterministic placement).
 *   - 1 EMU = 1/914400 inch; 1 pt = 1/72 inch ⇒ 1 pt = 12700 EMU.
 *   - History/undo is intentionally out of scope for MVP — image insertion
 *     bypasses the Command pipeline because it mutates bundle-level state
 *     (rawArchive + relationships) that the pipeline does not own.
 */

import type { DocxBundle } from '../index'
import type { Document, Drawing, Paragraph, ParagraphChild, Run, Section } from '../model/document'
import type { Relationship } from '../parser/relationships'
import type { ContentTypesPart } from '../serializer/contentTypesWriter'
import { ensureMediaContentType } from '../serializer/contentTypesWriter'

import type { Position, Range } from './commandTypes'

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

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

export function insertImageIntoBundle(
  bundle: DocxBundle,
  position: Position,
  image: InsertImageInput,
): InsertImageResult {
  const document = bundle.document
  const sectionIndex = findSectionIndex(document, position.paragraphPath)
  if (sectionIndex === -1) {
    throw new Error('Target paragraph not found in document')
  }

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

  const drawingRun: Run = {
    kind: 'run',
    children: [drawing],
  }

  const { document: nextDocument, insertedRunIndex } = appendRunToParagraph(
    document,
    sectionIndex,
    position.paragraphPath,
    drawingRun,
  )

  const cursor: Position = {
    paragraphPath: position.paragraphPath,
    runIndex: insertedRunIndex + 1,
    charOffset: 0,
  }

  return {
    bundle: {
      ...bundle,
      document: nextDocument,
      relationships: nextRelationships,
      rawArchive: nextRawArchive,
      contentTypes: nextContentTypes,
    },
    document: nextDocument,
    range: { anchor: cursor, focus: cursor },
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function findSectionIndex(document: Document, paragraphPath: ReadonlyArray<number>): number {
  const sectionIndex = paragraphPath[0] ?? -1
  if (sectionIndex < 0 || sectionIndex >= document.sections.length) {
    return -1
  }
  return sectionIndex
}

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

function appendRunToParagraph(
  document: Document,
  sectionIndex: number,
  paragraphPath: ReadonlyArray<number>,
  newRun: Run,
): { readonly document: Document; readonly insertedRunIndex: number } {
  const section = document.sections[sectionIndex]
  if (section === undefined) {
    throw new Error('Section not found')
  }

  const blockPath = paragraphPath.slice(1)
  if (blockPath.length === 0) {
    throw new Error('Paragraph path missing block index')
  }

  const updateResult = updateBlocks(section.blocks, blockPath, newRun)
  const nextSection: Section = { ...section, blocks: updateResult.blocks }

  const nextSections = document.sections.map((s, i) => (i === sectionIndex ? nextSection : s))
  const nextDocument: Document = { ...document, sections: nextSections }

  return { document: nextDocument, insertedRunIndex: updateResult.insertedRunIndex }
}

function updateBlocks(
  blocks: ReadonlyArray<Section['blocks'][number]>,
  path: ReadonlyArray<number>,
  newRun: Run,
): {
  readonly blocks: ReadonlyArray<Section['blocks'][number]>
  readonly insertedRunIndex: number
} {
  const head = path[0] ?? -1
  if (head < 0 || head >= blocks.length) {
    throw new Error('Block index out of range')
  }

  const target = blocks[head]
  if (target === undefined) {
    throw new Error('Block missing')
  }

  if (path.length === 1) {
    if (target.kind !== 'paragraph') {
      throw new Error('Target block is not a paragraph')
    }

    const nextChildren: ReadonlyArray<ParagraphChild> = [...target.children, newRun]
    const nextParagraph: Paragraph = { ...target, children: nextChildren }
    const nextBlocks = blocks.map((b, i) => (i === head ? nextParagraph : b))
    return { blocks: nextBlocks, insertedRunIndex: nextChildren.length - 1 }
  }

  // Recurse through tables (table → row → cell → blocks).
  if (target.kind !== 'table') {
    throw new Error('Cannot recurse into non-table block')
  }

  const rowIdx = path[1] ?? -1
  const cellIdx = path[2] ?? -1
  const innerPath = path.slice(3)
  const row = target.rows[rowIdx]
  if (row === undefined || row.kind !== 'table-row') {
    throw new Error('Table row not found')
  }
  const cell = row.cells[cellIdx]
  if (cell === undefined || cell.kind !== 'table-cell') {
    throw new Error('Table cell not found')
  }

  const inner = updateBlocks(cell.blocks, innerPath, newRun)
  const nextCell: TableCellLike = { ...cell, blocks: inner.blocks }
  const nextRow = {
    ...row,
    cells: row.cells.map((c, i) => (i === cellIdx ? nextCell : c)),
  }
  const nextTable = {
    ...target,
    rows: target.rows.map((r, i) => (i === rowIdx ? nextRow : r)),
  }
  const nextBlocks = blocks.map((b, i) => (i === head ? nextTable : b))
  return { blocks: nextBlocks, insertedRunIndex: inner.insertedRunIndex }
}

// Local alias to avoid pulling TableCell type into this isolated helper file.
type TableCellLike = {
  readonly kind: 'table-cell'
  readonly blocks: ReadonlyArray<Section['blocks'][number]>
}
