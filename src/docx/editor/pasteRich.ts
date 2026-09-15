/**
 * Atlas — rich paste integration (DXE-19)
 *
 * `pasteBlocks.ts` parses pasted HTML into an ordered `PasteBlock` tree
 * (paragraphs and tables, with inline formatting/color/hyperlinks/images).
 * This module turns that tree into the editor's own Command/History
 * pipeline, plus the bundle-level resources (relationships, media parts,
 * content-type defaults, numbering) a hyperlink/image/list needs alongside
 * the pure document edit — the same reason `insertImage.ts`/
 * `insertHyperlink.ts`/`insertList.ts` each need their own bundle-aware
 * helper instead of a plain Command.
 *
 * Two different strategies are used for the two shapes of pasted content:
 *
 *   - The top-level body streams in as a sequence of commands, extending
 *     the MVP paste helper's approach (`htmlPaste.ts`'s `buildPasteCommands`)
 *     of tracking a running caret position across inserts. Unlike that
 *     helper (whose positions are all hand-computed arithmetic, safe only
 *     because plain-text/simple-format inserts never change the run count a
 *     position addresses), this module derives each next position by
 *     actually *applying* every command against a throwaway working copy of
 *     the document as it goes, then reading the engine's own returned
 *     `range` — required because inserting an inline image or wrapping a
 *     hyperlink can restructure run addressing in ways only the engine
 *     itself really knows. The one hand-computed exception is what happens
 *     after a pasted *table*: `insert-table`'s own returned cursor
 *     deliberately lands inside the table's first cell (matching the
 *     toolbar's "insert table" UX), so content continuing after a pasted
 *     table instead resumes at `advanceSiblingPosition(<position before the
 *     table>, 2)` — the "after" paragraph `insert-table` always creates
 *     alongside the table itself (see that helper's own doc comment).
 *   - A pasted table's cells are instead built as a fully-formed `Table`
 *     model object up front (there is no caret to track inside a cell —
 *     its content is static, not streamed) and inserted verbatim via
 *     `insert-table`'s existing `table` field (DXE-19 groundwork).
 *
 * Every image/hyperlink needs a relationship id, allocated in one pre-pass
 * across the *whole* paste — both the streamed body and every table cell —
 * so ids never collide within the same paste, mirroring what
 * `insertImage.ts`/`insertHyperlink.ts` each do for a single insertion.
 *
 * Remaining documented gaps (matching `pasteBlocks.ts`'s own doc comment):
 * vertical cell merge, nested tables, and per-level list indent (every
 * pasted list is inserted flat at level 0, one shared `numId` per bullet/
 * number kind across the whole paste) — plus, new to this module, no list
 * formatting for a list pasted *inside a table cell* (its cell content is
 * built as a static tree, not applied through `insert-list`).
 */

import type {
  Document,
  Drawing,
  Hyperlink as HyperlinkNode,
  NumberingDef,
  Paragraph,
  ParagraphChild,
  Run,
  Table,
  TableCell as ModelTableCell,
  TableRow as ModelTableRow,
  TextNode,
} from '../model'
import { twip } from '../model'
import type { DocxBundle } from '../index'
import type { NumberingPart } from '../parser/numbering'
import type { Relationship } from '../parser/relationships'
import { ensureMediaContentType, type ContentTypesPart } from '../serializer/contentTypesWriter'

import { applyCommand } from './commands'
import type { Command, Position, Range } from './commandTypes'
import { advanceSiblingPosition, hasAnyFormat, orderedStart, positionsEqual } from './htmlPaste'
import { decodeImageNaturalSizePt, type ImageMimeType } from './insertImage'
import { createListNumberingEntry, pickListNumId, type ListKind } from './insertList'
import type {
  PasteBlock,
  PasteImageRun,
  PasteParagraph,
  PasteTable,
  PasteTableCell,
  PasteTextRun,
} from './pasteBlocks'
import { nextRelationshipIdNumber } from './relationshipIds'

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/** The bundle-level state rich paste's resource allocator needs to read
 * from — a plain snapshot rather than a whole `DocxBundle` so callers (and
 * tests) don't need to fabricate an entire bundle just to build commands. */
export interface RichPasteBundleContext {
  readonly relationships: ReadonlyArray<Relationship>
  readonly rawArchive: ReadonlyMap<string, Uint8Array>
  readonly contentTypes: ContentTypesPart
  readonly numberingPart: NumberingPart | undefined
  readonly numbering: ReadonlyMap<string, NumberingDef>
}

/** Bundle-level additions the paste's images/hyperlinks/lists need merged
 * into the live bundle alongside the document edit itself — `null` when the
 * paste added no such resource (plain formatted text/tables of text only),
 * so the caller can skip an unnecessary `onBundleChange`. */
export interface RichPasteBundlePatch {
  readonly relationships: ReadonlyArray<Relationship>
  readonly rawArchive: ReadonlyMap<string, Uint8Array>
  readonly contentTypes: ContentTypesPart
  readonly numberingPart: NumberingPart | undefined
  readonly numbering: ReadonlyMap<string, NumberingDef>
}

export interface RichPasteResult {
  readonly commands: ReadonlyArray<Command>
  readonly finalCursor: Position
  readonly bundlePatch: RichPasteBundlePatch | null
}

/** Builds a `RichPasteBundleContext` straight from a live `DocxBundle`. */
export function bundleContextFor(bundle: DocxBundle): RichPasteBundleContext {
  return {
    relationships: bundle.relationships ?? [],
    rawArchive: bundle.rawArchive ?? new Map(),
    contentTypes: bundle.contentTypes ?? { defaults: [], overrides: [] },
    numberingPart: bundle.numberingPart,
    numbering: bundle.document.numbering,
  }
}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const HYPERLINK_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const MIME_EXTENSION: Readonly<Record<ImageMimeType, string>> = {
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/gif': 'gif',
}
const EMU_PER_POINT = 12700
/** Matches `buildEmptyTable`'s own default in `commands.ts` — a pasted
 * table has no natural page width of its own to measure against here, so
 * this module targets the same default content width. */
const TABLE_TOTAL_WIDTH_TWIPS = 9000

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Builds the command sequence (plus bundle patch) for pasting `blocks` at
 * `cursor`, optionally replacing `selection` first. Returns `null` when
 * there is nothing to paste (an empty block list) so the caller can fall
 * back to its own plain-text path, mirroring `buildPasteCommands`'s
 * contract.
 */
export async function buildRichPasteCommands(
  document: Document,
  bundle: RichPasteBundleContext,
  blocks: ReadonlyArray<PasteBlock>,
  cursor: Position,
  selection: Range | null,
): Promise<RichPasteResult | null> {
  if (blocks.length === 0) {
    return null
  }

  const allocator = new PasteResourceAllocator(bundle)
  await allocator.allocateForBlocks(blocks)

  const commands: Command[] = []
  let workingDocument = document
  let pos = cursor

  if (selection !== null && !positionsEqual(selection.anchor, selection.focus)) {
    const deleteCmd: Command = { kind: 'delete-range', range: selection }
    const result = applyCommand(workingDocument, deleteCmd)
    commands.push(deleteCmd)
    workingDocument = result.document
    pos = result.range?.focus ?? orderedStart(selection)
  }

  const listGroups = new Map<ListKind, Array<ReadonlyArray<number>>>()

  for (const [blockIndex, block] of blocks.entries()) {
    if (blockIndex > 0) {
      const breakCmd: Command = { kind: 'insert-paragraph-break', at: pos }
      const result = applyCommand(workingDocument, breakCmd)
      commands.push(breakCmd)
      workingDocument = result.document
      pos = result.range?.focus ?? advanceSiblingPosition(pos, 1)
    }

    if (block.kind === 'table') {
      const positionBeforeTable = pos
      const table = buildTableModel(block, allocator)
      const tableCmd: Command = {
        kind: 'insert-table',
        at: pos,
        rows: table.rows.length,
        cols: table.tblGrid?.length ?? 1,
        table,
      }
      const result = applyCommand(workingDocument, tableCmd)
      commands.push(tableCmd)
      workingDocument = result.document
      // insert-table's own returned cursor lands inside the table's first
      // cell (matching the toolbar's "insert table" UX) — paste content
      // continuing after the table instead resumes in the paragraph the
      // command creates right after it (before/table/after — see the
      // module doc comment).
      pos = advanceSiblingPosition(positionBeforeTable, 2)
      continue
    }

    if (block.list !== undefined) {
      const kind: ListKind = block.list.ordered ? 'number' : 'bullet'
      const paths = listGroups.get(kind) ?? []
      paths.push(pos.paragraphPath)
      listGroups.set(kind, paths)
    }

    for (const run of block.runs) {
      if (run.kind === 'image') {
        const drawing = allocator.buildDrawing(run)
        if (drawing === null) {
          continue
        }
        const inlineCmd: Command = { kind: 'insert-inline', at: pos, child: drawing }
        const result = applyCommand(workingDocument, inlineCmd)
        commands.push(inlineCmd)
        workingDocument = result.document
        pos = result.range?.focus ?? pos
        continue
      }

      if (run.text.length === 0) {
        continue
      }

      const textStart = pos
      const textCmd: Command = { kind: 'insert-text', at: textStart, text: run.text }
      const textResult = applyCommand(workingDocument, textCmd)
      commands.push(textCmd)
      workingDocument = textResult.document
      // Tracks the run's own end position separately from the streaming
      // `pos` cursor: `apply-run-format`/`insert-hyperlink` below each need
      // the *original* (pre-their-own-restructuring) span every time, so
      // reassigning `pos` after one must never feed into the other's `range`
      // — each is computed from `textStart`/`textEnd` directly instead.
      const textEnd = textResult.range?.focus ?? textStart
      pos = textEnd

      if (hasAnyFormat(run.format)) {
        const formatCmd: Command = {
          kind: 'apply-run-format',
          range: { anchor: textStart, focus: textEnd },
          format: run.format,
        }
        const formatResult = applyCommand(workingDocument, formatCmd)
        commands.push(formatCmd)
        workingDocument = formatResult.document
        // `apply-run-format` can split the single run `insert-text` just
        // created into several (e.g. isolating a bolded word from its
        // plain-text neighbors), which shifts every run index after the
        // split — `pos` must follow the command's own returned range
        // (already recomputed against the *new* run layout by
        // `applyRunFormatSingleParagraph`/`applyRunFormatAcrossParagraphs`)
        // rather than the pre-split `textEnd`, or the next run's insert
        // addresses a run index that no longer means what it did.
        pos = formatResult.range?.focus ?? pos
      }

      if (run.href !== undefined) {
        const relationshipId = allocator.hyperlinkRelationshipId(run)
        if (relationshipId !== null) {
          const linkCmd: Command = {
            kind: 'insert-hyperlink',
            range: { anchor: textStart, focus: pos },
            url: run.href,
            relationshipId,
          }
          const linkResult = applyCommand(workingDocument, linkCmd)
          commands.push(linkCmd)
          workingDocument = linkResult.document
          // `insert-hyperlink` wraps the run(s) in that range in a new
          // `w:hyperlink` element — like `apply-run-format` above, this
          // command's own returned range is the one to trust afterward.
          pos = linkResult.range?.focus ?? pos
        }
      }
    }
  }

  // Lists are applied last, over every pasted paragraph's *final* path, once
  // every paragraph the paste creates actually exists in `workingDocument` —
  // `insert-list` (like the toolbar's own list toggle) only ever formats
  // paragraphs that are already there.
  for (const [kind, paragraphPaths] of listGroups) {
    const numId = allocator.numIdForList(kind)
    const listCmd: Command = {
      kind: 'insert-list',
      paragraphPaths: Object.freeze(paragraphPaths.map((path) => Object.freeze([...path]))),
      numId,
      level: 0,
    }
    const result = applyCommand(workingDocument, listCmd)
    commands.push(listCmd)
    workingDocument = result.document
  }

  if (commands.length === 0) {
    return null
  }

  return {
    commands: Object.freeze(commands),
    finalCursor: pos,
    bundlePatch: allocator.buildPatch(),
  }
}

// ---------------------------------------------------------------------------
// Internals — resource allocation (relationships / media / content types /
// numbering) shared between the streamed body and pasted tables' cells.
// ---------------------------------------------------------------------------

class PasteResourceAllocator {
  private readonly relationships: Relationship[]
  private nextRelIdNumber: number
  private readonly rawArchive: Map<string, Uint8Array>
  private contentTypes: ContentTypesPart
  private numberingPart: NumberingPart | undefined
  private readonly numbering: Map<string, NumberingDef>
  private readonly mediaCounters = new Map<string, number>()
  private readonly hyperlinkIds = new Map<PasteTextRun, string>()
  private readonly imageAssignments = new Map<PasteImageRun, { relationshipId: string; cx: number; cy: number }>()
  private changed = false

  constructor(context: RichPasteBundleContext) {
    this.relationships = [...context.relationships]
    this.nextRelIdNumber = nextRelationshipIdNumber(context.relationships)
    this.rawArchive = new Map(context.rawArchive)
    this.contentTypes = context.contentTypes
    this.numberingPart = context.numberingPart
    this.numbering = new Map(context.numbering)
  }

  async allocateForBlocks(blocks: ReadonlyArray<PasteBlock>): Promise<void> {
    for (const block of blocks) {
      if (block.kind === 'table') {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            await this.allocateForParagraphs(cell.paragraphs)
          }
        }
        continue
      }
      await this.allocateForParagraphs([block])
    }
  }

  private async allocateForParagraphs(paragraphs: ReadonlyArray<PasteParagraph>): Promise<void> {
    for (const paragraph of paragraphs) {
      for (const run of paragraph.runs) {
        if (run.kind === 'image') {
          await this.allocateImage(run)
        } else if (run.href !== undefined) {
          this.allocateHyperlink(run)
        }
      }
    }
  }

  private allocateHyperlink(run: PasteTextRun): void {
    if (run.href === undefined) {
      return
    }
    const id = this.mintRelationshipId()
    this.relationships.push({ id, type: HYPERLINK_REL_TYPE, target: run.href, targetMode: 'External' })
    this.hyperlinkIds.set(run, id)
    this.changed = true
  }

  private async allocateImage(run: PasteImageRun): Promise<void> {
    const decoded = decodeImageDataUrl(run.dataUrl)
    if (decoded === null) {
      return
    }

    const { bytes, mime } = decoded
    const extension = MIME_EXTENSION[mime]
    const mediaPath = this.allocateMediaPath(extension)
    const id = this.mintRelationshipId()

    this.relationships.push({ id, type: IMAGE_REL_TYPE, target: mediaPath.replace(/^word\//, '') })
    this.rawArchive.set(mediaPath, bytes)
    this.contentTypes = ensureMediaContentType(this.contentTypes, extension)

    const { widthPt, heightPt } = await decodeImageNaturalSizePt(bytes, mime)
    this.imageAssignments.set(run, {
      relationshipId: id,
      cx: Math.round(widthPt * EMU_PER_POINT),
      cy: Math.round(heightPt * EMU_PER_POINT),
    })
    this.changed = true
  }

  private mintRelationshipId(): string {
    const id = `rId${this.nextRelIdNumber}`
    this.nextRelIdNumber += 1
    return id
  }

  private allocateMediaPath(extension: string): string {
    let counter = this.mediaCounters.get(extension) ?? 1
    let path = `word/media/image${counter}.${extension}`
    while (this.rawArchive.has(path)) {
      counter += 1
      path = `word/media/image${counter}.${extension}`
    }
    this.mediaCounters.set(extension, counter + 1)
    return path
  }

  hyperlinkRelationshipId(run: PasteTextRun): string | null {
    return this.hyperlinkIds.get(run) ?? null
  }

  buildDrawing(run: PasteImageRun): Drawing | null {
    const assignment = this.imageAssignments.get(run)
    if (assignment === undefined) {
      return null
    }
    return {
      kind: 'drawing',
      layout: 'inline',
      name: 'Pasted image',
      relationshipId: assignment.relationshipId,
      extent: { cx: assignment.cx, cy: assignment.cy },
    }
  }

  /** Picks (allocating if needed) the numId a pasted list of `kind` should
   * use, mirroring the toolbar's own `pickListNumId`/`ensureListNumbering`
   * pair — see `insertList.ts`'s doc comment — but against this allocator's
   * own locally-tracked numbering map rather than a whole `DocxBundle`, so a
   * paste that mints one runs no risk of colliding with one it mints for a
   * *different* list kind earlier in the very same paste. */
  numIdForList(kind: ListKind): number {
    const numId = pickListNumId(this.numbering, kind)
    const numIdStr = String(numId)
    if (!this.numbering.has(numIdStr)) {
      const entry = createListNumberingEntry(numId, kind)
      this.numbering.set(numIdStr, entry.numberingDef)

      const previousPart = this.numberingPart ?? { abstractNums: new Map(), nums: new Map() }
      this.numberingPart = {
        abstractNums: new Map(previousPart.abstractNums).set(entry.abstractNum.abstractNumId, entry.abstractNum),
        nums: new Map(previousPart.nums).set(entry.numInstance.numId, entry.numInstance),
      }
      this.changed = true
    }
    return numId
  }

  /** Builds a pasted table cell's static content directly as model nodes —
   * there is no caret to stream through a cell, so images/hyperlinks inside
   * one are resolved against this same allocator's pre-pass assignments
   * rather than through the Command pipeline. */
  buildParagraph(paragraph: PasteParagraph): Paragraph {
    const children: ParagraphChild[] = []

    for (const run of paragraph.runs) {
      if (run.kind === 'image') {
        const drawing = this.buildDrawing(run)
        if (drawing !== null) {
          children.push({ kind: 'run', children: Object.freeze([drawing]) })
        }
        continue
      }

      if (run.text.length === 0) {
        continue
      }

      const textNode: TextNode = { kind: 'text', value: run.text }
      const modelRun: Run = {
        kind: 'run',
        ...(hasAnyFormat(run.format) ? { props: run.format } : {}),
        children: Object.freeze([textNode]),
      }

      const relationshipId = run.href !== undefined ? this.hyperlinkRelationshipId(run) : null
      if (relationshipId !== null) {
        const hyperlink: HyperlinkNode = {
          kind: 'hyperlink',
          relationshipId,
          history: true,
          children: Object.freeze([modelRun]),
        }
        children.push(hyperlink)
        continue
      }

      children.push(modelRun)
    }

    return { kind: 'paragraph', children: Object.freeze(children) }
  }

  buildPatch(): RichPasteBundlePatch | null {
    if (!this.changed) {
      return null
    }
    return {
      relationships: Object.freeze([...this.relationships]),
      rawArchive: this.rawArchive,
      contentTypes: this.contentTypes,
      numberingPart: this.numberingPart,
      numbering: this.numbering,
    }
  }
}

// ---------------------------------------------------------------------------
// Internals — building a pasted table's verbatim `Table` model
// ---------------------------------------------------------------------------

function buildTableModel(pasteTable: PasteTable, allocator: PasteResourceAllocator): Table {
  const cols = pasteTable.rows.reduce((max, row) => {
    const rowCols = row.cells.reduce((sum, cell) => sum + Math.max(1, cell.colSpan), 0)
    return Math.max(max, rowCols)
  }, 1)

  const columnWidth = twip(Math.max(1, Math.floor(TABLE_TOTAL_WIDTH_TWIPS / cols)))
  const tblGrid = Object.freeze(Array.from({ length: cols }, () => columnWidth))

  const rows: ModelTableRow[] = pasteTable.rows.map((row) => ({
    kind: 'table-row',
    cells: Object.freeze(row.cells.map((cell) => buildTableCellModel(cell, allocator))),
  }))

  return { kind: 'table', tblGrid, rows: Object.freeze(rows) }
}

function buildTableCellModel(cell: PasteTableCell, allocator: PasteResourceAllocator): ModelTableCell {
  const blocks =
    cell.paragraphs.length > 0
      ? cell.paragraphs.map((paragraph) => allocator.buildParagraph(paragraph))
      : [{ kind: 'paragraph' as const, children: Object.freeze([]) }]

  return {
    kind: 'table-cell',
    ...(cell.colSpan > 1 ? { props: { gridSpan: cell.colSpan } } : {}),
    blocks: Object.freeze(blocks),
  }
}

// ---------------------------------------------------------------------------
// Internals — data: URL decoding
// ---------------------------------------------------------------------------

const DATA_IMAGE_URL_RE = /^data:image\/(png|jpe?g|gif)(?:;[^,]*)?;base64,(.+)$/i

function decodeImageDataUrl(dataUrl: string): { bytes: Uint8Array; mime: ImageMimeType } | null {
  const match = DATA_IMAGE_URL_RE.exec(dataUrl)
  if (match === null) {
    return null
  }

  const rawType = match[1].toLowerCase()
  const mime: ImageMimeType = rawType === 'jpg' ? 'image/jpeg' : (`image/${rawType}` as ImageMimeType)

  try {
    const binary = atob(match[2])
    const bytes = new Uint8Array(binary.length)
    for (let i = 0; i < binary.length; i += 1) {
      bytes[i] = binary.charCodeAt(i)
    }
    return { bytes, mime }
  } catch {
    return null
  }
}
