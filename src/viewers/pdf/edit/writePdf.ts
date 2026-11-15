/**
 * Writes the PDF edit model into a PDF (PDF-EDIT-1).
 *
 * Takes the bytes to start from (the original file, or pdf.js's
 * `saveDocument()` output when form fields were filled in — see
 * `buildEditedPdf`) and applies, in order:
 *
 *   1. text edits — the page's own text rewritten in its drawing
 *      instructions (see `content/rewriteText.ts`): old glyphs removed, new
 *      text written in the document's font or a metric-compatible fallback;
 *   2. annotations — real `/FreeText`, `/Highlight` and `/Ink` annotation
 *      objects, each with its own appearance stream, so every reader shows
 *      them the same way and annotation-aware readers (Acrobat, Preview, …)
 *      still list them as comments the user can select or delete;
 *   3. page rotation — `/Rotate` set to the page's own rotation plus the saved
 *      turn;
 *   4. page order and deletions — the page tree rebuilt as one flat `/Kids`
 *      array, and form fields whose widgets lived only on deleted pages
 *      removed from the AcroForm so other readers don't show orphans.
 *
 * pdf-lib is passed in rather than imported: the viewer loads it on demand
 * (it is only needed once the user edits something), and tests hand it the
 * real module.
 */

import type * as PdfLibModule from 'pdf-lib'
import type { PDFArray, PDFContext, PDFDict, PDFDocument, PDFFont, PDFHexString, PDFOperator, PDFPage, PDFRef } from 'pdf-lib'

import { combineRotation, normalizeRotation } from '../rotation'
import { applyTextEditsToPage, coveringFallback, finalizeFallbackFonts, type FallbackFontSource } from './content/rewriteText'
import { hasDocumentEdits } from './editState'
import { UnsupportedCharacterError } from './errors'
import {
  TEXT_BOX_PADDING,
  boundingRect,
  layoutTextBox,
  textBoxCorners,
} from './textLayout'
import type {
  HighlightAnnotation,
  InkAnnotation,
  PdfEditState,
  PdfRect,
  RgbColor,
  TextBoxAnnotation,
  TextEdit,
} from './types'

export type PdfLib = typeof PdfLibModule

export { StaleTextEditError, UnsupportedCharacterError } from './errors'
export type { FallbackFontSource } from './content/rewriteText'

const ANNOTATION_FLAG_PRINT = 4
const HIGHLIGHT_STATE = 'GSHighlight'
const TEXT_BOX_FONT_KEY = 'Helv'

type WriteContext = {
  readonly lib: PdfLib
  readonly doc: PDFDocument
  readonly context: PDFContext
  readonly fonts: Map<string, PDFFont>
  readonly now: Date
  /** Bundled fonts for characters the standard fonts lack; null when none were supplied. */
  readonly fallback: FallbackFontSource | null
}

async function getFont(ctx: WriteContext, postScriptName: string): Promise<PDFFont> {
  const cached = ctx.fonts.get(postScriptName)
  if (cached) return cached
  const font = await ctx.doc.embedFont(postScriptName)
  ctx.fonts.set(postScriptName, font)
  return font
}

function encodeOrThrow(font: PDFFont, text: string): ReturnType<PDFFont['encodeText']> {
  try {
    return font.encodeText(text)
  } catch {
    throw new UnsupportedCharacterError(text)
  }
}

function rgbArray(color: RgbColor): number[] {
  return [color.r, color.g, color.b]
}

/** An annotation dictionary with the entries every annotation written here shares. */
function baseAnnotation(
  ctx: WriteContext,
  page: PDFPage,
  subtype: string,
  id: string,
  rect: PdfRect,
  color: RgbColor,
  appearance: PDFRef,
): PDFDict {
  const { lib, context, now } = ctx
  return context.obj({
    Type: 'Annot',
    Subtype: subtype,
    Rect: [rect.x, rect.y, rect.x + rect.width, rect.y + rect.height],
    C: rgbArray(color),
    F: ANNOTATION_FLAG_PRINT,
    P: page.ref,
    NM: lib.PDFHexString.fromText(id),
    M: lib.PDFString.fromDate(now),
    CreationDate: lib.PDFString.fromDate(now),
    AP: { N: appearance },
  })
}

function attachAnnotation(ctx: WriteContext, page: PDFPage, annotation: PDFDict): void {
  const ref = ctx.context.register(annotation)
  page.node.addAnnot(ref)
}

function inflate(rect: PdfRect, by: number): PdfRect {
  return { x: rect.x - by, y: rect.y - by, width: rect.width + by * 2, height: rect.height + by * 2 }
}

// ---- Annotations ----------------------------------------------------------------

function writeHighlight(ctx: WriteContext, page: PDFPage, annotation: HighlightAnnotation): void {
  if (annotation.rects.length === 0) return
  const { lib, context } = ctx
  const bounds = boundingRect(annotation.rects.flatMap((r) => [
    { x: r.x, y: r.y },
    { x: r.x + r.width, y: r.y + r.height },
  ]))
  const operators: PDFOperator[] = [
    lib.setGraphicsState(HIGHLIGHT_STATE),
    lib.setFillingRgbColor(annotation.color.r, annotation.color.g, annotation.color.b),
    ...annotation.rects.map((r) => lib.rectangle(r.x, r.y, r.width, r.height)),
    lib.fill(),
  ]
  // Form space is user space here (identity Matrix, BBox = the bounds), so the
  // appearance maps onto Rect without any scaling.
  const appearance = context.register(
    context.formXObject(operators, {
      BBox: [bounds.x, bounds.y, bounds.x + bounds.width, bounds.y + bounds.height],
      Resources: { ExtGState: { [HIGHLIGHT_STATE]: { Type: 'ExtGState', BM: 'Multiply' } } },
    }),
  )
  const dict = baseAnnotation(ctx, page, 'Highlight', annotation.id, bounds, annotation.color, appearance)
  // UL, UR, LL, LR order — what Acrobat writes and every reader accepts (the
  // spec's own wording says counter-clockwise, which nobody implements).
  dict.set(
    lib.PDFName.of('QuadPoints'),
    context.obj(annotation.rects.flatMap((r) => [
      r.x, r.y + r.height,
      r.x + r.width, r.y + r.height,
      r.x, r.y,
      r.x + r.width, r.y,
    ])),
  )
  attachAnnotation(ctx, page, dict)
}

function writeInk(ctx: WriteContext, page: PDFPage, annotation: InkAnnotation): void {
  const [first, ...rest] = annotation.points
  if (!first) return
  const { lib, context } = ctx
  const bounds = inflate(boundingRect(annotation.points), annotation.width / 2 + 1)
  const operators: PDFOperator[] = [
    lib.setStrokingRgbColor(annotation.color.r, annotation.color.g, annotation.color.b),
    lib.setLineWidth(annotation.width),
    lib.setLineCap(lib.LineCapStyle.Round),
    lib.setLineJoin(lib.LineJoinStyle.Round),
    lib.moveTo(first.x, first.y),
    // A single click still leaves a dot: a zero-length segment with round caps.
    ...(rest.length > 0 ? rest : [first]).map((point) => lib.lineTo(point.x, point.y)),
    lib.stroke(),
  ]
  const appearance = context.register(
    context.formXObject(operators, {
      BBox: [bounds.x, bounds.y, bounds.x + bounds.width, bounds.y + bounds.height],
    }),
  )
  const dict = baseAnnotation(ctx, page, 'Ink', annotation.id, bounds, annotation.color, appearance)
  dict.set(lib.PDFName.of('InkList'), context.obj([annotation.points.flatMap((p) => [p.x, p.y])]))
  dict.set(lib.PDFName.of('BS'), context.obj({ Type: 'Border', W: annotation.width, S: 'S' }))
  attachAnnotation(ctx, page, dict)
}

async function writeTextBox(ctx: WriteContext, page: PDFPage, annotation: TextBoxAnnotation): Promise<void> {
  if (annotation.text.trim().length === 0) return
  const { lib, context } = ctx
  const { fontSize, color } = annotation
  // Helvetica (never embedded) when it can draw every character; otherwise a
  // subset of the bundled sans-serif fallback (see `fallbackFontFile`).
  const helvetica = await getFont(ctx, 'Helvetica')
  const lines = annotation.text.split(/\r\n|\r|\n/)
  const helveticaCan = lines.every((line) => {
    try {
      helvetica.encodeText(line)
      return true
    } catch {
      return false
    }
  })
  let font: { readonly ref: PDFRef; width(text: string): number; encode(text: string): PDFHexString } = {
    ref: helvetica.ref,
    width: (text) => helvetica.widthOfTextAtSize(text, fontSize),
    encode: (text) => encodeOrThrow(helvetica, text),
  }
  if (!helveticaCan) {
    if (!ctx.fallback) throw new UnsupportedCharacterError(annotation.text)
    const fallback = await coveringFallback(lib, ctx.doc, ctx.fallback, { family: 'helvetica', bold: false, italic: false }, annotation.text)
    font = {
      ref: fallback.ref,
      width: (text) => fallback.width(text, fontSize),
      encode: (text) => lib.PDFHexString.of(fallback.encode(text).slice(1, -1)),
    }
  }
  const layout = layoutTextBox(annotation.text, fontSize, (line) => font.width(line))

  const textOperators: PDFOperator[] = []
  layout.lines.forEach((line, index) => {
    if (line.length === 0) return
    textOperators.push(
      lib.beginText(),
      lib.setFontAndSize(TEXT_BOX_FONT_KEY, fontSize),
      lib.setFillingRgbColor(color.r, color.g, color.b),
      lib.moveText(TEXT_BOX_PADDING, layout.height - layout.baselines[index]),
      lib.showText(font.encode(line)),
      lib.endText(),
    )
  })

  const radians = (annotation.rotation * Math.PI) / 180
  const cos = Math.cos(radians)
  const sin = Math.sin(radians)
  // BBox is the unrotated box; Matrix turns it, and a reader then fits the
  // turned box onto Rect — which is exactly its bounding box, so the fit is a
  // pure translation.
  const appearance = context.register(
    context.formXObject(textOperators, {
      BBox: [0, 0, layout.width, layout.height],
      Matrix: [cos, sin, -sin, cos, 0, 0],
      Resources: { Font: { [TEXT_BOX_FONT_KEY]: font.ref } },
    }),
  )
  const rect = boundingRect(textBoxCorners(annotation.anchor, annotation.rotation, layout.width, layout.height))
  const dict = baseAnnotation(ctx, page, 'FreeText', annotation.id, rect, color, appearance)
  dict.set(lib.PDFName.of('Contents'), lib.PDFHexString.fromText(annotation.text))
  dict.set(
    lib.PDFName.of('DA'),
    lib.PDFString.of(`/${TEXT_BOX_FONT_KEY} ${fontSize} Tf ${color.r} ${color.g} ${color.b} rg`),
  )
  dict.set(lib.PDFName.of('BS'), context.obj({ Type: 'Border', W: 0 }))
  attachAnnotation(ctx, page, dict)
}

// ---- Page tree --------------------------------------------------------------------------

const INHERITABLE_PAGE_KEYS = ['Resources', 'MediaBox', 'CropBox', 'Rotate'] as const

/**
 * Replaces the page tree with one flat `/Kids` array in the new order.
 *
 * pdf-lib's `removePage`/`insertPage` keep the tree's shape, which leaves
 * empty intermediate nodes behind when pages move between branches; a flat
 * array is valid at any size that matters here and has no such corner cases.
 * Each page first gets its own copy of every attribute it used to inherit
 * from a parent node it is about to be detached from.
 */
function rebuildPageTree(ctx: WriteContext, orderedPages: ReadonlyArray<PDFPage>): void {
  const { lib, doc, context } = ctx
  const pagesRef = doc.catalog.get(lib.PDFName.of('Pages'))
  const pagesDict = doc.catalog.Pages()
  if (!(pagesRef instanceof lib.PDFRef)) {
    throw new Error('This PDF has no indirect page tree root.')
  }
  for (const page of orderedPages) {
    for (const key of INHERITABLE_PAGE_KEYS) {
      const name = lib.PDFName.of(key)
      if (page.node.get(name) === undefined) {
        const inherited = page.node.getInheritableAttribute(name)
        if (inherited !== undefined) page.node.set(name, inherited)
      }
    }
    page.node.set(lib.PDFName.of('Parent'), pagesRef)
  }
  pagesDict.set(lib.PDFName.of('Kids'), context.obj(orderedPages.map((page) => page.ref)))
  pagesDict.set(lib.PDFName.of('Count'), context.obj(orderedPages.length))
}

function collectAnnotationRefs(ctx: WriteContext, page: PDFPage): Set<string> {
  const refs = new Set<string>()
  const annots = page.node.lookupMaybe(ctx.lib.PDFName.of('Annots'), ctx.lib.PDFArray)
  annots?.asArray().forEach((entry) => {
    if (entry instanceof ctx.lib.PDFRef) refs.add(entry.toString())
  })
  return refs
}

/**
 * Drops form fields whose widgets were all on deleted pages. Without it the
 * AcroForm still lists them, and readers that enumerate fields (Acrobat's form
 * panel, printing "form fields only", JavaScript) find widgets on no page.
 * Edits `fields` in place; a parent left with no kids goes too.
 */
function pruneFieldTree(ctx: WriteContext, fields: PDFArray, deleted: ReadonlySet<string>): void {
  const { lib } = ctx
  for (let index = fields.size() - 1; index >= 0; index -= 1) {
    const entry = fields.get(index)
    if (entry instanceof lib.PDFRef && deleted.has(entry.toString())) {
      fields.remove(index)
      continue
    }
    const field = entry instanceof lib.PDFRef ? ctx.context.lookupMaybe(entry, lib.PDFDict) : undefined
    const kids = field?.lookupMaybe(lib.PDFName.of('Kids'), lib.PDFArray)
    if (field && kids) {
      const before = kids.size()
      pruneFieldTree(ctx, kids, deleted)
      if (before > 0 && kids.size() === 0) fields.remove(index)
    }
  }
}

function pruneDeletedWidgets(ctx: WriteContext, deletedPages: ReadonlyArray<PDFPage>): void {
  const { lib, doc } = ctx
  const deleted = new Set<string>()
  deletedPages.forEach((page) => collectAnnotationRefs(ctx, page).forEach((ref) => deleted.add(ref)))
  if (deleted.size === 0) return
  const acroForm = doc.catalog.lookupMaybe(lib.PDFName.of('AcroForm'), lib.PDFDict)
  const fields = acroForm?.lookupMaybe(lib.PDFName.of('Fields'), lib.PDFArray)
  if (fields) pruneFieldTree(ctx, fields, deleted)
}

/**
 * Removes deleted pages from the FILE, not just from the page tree.
 *
 * pdf-lib writes every object it holds, reachable or not, so a page that was
 * only unlinked from `/Kids` would still be in the saved bytes — its content
 * stream, images and fonts recoverable by anyone who opens the file in a text
 * editor. For a user deleting a page precisely because it holds something
 * private, that is a leak. So: every remaining reference to a deleted page
 * (a bookmark, a named destination, a link on another page) is replaced by
 * null — readers treat such a destination as going nowhere — and then every
 * object no longer reachable from the trailer is dropped.
 */
function dropUnreachableObjects(ctx: WriteContext, deletedPages: ReadonlyArray<PDFPage>): void {
  const { lib, context } = ctx
  const deleted = new Set(deletedPages.map((page) => page.ref.toString()))

  const scrub = (value: unknown): void => {
    if (value instanceof lib.PDFArray) {
      for (let index = 0; index < value.size(); index += 1) {
        const entry = value.get(index)
        if (entry instanceof lib.PDFRef && deleted.has(entry.toString())) value.set(index, lib.PDFNull)
        else scrub(entry)
      }
    } else if (value instanceof lib.PDFDict) {
      for (const [key, entry] of value.entries()) {
        if (entry instanceof lib.PDFRef && deleted.has(entry.toString())) value.set(key, lib.PDFNull)
        else scrub(entry)
      }
    } else if (value instanceof lib.PDFStream) {
      scrub(value.dict)
    }
  }
  for (const [ref, object] of context.enumerateIndirectObjects()) {
    if (!deleted.has(ref.toString())) scrub(object)
  }

  const reachable = new Set<string>()
  const pending: unknown[] = Object.values(context.trailerInfo)
  while (pending.length > 0) {
    const value = pending.pop()
    if (value instanceof lib.PDFRef) {
      const key = value.toString()
      if (reachable.has(key)) continue
      reachable.add(key)
      pending.push(context.lookup(value))
    } else if (value instanceof lib.PDFArray) {
      pending.push(...value.asArray())
    } else if (value instanceof lib.PDFDict) {
      pending.push(...value.values())
    } else if (value instanceof lib.PDFStream) {
      pending.push(value.dict)
    }
  }
  for (const [ref] of context.enumerateIndirectObjects()) {
    if (!reachable.has(ref.toString())) context.delete(ref)
  }
}

// ---- Entry points ---------------------------------------------------------------------------

/** Applies `edits` to `baseBytes`. Returns `baseBytes` untouched when there is nothing to apply. */
export async function applyEditsToPdf(
  lib: PdfLib,
  baseBytes: Uint8Array,
  edits: PdfEditState,
  sourcePageCount: number,
  fallback: FallbackFontSource | null = null,
): Promise<Uint8Array> {
  if (!hasDocumentEdits(edits, sourcePageCount)) return baseBytes

  const doc = await lib.PDFDocument.load(baseBytes, { updateMetadata: false })
  const pages = doc.getPages()
  if (pages.length !== sourcePageCount) {
    throw new Error(`Expected ${sourcePageCount} pages but the file has ${pages.length}.`)
  }
  const ctx: WriteContext = { lib, doc, context: doc.context, fonts: new Map(), now: new Date(), fallback }
  const pageFor = (source: number): PDFPage => pages[source - 1]

  // Text edits first: they are located by the original page structure, which
  // every later step (annotations, rotation, page order) leaves untouched.
  const textEditsByPage = new Map<number, TextEdit[]>()
  for (const edit of edits.textEdits) textEditsByPage.set(edit.page, [...(textEditsByPage.get(edit.page) ?? []), edit])
  for (const [source, pageEdits] of textEditsByPage) {
    if (!fallback) throw new Error('Text edits need the bundled fallback fonts.')
    await applyTextEditsToPage(lib, doc, pageFor(source), pageEdits, fallback)
  }

  for (const annotation of edits.annotations) {
    const page = pageFor(annotation.page)
    switch (annotation.kind) {
      case 'text':
        await writeTextBox(ctx, page, annotation)
        break
      case 'highlight':
        writeHighlight(ctx, page, annotation)
        break
      case 'ink':
        writeInk(ctx, page, annotation)
        break
    }
  }

  // Every piece of text is encoded by now: write the fallback font subsets.
  await finalizeFallbackFonts(doc)

  for (const slot of edits.pages) {
    if (slot.rotation === 0) continue
    const page = pageFor(slot.source)
    page.setRotation(lib.degrees(combineRotation(normalizeRotation(page.getRotation().angle), slot.rotation)))
  }

  const deleted: PDFPage[] = []
  const isIdentityOrder =
    edits.pages.length === sourcePageCount && edits.pages.every((slot, index) => slot.source === index + 1)
  if (!isIdentityOrder) {
    const kept = new Set(edits.pages.map((slot) => slot.source))
    const deletedPages = pages.filter((_, index) => !kept.has(index + 1))
    pruneDeletedWidgets(ctx, deletedPages)
    rebuildPageTree(ctx, edits.pages.map((slot) => pageFor(slot.source)))
    deleted.push(...deletedPages)
  }

  // A deleted page, or a content stream a text edit replaced, must not stay
  // in the file as an orphaned object. Fonts are embedded lazily; flushing
  // first makes them real objects the reachability walk can see.
  if (deleted.length > 0 || edits.textEdits.length > 0) {
    await doc.flush()
    dropUnreachableObjects(ctx, deleted)
  }

  return doc.save()
}

export type BuildEditedPdfInput = {
  readonly original: Uint8Array
  /**
   * pdf.js's `saveDocument()` for the loaded document, when form fields were
   * changed — it writes field values with proper appearances as an
   * incremental update, which pdf-lib then builds on. `null` when no field
   * was touched, so an edit-free save never round-trips through pdf.js.
   */
  readonly saveForms: (() => Promise<Uint8Array>) | null
  readonly edits: PdfEditState
  readonly sourcePageCount: number
  /** Loads pdf-lib; only called when there are edits beyond form values. */
  readonly loadPdfLib: () => Promise<PdfLib>
  /** Bundled fonts for text edits and non-Latin text; loaded only when needed. */
  readonly loadFallbackFonts?: () => Promise<FallbackFontSource>
}

/** The bytes a save writes: forms first (pdf.js), then everything else (pdf-lib). */
export async function buildEditedPdf(input: BuildEditedPdfInput): Promise<Uint8Array> {
  const base = input.saveForms ? await input.saveForms() : input.original
  if (!hasDocumentEdits(input.edits, input.sourcePageCount)) return base
  const lib = await input.loadPdfLib()
  const needsFallback = input.edits.textEdits.length > 0 || input.edits.annotations.some((annotation) => annotation.kind === 'text')
  const fallback = needsFallback && input.loadFallbackFonts ? await input.loadFallbackFonts() : null
  return applyEditsToPdf(lib, base, input.edits, input.sourcePageCount, fallback)
}
