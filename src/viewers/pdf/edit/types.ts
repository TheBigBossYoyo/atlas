/**
 * The PDF edit model (PDF-EDIT-1).
 *
 * Everything a user changes in a PDF, other than form-field values, is recorded
 * here as plain data and only turned into PDF objects when the document is
 * written (`writePdf.ts`). Keeping it as data rather than mutating a parsed
 * document is what makes undo/redo a stack of states, lets the viewer draw the
 * pending changes over the untouched pdf.js render, and keeps the original
 * bytes as the single base every save starts from.
 *
 * Coordinates are PDF user space of the SOURCE page (points, origin at the
 * page's lower-left, y up) — the space pdf.js's `PageViewport` converts from
 * and pdf-lib draws in — so a change is pinned to the page content it was made
 * on, whatever the view's zoom, rotation or page order.
 *
 * Form-field values are deliberately NOT in this model: pdf.js already keeps
 * them in `pdfDoc.annotationStorage` and writes them itself, with proper field
 * appearances, through `saveDocument()`.
 */

import type { PageRotation } from '../types'
import type { PieceRef, SpanStyle, StyledSpan, TextAlign } from './content/paragraphs'

/** 0..1 per channel, as PDF colour operators take them. */
export type RgbColor = {
  readonly r: number
  readonly g: number
  readonly b: number
}

/** An axis-aligned rectangle in PDF user space; (x, y) is its lower-left corner. */
export type PdfRect = {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

export type PdfPoint = {
  readonly x: number
  readonly y: number
}

/** The three PDF standard font families: always available, never embedded. */
export type StandardFontFamily = 'helvetica' | 'times' | 'courier'

export type FontStyle = {
  readonly family: StandardFontFamily
  readonly bold: boolean
  readonly italic: boolean
}

/** A free-text box, written as a `/FreeText` annotation. */
export type TextBoxAnnotation = {
  readonly kind: 'text'
  readonly id: string
  /** 1-based source page number. */
  readonly page: number
  /** Top-left corner of the first line's box, in the box's own (rotated) frame. */
  readonly anchor: PdfPoint
  /**
   * Counter-clockwise rotation of the text in user space, in degrees: the
   * page's own `/Rotate` plus any saved page rotation at the time the box was
   * placed, so the text reads upright the way the page is shown.
   */
  readonly rotation: PageRotation
  readonly text: string
  readonly fontSize: number
  readonly color: RgbColor
}

/** A highlight over one or more line boxes, written as a `/Highlight` annotation. */
export type HighlightAnnotation = {
  readonly kind: 'highlight'
  readonly id: string
  readonly page: number
  readonly rects: ReadonlyArray<PdfRect>
  readonly color: RgbColor
}

/** A freehand stroke, written as an `/Ink` annotation. */
export type InkAnnotation = {
  readonly kind: 'ink'
  readonly id: string
  readonly page: number
  readonly points: ReadonlyArray<PdfPoint>
  readonly color: RgbColor
  readonly width: number
}

export type PdfAnnotationDraft = TextBoxAnnotation | HighlightAnnotation | InkAnnotation

/**
 * A real edit of the page's own text (PDF-EDIT-2), the way Acrobat's Edit
 * Text works: the original glyphs are taken OUT of the page's drawing
 * instructions and new text is written in their place — in the document's own
 * font when it has every glyph needed, otherwise in a bundled font that
 * matches its metrics. Nothing of the old text remains in the saved file.
 *
 * Identity is the paragraph's position in the ORIGINAL file (stream, operator
 * indices), so every save re-applies all edits to the bytes the document was
 * opened with and an edit can never land twice.
 */
export type TextEdit = {
  readonly kind: 'text-edit'
  readonly id: string
  /** 1-based source page number. */
  readonly page: number
  /** The edited paragraph's id, or null for new text added to the page. */
  readonly paragraphId: string | null
  /** `page`, or a form XObject's reference: which stream the paragraph lives in. */
  readonly streamKey: string
  /** Text-showing operators to remove (op index, element index in a TJ array). */
  readonly pieces: ReadonlyArray<PieceRef>
  /** New text goes right after this `ET`; -1 appends it to the end of the page. */
  readonly insertAfterOp: number
  /** First baseline's start and the baseline direction (unit vector), user space. */
  readonly origin: PdfPoint
  readonly direction: PdfPoint
  readonly leading: number
  /** Width to wrap at, or null to keep each line as typed. */
  readonly width: number | null
  readonly align: TextAlign
  readonly firstLineIndent: number
  /** Tab stops, as distances along the baseline from `origin` (see TextParagraph.tabStops). */
  readonly tabStops: ReadonlyArray<number>
  readonly styles: ReadonlyArray<SpanStyle>
  readonly spans: ReadonlyArray<StyledSpan>
  /** The original paragraph's area (four corners), for hit-testing and the interim preview. */
  readonly bounds: ReadonlyArray<PdfPoint>
}

/** One page of the output document: which source page, and how much to turn it. */
export type PageSlot = {
  /** 1-based source page number. */
  readonly source: number
  /** Clockwise rotation added on top of the page's own `/Rotate`, saved into the file. */
  readonly rotation: PageRotation
}

export type PdfEditState = {
  /** Output page order; a source page missing from it has been deleted. */
  readonly pages: ReadonlyArray<PageSlot>
  readonly annotations: ReadonlyArray<PdfAnnotationDraft>
  readonly textEdits: ReadonlyArray<TextEdit>
}

export type PdfEditTool = 'select' | 'text' | 'highlight' | 'ink' | 'edit-text'
