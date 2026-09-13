/**
 * Shared types for the PDF viewer module (Wave 3-P).
 *
 * Kept dependency-free (no pdfjs-dist imports beyond the type-only
 * `PDFDocumentProxy`/`PDFPageProxy`) so pure helper modules can be unit
 * tested without loading the real pdfjs-dist runtime.
 */

import type { PDFDocumentProxy, PDFPageProxy, PageViewport } from 'pdfjs-dist'

export type PdfDocument = PDFDocumentProxy
export type PdfPage = PDFPageProxy

export type PdfDestination = string | readonly unknown[] | null

export type PdfOutlineNode = {
  readonly title: string
  readonly dest: PdfDestination
  readonly items: ReadonlyArray<PdfOutlineNode>
}

export type PdfRef = {
  readonly num: number
  readonly gen: number
}

/** Zoom can be a literal fraction (100% = 1) or one of the two fit modes. */
export type ZoomMode = number | 'fit-width' | 'fit-page'

/** Page rotation, additive with each page's own intrinsic /Rotate entry. */
export type PageRotation = 0 | 90 | 180 | 270

/** Unscaled (scale 1, rotation 0) page dimensions, in PDF points. */
export type PageGeometry = {
  readonly width: number
  readonly height: number
}

/** A rendered (or in-flight) text layer, matching pdfjs-dist's `TextLayer`
 * public instance shape — narrowed to just what this viewer calls. */
export type TextLayerInstance = {
  render(): Promise<unknown>
  cancel(): void
}

/** pdfjs-dist's `TextContent` type isn't re-exported from the package's
 * top-level entry (only from its internal `display/api` module), so it's
 * derived structurally from the already-exported `PDFPageProxy` instead of
 * reaching into an unexported path. */
export type PdfTextContent = Awaited<ReturnType<PdfPage['getTextContent']>>

/** The slice of the dynamically-imported `pdfjs-dist` module this viewer's
 * subcomponents need, threaded down as a prop instead of each doing its own
 * `import()` — keeps a single loaded instance and makes the components
 * mockable in tests without touching the real pdfjs-dist runtime. */
export type PdfjsRuntime = {
  readonly TextLayer: new (params: {
    textContentSource: PdfTextContent
    container: HTMLElement
    viewport: PageViewport
  }) => TextLayerInstance
}

