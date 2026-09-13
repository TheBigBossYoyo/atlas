/**
 * Pure annotation-classification/geometry helpers (PDF-05/P5, PDF-11/P10).
 *
 * pdf.js ships its own `AnnotationLayer` class, but it's built for — and
 * requires — the full `web/pdf_viewer.mjs` viewer framework (a
 * `PDFLinkService`, a `StructTreeLayerBuilder`, a `CommentManager`, an
 * `AnnotationEditorUIManager`, …): private, undocumented collaborators this
 * viewer doesn't otherwise use. Instantiating that class directly, outside
 * its intended host, would be fragile across pdfjs-dist versions. Instead
 * this viewer renders a small, purpose-built overlay directly from
 * `page.getAnnotations()`'s plain data, positioned with the same `viewport`
 * used for the page's canvas so it lines up pixel-for-pixel. This module is
 * the pure part of that: classifying an annotation dict and converting its
 * PDF-space rect into a viewport-space (CSS px) box.
 */

export type ViewportPointConverter = {
  // Real pdf.js `PageViewport.convertToViewportPoint` is typed as returning
  // `any[]` (it always returns exactly [x, y] at runtime) — kept as a plain
  // readonly array here rather than a 2-tuple so the real type is
  // structurally assignable to this one.
  convertToViewportPoint: (x: number, y: number) => ReadonlyArray<number>
}

export type ViewportBox = {
  readonly left: number
  readonly top: number
  readonly width: number
  readonly height: number
}

/**
 * Converts a PDF-space annotation rect (`[x1, y1, x2, y2]`, corners in any
 * order) into a viewport-space box. `PageViewport` only exposes point
 * conversion (not rect conversion), so this converts both corners and takes
 * the bounding box — rotation can swap which corner ends up top-left.
 */
export function annotationRectToViewportBox(
  viewport: ViewportPointConverter,
  rect: readonly [number, number, number, number],
): ViewportBox {
  const [x1, y1, x2, y2] = rect
  const [vx1, vy1] = viewport.convertToViewportPoint(x1, y1)
  const [vx2, vy2] = viewport.convertToViewportPoint(x2, y2)

  return {
    left: Math.min(vx1, vx2),
    top: Math.min(vy1, vy2),
    width: Math.abs(vx2 - vx1),
    height: Math.abs(vy2 - vy1),
  }
}

export type PdfAnnotationKind =
  | 'internal-link'
  | 'external-link'
  | 'text-field'
  | 'checkbox'
  | 'radio'
  | 'ignored'

/** The subset of a pdf.js annotation data dict this module reads. Loosely
 * typed (`unknown`-ish) on purpose — pdf.js's own typings for this shape are
 * `any`, and the real objects carry many more fields we don't use. */
export type RawPdfAnnotation = {
  readonly id: string
  readonly subtype?: string
  readonly rect?: ReadonlyArray<number>
  /** Sanitized, validated absolute URL — pdf.js only sets this when the raw
   * action URL passed its own scheme/format validation. Deliberately the
   * ONLY url-ish field this module trusts; `unsafeUrl` (the raw, unvalidated
   * string pdf.js also exposes) is never rendered as a clickable link. */
  readonly url?: string | null
  readonly dest?: string | ReadonlyArray<unknown> | null
  readonly fieldType?: string | null
  readonly fieldName?: string | null
  readonly fieldValue?: unknown
  readonly checkBox?: boolean
  readonly radioButton?: boolean
  readonly exportValue?: string | null
  readonly buttonValue?: string | null
  readonly noView?: boolean
  readonly hidden?: boolean
}

/**
 * Classifies a raw pdf.js annotation dict into the kind of overlay element
 * this viewer knows how to render. Annotations flagged NoView/Hidden, or of
 * a subtype/action this viewer doesn't support, are 'ignored'.
 */
export function classifyAnnotation(annotation: RawPdfAnnotation): PdfAnnotationKind {
  if (annotation.noView || annotation.hidden) {
    return 'ignored'
  }

  if (annotation.subtype === 'Link') {
    if (typeof annotation.url === 'string' && annotation.url.length > 0) {
      return 'external-link'
    }
    if (annotation.dest !== null && annotation.dest !== undefined) {
      return 'internal-link'
    }
    return 'ignored'
  }

  if (annotation.subtype === 'Widget') {
    if (annotation.fieldType === 'Tx') {
      return 'text-field'
    }
    if (annotation.fieldType === 'Btn') {
      return annotation.radioButton ? 'radio' : 'checkbox'
    }
  }

  return 'ignored'
}

/** Whether a checkbox/radio widget's current value marks it as checked. */
export function isWidgetChecked(annotation: RawPdfAnnotation): boolean {
  if (annotation.checkBox) {
    return annotation.fieldValue === annotation.exportValue
  }
  if (annotation.radioButton) {
    return annotation.fieldValue === annotation.buttonValue
  }
  return false
}

export function isValidAnnotationRect(
  rect: ReadonlyArray<number> | undefined,
): rect is readonly [number, number, number, number] {
  return Array.isArray(rect) && rect.length === 4 && rect.every((value) => Number.isFinite(value))
}
