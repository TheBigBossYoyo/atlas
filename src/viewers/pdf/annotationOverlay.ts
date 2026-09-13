/**
 * DOM-building side of the annotation overlay (PDF-05/P5, PDF-11/P10).
 *
 * The pure classification/geometry math lives in `annotations.ts` (and is
 * unit tested there); this module is the thin, intentionally simple glue
 * that turns one classified annotation into a real positioned DOM element —
 * see `annotations.ts`'s file comment for why this hand-rolls the overlay
 * instead of instantiating pdf.js's own `AnnotationLayer` class.
 */

import type { PageViewport } from 'pdfjs-dist'

import { resolveDestinationPage } from './outline'
import {
  annotationRectToViewportBox,
  classifyAnnotation,
  isValidAnnotationRect,
  isWidgetChecked,
  type RawPdfAnnotation,
} from './annotations'
import type { PdfDocument, PdfRef } from './types'

export type AnnotationOverlayCallbacks = {
  readonly onInternalNavigate: (pageNumber: number) => void
}

function applyBoxStyle(el: HTMLElement, box: { left: number; top: number; width: number; height: number }): void {
  el.style.position = 'absolute'
  el.style.left = `${box.left}px`
  el.style.top = `${box.top}px`
  el.style.width = `${box.width}px`
  el.style.height = `${box.height}px`
}

function createExternalLinkElement(annotation: RawPdfAnnotation): HTMLAnchorElement {
  const anchor = document.createElement('a')
  anchor.className = 'pdf-viewer__annotation pdf-viewer__annotation--link'
  anchor.href = annotation.url ?? '#'
  // Electron's navigation guards (will-navigate / setWindowOpenHandler)
  // intercept this and route it through shell.openExternal after scheme
  // validation — this viewer doesn't need its own IPC round-trip.
  anchor.target = '_blank'
  anchor.rel = 'noopener noreferrer'
  anchor.title = annotation.url ?? ''
  return anchor
}

function createInternalLinkElement(
  annotation: RawPdfAnnotation,
  pdfDoc: PdfDocument,
  onInternalNavigate: (pageNumber: number) => void,
): HTMLAnchorElement {
  const anchor = document.createElement('a')
  anchor.className = 'pdf-viewer__annotation pdf-viewer__annotation--link'
  anchor.href = '#'
  anchor.addEventListener('click', (event) => {
    event.preventDefault()
    void resolveDestinationPage(
      (id: string) => pdfDoc.getDestination(id),
      (ref: PdfRef) => pdfDoc.getPageIndex(ref),
      annotation.dest ?? null,
    ).then((pageNumber) => {
      if (pageNumber !== null) {
        onInternalNavigate(pageNumber)
      }
    })
  })
  return anchor
}

function createTextFieldElement(annotation: RawPdfAnnotation, pdfDoc: PdfDocument): HTMLInputElement {
  const input = document.createElement('input')
  input.type = 'text'
  input.className = 'pdf-viewer__annotation pdf-viewer__annotation--text-field'
  input.value = typeof annotation.fieldValue === 'string' ? annotation.fieldValue : ''
  input.setAttribute('aria-label', annotation.fieldName ?? 'PDF form field')
  input.addEventListener('input', () => {
    pdfDoc.annotationStorage.setValue(annotation.id, { value: input.value })
  })
  return input
}

function createButtonWidgetElement(
  annotation: RawPdfAnnotation,
  kind: 'checkbox' | 'radio',
  pdfDoc: PdfDocument,
): HTMLInputElement {
  const input = document.createElement('input')
  input.type = kind
  input.className = `pdf-viewer__annotation pdf-viewer__annotation--${kind}`
  if (annotation.fieldName) {
    input.name = annotation.fieldName
  }
  input.checked = isWidgetChecked(annotation)
  input.setAttribute('aria-label', annotation.fieldName ?? 'PDF form field')
  input.addEventListener('change', () => {
    const value = kind === 'checkbox'
      ? (input.checked ? annotation.exportValue ?? 'Yes' : 'Off')
      : (input.checked ? annotation.buttonValue ?? 'Yes' : 'Off')
    pdfDoc.annotationStorage.setValue(annotation.id, { value })
  })
  return input
}

/**
 * Clears `container` and re-populates it with one positioned element per
 * clickable-link / fillable-form annotation on the page. Everything else
 * (`classifyAnnotation` returning 'ignored') is skipped entirely — the
 * canvas underneath already painted the annotation's static appearance
 * stream, if it had one.
 */
export function renderAnnotationOverlay(
  container: HTMLElement,
  annotations: ReadonlyArray<RawPdfAnnotation>,
  viewport: PageViewport,
  pdfDoc: PdfDocument,
  callbacks: AnnotationOverlayCallbacks,
): void {
  container.replaceChildren()

  for (const annotation of annotations) {
    if (!isValidAnnotationRect(annotation.rect)) continue

    const kind = classifyAnnotation(annotation)
    if (kind === 'ignored') continue

    const box = annotationRectToViewportBox(viewport, annotation.rect)
    if (box.width <= 0 || box.height <= 0) continue

    let element: HTMLElement
    switch (kind) {
      case 'external-link':
        element = createExternalLinkElement(annotation)
        break
      case 'internal-link':
        element = createInternalLinkElement(annotation, pdfDoc, callbacks.onInternalNavigate)
        break
      case 'text-field':
        element = createTextFieldElement(annotation, pdfDoc)
        break
      case 'checkbox':
      case 'radio':
        element = createButtonWidgetElement(annotation, kind, pdfDoc)
        break
      default:
        continue
    }

    applyBoxStyle(element, box)
    container.appendChild(element)
  }
}
