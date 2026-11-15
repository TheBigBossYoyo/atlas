/**
 * DOM-building side of the annotation overlay (PDF-05/P5, PDF-11/P10).
 *
 * The pure classification/geometry math lives in `annotations.ts` (and is
 * unit tested there); this module is the thin, intentionally simple glue
 * that turns one classified annotation into a real positioned DOM element —
 * see `annotations.ts`'s file comment for why this hand-rolls the overlay
 * instead of instantiating pdf.js's own `AnnotationLayer` class.
 *
 * Form fields (PDF-EDIT-1). Every value the user enters goes into pdf.js's
 * `annotationStorage`, which is what `saveDocument()` writes and what printing
 * renders with `AnnotationMode.ENABLE_STORAGE`. Three rules follow from that:
 *   - an element is (re)built from the STORED value, falling back to the
 *     file's own value — the overlay is torn down whenever its page scrolls
 *     out of the render window, and rebuilding it from the file alone threw
 *     away whatever the user had typed;
 *   - checkbox and radio values are booleans, the type pdf.js's writer and
 *     renderer read (a string such as "Off" is truthy, so storing one saved an
 *     unchecked box as checked);
 *   - a radio button turns its siblings off in storage too, since each widget
 *     of a group is its own storage entry and pdf.js writes each one's state.
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
  /** Called after any form value changes, so the viewer can mark the document dirty. */
  readonly onFormChange?: () => void
}

type StoredValue = { readonly value?: unknown }

/** The DOM element of every rendered widget carries its storage id here, for sibling updates. */
const ANNOTATION_ID_ATTRIBUTE = 'data-annotation-id'

function readStoredValue(pdfDoc: PdfDocument, annotation: RawPdfAnnotation): unknown {
  const stored = pdfDoc.annotationStorage.getValue(annotation.id, {}) as StoredValue | undefined
  return stored?.value
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
    )
      .then((pageNumber) => {
        if (pageNumber !== null) {
          onInternalNavigate(pageNumber)
        }
      })
      .catch((err: unknown) => {
        // A malformed/dangling destination (e.g. a ref pointing at a page
        // that doesn't exist) must not surface as an unhandled promise
        // rejection — just leave the click a no-op.
        console.error('Failed to resolve internal PDF link destination', err)
      })
  })
  return anchor
}

/** Font size for a field's text: the field's own, scaled to the view, else sized to the box. */
function fieldFontSize(annotation: RawPdfAnnotation, viewport: PageViewport, boxHeight: number): number {
  const own = annotation.defaultAppearanceData?.fontSize
  if (typeof own === 'number' && own > 0) return own * viewport.scale
  return annotation.multiLine ? Math.min(12 * viewport.scale, boxHeight) : Math.max(6, boxHeight * 0.7)
}

function createTextFieldElement(
  annotation: RawPdfAnnotation,
  pdfDoc: PdfDocument,
  callbacks: AnnotationOverlayCallbacks,
): HTMLInputElement | HTMLTextAreaElement {
  const field: HTMLInputElement | HTMLTextAreaElement = annotation.multiLine
    ? document.createElement('textarea')
    : document.createElement('input')
  if (field instanceof HTMLInputElement) field.type = annotation.password ? 'password' : 'text'
  field.className = 'pdf-viewer__annotation pdf-viewer__annotation--text-field'
  const stored = readStoredValue(pdfDoc, annotation)
  field.value = typeof stored === 'string' ? stored : typeof annotation.fieldValue === 'string' ? annotation.fieldValue : ''
  if (typeof annotation.maxLen === 'number' && annotation.maxLen > 0) field.maxLength = annotation.maxLen
  if (annotation.readOnly) field.readOnly = true
  if (annotation.textAlignment === 1) field.style.textAlign = 'center'
  if (annotation.textAlignment === 2) field.style.textAlign = 'right'
  field.setAttribute('aria-label', annotation.fieldName ?? 'PDF form field')
  field.addEventListener('input', () => {
    pdfDoc.annotationStorage.setValue(annotation.id, { value: field.value })
    callbacks.onFormChange?.()
  })
  return field
}

function choiceOptionValue(option: { readonly exportValue?: unknown; readonly displayValue?: unknown }): string {
  return typeof option.exportValue === 'string' ? option.exportValue : String(option.displayValue ?? '')
}

function createChoiceElement(
  annotation: RawPdfAnnotation,
  pdfDoc: PdfDocument,
  callbacks: AnnotationOverlayCallbacks,
): HTMLSelectElement {
  const select = document.createElement('select')
  select.className = 'pdf-viewer__annotation pdf-viewer__annotation--choice'
  select.multiple = annotation.multiSelect === true
  if (!annotation.combo) select.size = Math.max(2, Math.min(annotation.options?.length ?? 2, 8))
  const stored = readStoredValue(pdfDoc, annotation)
  const selectedValues = new Set<string>(
    ([] as unknown[]).concat(stored ?? annotation.fieldValue ?? []).filter((v): v is string => typeof v === 'string'),
  )
  for (const option of annotation.options ?? []) {
    const element = document.createElement('option')
    element.value = choiceOptionValue(option)
    element.textContent = String(option.displayValue ?? option.exportValue ?? '')
    element.selected = selectedValues.has(element.value)
    select.appendChild(element)
  }
  if (annotation.combo && !select.multiple && selectedValues.size === 0) select.selectedIndex = -1
  if (annotation.readOnly) select.disabled = true
  select.setAttribute('aria-label', annotation.fieldName ?? 'PDF form field')
  select.addEventListener('change', () => {
    const values = Array.from(select.selectedOptions, (option) => option.value)
    pdfDoc.annotationStorage.setValue(annotation.id, { value: select.multiple ? values : (values[0] ?? '') })
    callbacks.onFormChange?.()
  })
  return select
}

/**
 * Field name -> storage ids of every widget of that field, across the whole
 * document (a radio group's buttons can sit on several pages). Loaded once per
 * document and shared by every page's overlay.
 */
const fieldWidgetIdsCache = new WeakMap<PdfDocument, Promise<ReadonlyMap<string, ReadonlyArray<string>>>>()

function loadFieldWidgetIds(pdfDoc: PdfDocument): Promise<ReadonlyMap<string, ReadonlyArray<string>>> {
  const cached = fieldWidgetIdsCache.get(pdfDoc)
  if (cached) return cached
  const loading = (async () => {
    const byName = new Map<string, string[]>()
    const fieldObjects: unknown = await pdfDoc.getFieldObjects?.()
    if (!fieldObjects) return byName
    const entries: Iterable<[string, unknown]> =
      fieldObjects instanceof Map ? fieldObjects.entries() : Object.entries(fieldObjects as Record<string, unknown>)
    for (const [name, widgets] of entries) {
      if (!Array.isArray(widgets)) continue
      const ids = widgets
        .map((widget: { id?: unknown }) => widget?.id)
        .filter((id): id is string => typeof id === 'string')
      byName.set(name, ids)
    }
    return byName
  })().catch(() => new Map<string, ReadonlyArray<string>>())
  fieldWidgetIdsCache.set(pdfDoc, loading)
  return loading
}

function createButtonWidgetElement(
  annotation: RawPdfAnnotation,
  kind: 'checkbox' | 'radio',
  pdfDoc: PdfDocument,
  callbacks: AnnotationOverlayCallbacks,
): HTMLInputElement {
  const input = document.createElement('input')
  input.type = kind
  input.className = `pdf-viewer__annotation pdf-viewer__annotation--${kind}`
  if (annotation.fieldName) {
    input.name = annotation.fieldName
  }
  const stored = readStoredValue(pdfDoc, annotation)
  input.checked = typeof stored === 'boolean' ? stored : isWidgetChecked(annotation)
  if (annotation.readOnly) input.disabled = true
  input.setAttribute('aria-label', annotation.fieldName ?? 'PDF form field')
  input.addEventListener('change', () => {
    const checked = input.checked
    pdfDoc.annotationStorage.setValue(annotation.id, { value: checked })
    callbacks.onFormChange?.()
    // Only a radio being turned ON affects its siblings; a lone checkbox's
    // state is its own.
    if (kind !== 'radio' || !checked || !annotation.fieldName) return
    const fieldName = annotation.fieldName
    void loadFieldWidgetIds(pdfDoc).then((idsByName) => {
      for (const siblingId of idsByName.get(fieldName) ?? []) {
        if (siblingId === annotation.id) continue
        pdfDoc.annotationStorage.setValue(siblingId, { value: false })
        // A sibling already on screen (same page or another rendered one).
        document.querySelectorAll<HTMLInputElement>(`input[${ANNOTATION_ID_ATTRIBUTE}]`).forEach((element) => {
          if (element.getAttribute(ANNOTATION_ID_ATTRIBUTE) === siblingId) element.checked = false
        })
      }
    })
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
        element = createTextFieldElement(annotation, pdfDoc, callbacks)
        element.style.fontSize = `${fieldFontSize(annotation, viewport, box.height)}px`
        break
      case 'choice':
        element = createChoiceElement(annotation, pdfDoc, callbacks)
        element.style.fontSize = `${fieldFontSize(annotation, viewport, box.height)}px`
        break
      case 'checkbox':
      case 'radio':
        element = createButtonWidgetElement(annotation, kind, pdfDoc, callbacks)
        break
      default:
        continue
    }

    if (kind !== 'external-link' && kind !== 'internal-link') {
      element.setAttribute(ANNOTATION_ID_ATTRIBUTE, annotation.id)
    }
    applyBoxStyle(element, box)
    container.appendChild(element)
  }
}
