import { memo, useEffect, useRef, useState } from 'react'
import type { RenderTask } from 'pdfjs-dist'

import { resolveScaleForPage } from './geometry'
import { renderAnnotationOverlay } from './annotationOverlay'
import type { RawPdfAnnotation } from './annotations'
import type { PageGeometry, PageRotation, PdfDocument, PdfjsRuntime, TextLayerInstance, ZoomMode } from './types'

export type PdfPageProps = {
  readonly pdfjs: PdfjsRuntime
  readonly pdfDoc: PdfDocument
  readonly pageNumber: number
  /** Whether this page should have a live rendered canvas right now
   * (PDF-01) — false releases the canvas but keeps the wrapper's size. */
  readonly isActive: boolean
  readonly zoomMode: ZoomMode
  readonly rotation: PageRotation
  readonly dpr: number
  readonly containerWidth: number
  readonly containerHeight: number
  /** Known unscaled page size, once measured; undefined until then. */
  readonly geometry: PageGeometry | undefined
  readonly onGeometryResolved: (pageNumber: number, geometry: PageGeometry) => void
  readonly registerNode: (pageNumber: number, node: HTMLDivElement | null) => void
  readonly onInternalLinkNavigate: (pageNumber: number) => void
  /** Active find query (PDF-06); empty string when find is closed. */
  readonly findQuery: string
  /** Whether this page contains the currently-selected match. */
  readonly isActiveFindPage: boolean
  /** 0-based ordinal of the active match among this page's own matches, or
   * -1 if this isn't the active-find page. */
  readonly activeMatchLocalIndex: number
  readonly onRenderError?: (pageNumber: number, message: string) => void
}

const FALLBACK_GEOMETRY: PageGeometry = { width: 612, height: 792 } // US Letter, points

function clearHighlights(container: HTMLElement): void {
  const marks = container.querySelectorAll('mark.pdf-viewer__find-highlight')
  marks.forEach((mark) => {
    const parent = mark.parentNode
    if (!parent) return
    parent.replaceChild(document.createTextNode(mark.textContent ?? ''), mark)
    parent.normalize()
  })
}

/** Wraps every case-insensitive occurrence of `query` within `container`'s
 * text nodes in a `<mark>`, mirroring the DOM-rewrite approach `useSearch.ts`
 * already uses for markdown — scoped here to one page's text layer. Returns
 * the marks in document order so the caller can style/scroll to the active
 * one by its per-page ordinal. */
function highlightMatchesInTextLayer(container: HTMLElement, query: string): HTMLElement[] {
  clearHighlights(container)
  if (query.length === 0) return []

  const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT, null)
  const textNodes: Text[] = []
  let node: Node | null
  while ((node = walker.nextNode())) {
    textNodes.push(node as Text)
  }

  const lowerQuery = query.toLowerCase()
  const marks: HTMLElement[] = []

  for (const textNode of textNodes) {
    const text = textNode.textContent ?? ''
    const lowerText = text.toLowerCase()
    const positions: number[] = []
    let fromIndex = 0

    for (;;) {
      const index = lowerText.indexOf(lowerQuery, fromIndex)
      if (index === -1) break
      positions.push(index)
      fromIndex = index + lowerQuery.length
    }

    if (positions.length === 0) continue

    const fragment = document.createDocumentFragment()
    let lastEnd = 0
    for (const pos of positions) {
      if (pos > lastEnd) {
        fragment.appendChild(document.createTextNode(text.slice(lastEnd, pos)))
      }
      const mark = document.createElement('mark')
      mark.className = 'pdf-viewer__find-highlight'
      mark.textContent = text.slice(pos, pos + lowerQuery.length)
      fragment.appendChild(mark)
      marks.push(mark)
      lastEnd = pos + lowerQuery.length
    }
    if (lastEnd < text.length) {
      fragment.appendChild(document.createTextNode(text.slice(lastEnd)))
    }

    textNode.parentNode?.replaceChild(fragment, textNode)
  }

  return marks
}

function PdfPageBase({
  pdfjs,
  pdfDoc,
  pageNumber,
  isActive,
  zoomMode,
  rotation,
  dpr,
  containerWidth,
  containerHeight,
  geometry,
  onGeometryResolved,
  registerNode,
  onInternalLinkNavigate,
  findQuery,
  isActiveFindPage,
  activeMatchLocalIndex,
  onRenderError,
}: PdfPageProps) {
  const wrapperRef = useRef<HTMLDivElement | null>(null)
  const canvasRef = useRef<HTMLCanvasElement | null>(null)
  const textLayerRef = useRef<HTMLDivElement | null>(null)
  const annotationLayerRef = useRef<HTMLDivElement | null>(null)
  const [isTextLayerReady, setIsTextLayerReady] = useState(false)

  const knownGeometry = geometry ?? FALLBACK_GEOMETRY
  const scale = resolveScaleForPage(
    zoomMode,
    knownGeometry.width,
    knownGeometry.height,
    containerWidth,
    containerHeight,
  )
  const isLandscapeAfterRotation = rotation === 90 || rotation === 270
  const displayWidth = Math.ceil((isLandscapeAfterRotation ? knownGeometry.height : knownGeometry.width) * scale)
  const displayHeight = Math.ceil((isLandscapeAfterRotation ? knownGeometry.width : knownGeometry.height) * scale)

  const setWrapperNode = (node: HTMLDivElement | null) => {
    wrapperRef.current = node
    registerNode(pageNumber, node)
  }

  // Render (or release) this page's canvas + text layer + annotation layer.
  useEffect(() => {
    if (!isActive) {
      const canvas = canvasRef.current
      if (canvas) {
        canvas.width = 0
        canvas.height = 0
      }
      textLayerRef.current?.replaceChildren()
      annotationLayerRef.current?.replaceChildren()
      setIsTextLayerReady(false)
      return
    }

    let cancelled = false
    let renderTask: RenderTask | null = null
    let textLayerInstance: TextLayerInstance | null = null

    void (async () => {
      try {
        const page = await pdfDoc.getPage(pageNumber)
        if (cancelled) return

        let resolvedGeometry = geometry
        if (!resolvedGeometry) {
          const rawViewport = page.getViewport({ scale: 1, rotation: 0 })
          resolvedGeometry = { width: rawViewport.width, height: rawViewport.height }
          onGeometryResolved(pageNumber, resolvedGeometry)
        }

        const pageScale = resolveScaleForPage(
          zoomMode,
          resolvedGeometry.width,
          resolvedGeometry.height,
          containerWidth,
          containerHeight,
        )
        const viewport = page.getViewport({ scale: pageScale, rotation })

        const canvas = canvasRef.current
        if (!canvas) return

        canvas.width = Math.ceil(viewport.width * dpr)
        canvas.height = Math.ceil(viewport.height * dpr)
        canvas.style.width = `${Math.ceil(viewport.width)}px`
        canvas.style.height = `${Math.ceil(viewport.height)}px`

        const canvasContext = canvas.getContext('2d')
        if (!canvasContext) {
          throw new Error(`Failed to create a 2D canvas context for page ${pageNumber}.`)
        }
        canvasContext.setTransform(1, 0, 0, 1, 0, 0)
        canvasContext.scale(dpr, dpr)

        renderTask = page.render({ canvas, canvasContext, viewport })
        // `renderTask.cancel()` (called from this effect's cleanup, e.g. on
        // unmount or a zoom/rotation change) rejects this promise — but this
        // function may already have returned early below (the `if
        // (cancelled) return` guards while awaiting the text layer) by the
        // time that happens, leaving nothing awaiting it. Attaching a no-op
        // catch immediately avoids that becoming an unhandled rejection; the
        // real `await renderTask.promise` further down still throws/settles
        // independently for the actual error handling.
        renderTask.promise.catch(() => {})

        const textLayerDiv = textLayerRef.current
        const textContentPromise = page.getTextContent()
        const annotationsPromise = page.getAnnotations({ intent: 'display' })
        textContentPromise.catch(() => {})
        annotationsPromise.catch(() => {})

        if (textLayerDiv) {
          textLayerDiv.replaceChildren()
          textLayerDiv.style.setProperty('--total-scale-factor', String(pageScale))
          textLayerDiv.style.setProperty('--scale-round-x', '1px')
          textLayerDiv.style.setProperty('--scale-round-y', '1px')
          const textContent = await textContentPromise
          if (cancelled) return
          textLayerInstance = new pdfjs.TextLayer({
            textContentSource: textContent,
            container: textLayerDiv,
            viewport,
          })
          await textLayerInstance.render()
          if (cancelled) return
          setIsTextLayerReady(true)
        }

        await renderTask.promise
        if (cancelled) return

        const annotationLayerDiv = annotationLayerRef.current
        if (annotationLayerDiv) {
          const annotations = (await annotationsPromise) as ReadonlyArray<RawPdfAnnotation>
          if (cancelled) return
          renderAnnotationOverlay(annotationLayerDiv, annotations, viewport, pdfDoc, {
            onInternalNavigate: onInternalLinkNavigate,
          })
        }
      } catch (err) {
        if (err instanceof Error && err.name !== 'RenderingCancelledException' && !cancelled) {
          console.error(`Error rendering PDF page ${pageNumber}`, err)
          onRenderError?.(pageNumber, err.message)
        }
      }
    })()

    return () => {
      cancelled = true
      renderTask?.cancel()
      textLayerInstance?.cancel()
    }
    // containerWidth/containerHeight only matter for fit-width/fit-page —
    // included so a resize re-fits every active page (PDF-02). The parent
    // memoizes onGeometryResolved/onInternalLinkNavigate/onRenderError so
    // including them here doesn't cause extra re-renders.
  }, [
    pdfjs,
    pdfDoc,
    pageNumber,
    isActive,
    zoomMode,
    rotation,
    dpr,
    containerWidth,
    containerHeight,
    geometry,
    onGeometryResolved,
    onInternalLinkNavigate,
    onRenderError,
  ])

  // Find highlighting: independent of the render effect above so typing in
  // the find bar doesn't force a full page re-render (PDF-06).
  useEffect(() => {
    const textLayerDiv = textLayerRef.current
    if (!textLayerDiv || !isTextLayerReady) return

    if (findQuery.length === 0) {
      clearHighlights(textLayerDiv)
      return
    }

    const marks = highlightMatchesInTextLayer(textLayerDiv, findQuery)
    if (isActiveFindPage && activeMatchLocalIndex >= 0 && activeMatchLocalIndex < marks.length) {
      const activeMark = marks[activeMatchLocalIndex]
      activeMark.classList.add('pdf-viewer__find-highlight--active')
      activeMark.scrollIntoView({ behavior: 'smooth', block: 'center' })
    }
  }, [findQuery, isTextLayerReady, isActiveFindPage, activeMatchLocalIndex])

  return (
    <div
      ref={setWrapperNode}
      className="pdf-viewer__page-wrapper"
      data-page={pageNumber}
      style={{ width: displayWidth, height: displayHeight }}
    >
      {isActive && (
        <>
          <canvas className="pdf-viewer__page" ref={canvasRef} />
          <div className="pdf-viewer__text-layer" ref={textLayerRef} />
          <div className="pdf-viewer__annotation-layer" ref={annotationLayerRef} />
        </>
      )}
    </div>
  )
}

export const PdfPage = memo(PdfPageBase)
