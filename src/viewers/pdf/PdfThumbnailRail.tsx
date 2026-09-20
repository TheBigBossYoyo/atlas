import { memo, useEffect, useMemo, useRef, useState } from 'react'

import { computeActivePageWindow } from './virtualization'
import { combineRotation } from './rotation'
import type { PageRotation, PdfDocument } from './types'
import { useTranslate } from '../../i18n'

const THUMBNAIL_WIDTH = 120
const THUMBNAIL_OVERSCAN = 6

type PdfThumbnailProps = {
  readonly pdfDoc: PdfDocument
  readonly pageNumber: number
  readonly rotation: PageRotation
  readonly isActive: boolean
  readonly isCurrent: boolean
  readonly onSelect: (pageNumber: number) => void
  readonly registerNode: (pageNumber: number, node: HTMLButtonElement | null) => void
}

function PdfThumbnailBase({
  pdfDoc,
  pageNumber,
  rotation,
  isActive,
  isCurrent,
  onSelect,
  registerNode,
}: PdfThumbnailProps) {
  const t = useTranslate()
  const canvasRef = useRef<HTMLCanvasElement | null>(null)

  useEffect(() => {
    if (!isActive) {
      const canvas = canvasRef.current
      if (canvas) {
        canvas.width = 0
        canvas.height = 0
      }
      return
    }

    let cancelled = false

    void (async () => {
      try {
        const page = await pdfDoc.getPage(pageNumber)
        if (cancelled) return
        // TOTAL rotation (intrinsic /Rotate + the Rotate-button state) so a
        // thumbnail's orientation always matches the corresponding main page
        // (see rotation.ts's `combineRotation`).
        const totalRotation = combineRotation(page.rotate, rotation)
        const unscaledViewport = page.getViewport({ scale: 1, rotation: totalRotation })
        const scale = THUMBNAIL_WIDTH / unscaledViewport.width
        const viewport = page.getViewport({ scale, rotation: totalRotation })

        const canvas = canvasRef.current
        if (!canvas) return
        canvas.width = Math.ceil(viewport.width)
        canvas.height = Math.ceil(viewport.height)
        const canvasContext = canvas.getContext('2d')
        if (!canvasContext) return

        await page.render({ canvas, canvasContext, viewport }).promise
      } catch (err) {
        if (err instanceof Error && err.name !== 'RenderingCancelledException') {
          console.error(`Error rendering thumbnail for page ${pageNumber}`, err)
        }
      }
    })()

    return () => {
      cancelled = true
    }
  }, [pdfDoc, pageNumber, rotation, isActive])

  return (
    <button
      ref={(node) => registerNode(pageNumber, node)}
      type="button"
      className={`pdf-viewer__thumbnail ${isCurrent ? 'pdf-viewer__thumbnail--current' : ''}`}
      data-page={pageNumber}
      onClick={() => onSelect(pageNumber)}
      aria-label={t('pdf.thumbnails.goToPage', { n: pageNumber })}
      aria-current={isCurrent}
    >
      <canvas className="pdf-viewer__thumbnail-canvas" ref={canvasRef} />
      <span className="pdf-viewer__thumbnail-label">{pageNumber}</span>
    </button>
  )
}

const PdfThumbnail = memo(PdfThumbnailBase)

export type PdfThumbnailRailProps = {
  readonly isOpen: boolean
  readonly pdfDoc: PdfDocument
  readonly pageCount: number
  readonly currentPage: number
  readonly rotation: PageRotation
  readonly onSelectPage: (pageNumber: number) => void
}

/**
 * Collapsible sidebar of lazily-rendered per-page thumbnails (PDF-13/P10).
 * `NavItem` (the app-wide sidebar's data model) only carries a label/icon,
 * not an image, so this renders its own small panel inside the PDF viewer
 * rather than trying to route thumbnails through the shared Sidebar.
 * Bounded-memory the same way as the main page canvases (P1): only
 * thumbnails within the scrolled viewport (± overscan) keep a live canvas.
 */
export function PdfThumbnailRail({
  isOpen,
  pdfDoc,
  pageCount,
  currentPage,
  rotation,
  onSelectPage,
}: PdfThumbnailRailProps) {
  const t = useTranslate()
  const railRef = useRef<HTMLDivElement | null>(null)
  const nodesRef = useRef<Map<number, HTMLButtonElement>>(new Map())
  const observerRef = useRef<IntersectionObserver | null>(null)
  const [visiblePages, setVisiblePages] = useState<ReadonlyArray<number>>([])

  const registerNode = (pageNumber: number, node: HTMLButtonElement | null) => {
    const observer = observerRef.current
    const previous = nodesRef.current.get(pageNumber)
    if (previous && observer) observer.unobserve(previous)

    if (node) {
      nodesRef.current.set(pageNumber, node)
      observer?.observe(node)
    } else {
      nodesRef.current.delete(pageNumber)
    }
  }

  useEffect(() => {
    if (!isOpen) return undefined
    const rail = railRef.current
    if (!rail) return undefined

    const observer = new IntersectionObserver(
      (entries) => {
        setVisiblePages((prev) => {
          const next = new Set(prev)
          for (const entry of entries) {
            const pageNumber = Number((entry.target as HTMLElement).dataset.page)
            if (entry.isIntersecting) {
              next.add(pageNumber)
            } else {
              next.delete(pageNumber)
            }
          }
          return [...next]
        })
      },
      { root: rail, rootMargin: '200px 0px' },
    )
    observerRef.current = observer
    nodesRef.current.forEach((node) => observer.observe(node))

    return () => {
      observer.disconnect()
      observerRef.current = null
    }
  }, [isOpen])

  const activePages = useMemo(
    () => computeActivePageWindow(visiblePages, pageCount, THUMBNAIL_OVERSCAN),
    [visiblePages, pageCount],
  )

  if (!isOpen) return null

  return (
    <div className="pdf-viewer__thumbnail-rail" ref={railRef} role="group" aria-label={t('pdf.thumbnails.railAria')}>
      {Array.from({ length: pageCount }, (_, index) => {
        const pageNumber = index + 1
        return (
          <PdfThumbnail
            key={pageNumber}
            pdfDoc={pdfDoc}
            pageNumber={pageNumber}
            rotation={rotation}
            isActive={activePages.has(pageNumber)}
            isCurrent={pageNumber === currentPage}
            onSelect={onSelectPage}
            registerNode={registerNode}
          />
        )
      })}
    </div>
  )
}
