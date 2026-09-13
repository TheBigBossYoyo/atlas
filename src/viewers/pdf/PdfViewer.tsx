import { memo, useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent } from 'react'

import type { NavItem, ViewerProps } from '../../formats/types'
import { useSetNavItems, useSetViewerStats } from '../shared/useViewerContext'
import '../__styles__/viewer-pdf.css'

import { buildFallbackNavItems, buildOutlineNavItems } from './outline'
import { DEFAULT_ZOOM, FIT_PADDING_PX, clampZoom, zoomIn as computeZoomIn, zoomOut as computeZoomOut } from './geometry'
import { rotateClockwise } from './rotation'
import {
  DEFAULT_OVERSCAN,
  clampPageNumber,
  computeActivePageWindow,
  includePendingJumpTarget,
  pickMostVisiblePage,
  updateVisibilityRatios,
  type PageRatioEntry,
} from './virtualization'
import { localMatchIndexOnPage } from './search'
import { usePdfFind } from './usePdfFind'
import { usePdfPageGeometry } from './usePdfPageGeometry'
import { runDetectingFakeWorkerFallback } from './workerSetup'
import { PdfPage } from './PdfPage'
import { PdfToolbar } from './PdfToolbar'
import { PdfFindBar } from './PdfFindBar'
import { PdfPasswordDialog } from './PdfPasswordDialog'
import { PdfThumbnailRail } from './PdfThumbnailRail'
import type { PageGeometry, PageRotation, PdfDocument, PdfOutlineNode, PdfjsRuntime, ZoomMode } from './types'

type PasswordPromptState = { readonly isIncorrect: boolean }

const DEFAULT_PAGE_GEOMETRY: PageGeometry = { width: 612, height: 792 }
const JUMP_SCROLL_TIMEOUT_MS = 600
const OBSERVER_ROOT_MARGIN = '200px 0px 200px 0px'
const OBSERVER_THRESHOLDS = [0, 0.1, 0.25, 0.5, 0.75, 0.9, 1]
/** PDF points -> CSS px at 100% print scale (96 CSS px/in over 72 pt/in). */
const PDF_POINTS_TO_CSS_PX = 96 / 72

function PdfViewerBase({ file }: ViewerProps) {
  const containerRef = useRef<HTMLDivElement | null>(null)
  const viewerRef = useRef<HTMLDivElement | null>(null)
  const printRootRef = useRef<HTMLDivElement | null>(null)

  const [errorMessage, setErrorMessage] = useState<string | null>(null)
  const [pageCount, setPageCount] = useState(0)
  const [currentPage, setCurrentPage] = useState(1)
  const [pageInput, setPageInput] = useState('1')
  const [zoomMode, setZoomMode] = useState<ZoomMode>(DEFAULT_ZOOM)
  const [isZoomMenuOpen, setIsZoomMenuOpen] = useState(false)
  const [rotation, setRotation] = useState<PageRotation>(0)
  const [pdfDoc, setPdfDoc] = useState<PdfDocument | null>(null)
  const [pdfjs, setPdfjs] = useState<PdfjsRuntime | null>(null)
  const [isThumbnailRailOpen, setIsThumbnailRailOpen] = useState(false)
  const [passwordPrompt, setPasswordPrompt] = useState<PasswordPromptState | null>(null)
  // Bumped on every onPassword callback so <PdfPasswordDialog key={...}>
  // remounts (starting the input empty) instead of the dialog resetting its
  // own state in an effect.
  const [passwordAttemptId, setPasswordAttemptId] = useState(0)
  const [isPrinting, setIsPrinting] = useState(false)
  const [dpr, setDpr] = useState(() => (typeof window === 'undefined' ? 1 : window.devicePixelRatio || 1))
  const [containerSize, setContainerSize] = useState({ width: 800, height: 600 })

  const [intersectingPages, setIntersectingPages] = useState<ReadonlyArray<number>>([])
  const [visibilityRatios, setVisibilityRatios] = useState<ReadonlyMap<number, number>>(new Map())
  const [pendingJumpPage, setPendingJumpPage] = useState<number | null>(null)

  const currentPageRef = useRef(1)
  const pageCountRef = useRef(0)
  const nodesRef = useRef<Map<number, HTMLDivElement>>(new Map())
  const observerRef = useRef<IntersectionObserver | null>(null)
  const isProgrammaticScrollRef = useRef(false)
  const updatePasswordRef = useRef<((password: string) => void) | null>(null)
  const destroyLoadingTaskRef = useRef<(() => Promise<void>) | null>(null)
  const cancelledPasswordRef = useRef(false)

  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()

  const { geometry: pageGeometry, reportGeometry } = usePdfPageGeometry(pdfDoc, pageCount)

  useEffect(() => {
    currentPageRef.current = currentPage
  }, [currentPage])

  useEffect(() => {
    pageCountRef.current = pageCount
  }, [pageCount])

  // ---- Page navigation ----------------------------------------------------

  const scrollToPage = useCallback((pageNumber: number) => {
    const clamped = clampPageNumber(pageNumber, pageCountRef.current)
    setPendingJumpPage(clamped)
    setCurrentPage(clamped)
    setPageInput(String(clamped))
  }, [])

  // Actually perform the scroll once the pending target's wrapper is
  // mounted — it's guaranteed to mount because `includePendingJumpTarget`
  // folds it into the active window even when it wasn't already
  // intersecting (PDF-07: previously a jump to an unrendered page silently
  // did nothing).
  useEffect(() => {
    if (pendingJumpPage === null) return undefined
    const node = nodesRef.current.get(pendingJumpPage)
    if (!node) return undefined

    isProgrammaticScrollRef.current = true
    node.scrollIntoView({ behavior: 'smooth', block: 'start' })
    const timeout = setTimeout(() => {
      isProgrammaticScrollRef.current = false
      setPendingJumpPage(null)
    }, JUMP_SCROLL_TIMEOUT_MS)
    return () => clearTimeout(timeout)
    // Re-checks whenever the active window changes — that's what causes the
    // pending target's node to actually mount.
  }, [pendingJumpPage, intersectingPages])

  const handlePrevPage = useCallback(() => {
    if (currentPageRef.current > 1) scrollToPage(currentPageRef.current - 1)
  }, [scrollToPage])

  const handleNextPage = useCallback(() => {
    if (currentPageRef.current < pageCountRef.current) scrollToPage(currentPageRef.current + 1)
  }, [scrollToPage])

  const handlePageInputKeyDown = useCallback(
    (e: ReactKeyboardEvent<HTMLInputElement>) => {
      if (e.key === 'Enter') {
        const parsed = parseInt(pageInput, 10)
        scrollToPage(Number.isNaN(parsed) ? currentPageRef.current : parsed)
      }
    },
    [pageInput, scrollToPage],
  )

  const handlePageInputBlur = useCallback(() => {
    setPageInput(String(currentPageRef.current))
  }, [])

  // ---- Zoom / rotation ------------------------------------------------------

  const handleZoomIn = useCallback(() => {
    setZoomMode((prev) => computeZoomIn(typeof prev === 'number' ? prev : 1))
  }, [])

  const handleZoomOut = useCallback(() => {
    setZoomMode((prev) => computeZoomOut(typeof prev === 'number' ? prev : 1))
  }, [])

  const handleRotate = useCallback(() => {
    setRotation((prev) => rotateClockwise(prev))
  }, [])

  // ---- Print (PDF-10/P9) ------------------------------------------------

  const handlePrint = useCallback(async () => {
    if (!pdfDoc || pageCount === 0 || isPrinting) return
    const printRoot = printRootRef.current
    if (!printRoot) return

    setIsPrinting(true)
    printRoot.replaceChildren()

    try {
      for (let pageNumber = 1; pageNumber <= pageCount; pageNumber += 1) {
        const page = await pdfDoc.getPage(pageNumber)
        const viewport = page.getViewport({ scale: PDF_POINTS_TO_CSS_PX, rotation })

        const canvas = document.createElement('canvas')
        canvas.className = 'pdf-viewer__print-page-canvas'
        canvas.width = Math.ceil(viewport.width)
        canvas.height = Math.ceil(viewport.height)
        const canvasContext = canvas.getContext('2d')
        if (canvasContext) {
          await page.render({ canvas, canvasContext, viewport }).promise
        }

        const pageDiv = document.createElement('div')
        pageDiv.className = 'pdf-viewer__print-page'
        pageDiv.appendChild(canvas)
        printRoot.appendChild(pageDiv)
        page.cleanup()
      }

      // Mirrors DocxViewer's print pattern: the print stylesheet in
      // viewer-pdf.css keys off this class + @media print to show only
      // `.pdf-viewer__print-root` and hide the interactive toolbar/pages.
      document.body.classList.add('atlas-printing')
      window.print()
    } finally {
      document.body.classList.remove('atlas-printing')
      printRoot.replaceChildren()
      setIsPrinting(false)
    }
  }, [pdfDoc, pageCount, rotation, isPrinting])

  // ---- Find (PDF-06/P6) --------------------------------------------------

  const find = usePdfFind(pdfDoc, pageCount, scrollToPage)
  const activeMatch = find.matches[find.currentMatchIndex]
  const activeMatchLocalIndex = activeMatch
    ? localMatchIndexOnPage(find.matches, find.currentMatchIndex)
    : -1

  // find.isOpen/open/close are read via a ref inside the keydown handler
  // further below instead of that effect's dependency array — `find`
  // (usePdfFind's return value) is a fresh object every render, and
  // depending on it directly would tear down/rebuild the window listener on
  // every render.
  const findIsOpenRef = useRef(find.isOpen)
  useEffect(() => {
    findIsOpenRef.current = find.isOpen
  }, [find.isOpen])
  const findOpen = find.open
  const findClose = find.close

  // ---- Container size (P2: drives fit-width/fit-page re-fit on resize) --

  useEffect(() => {
    const viewer = viewerRef.current
    if (!viewer) return undefined

    const measure = () => {
      setContainerSize({
        width: Math.max(0, viewer.clientWidth - FIT_PADDING_PX),
        height: Math.max(0, viewer.clientHeight - FIT_PADDING_PX),
      })
    }
    measure()

    const observer = new ResizeObserver(() => measure())
    observer.observe(viewer)
    return () => observer.disconnect()
  }, [])

  // ---- Live devicePixelRatio changes (PDF-18/P11) ------------------------

  useEffect(() => {
    if (typeof window === 'undefined' || typeof window.matchMedia !== 'function') {
      return undefined
    }
    const mediaQuery = window.matchMedia(`(resolution: ${dpr}dppx)`)
    const handleChange = () => setDpr(window.devicePixelRatio || 1)
    mediaQuery.addEventListener('change', handleChange)
    return () => mediaQuery.removeEventListener('change', handleChange)
  }, [dpr])

  // ---- Virtualization (PDF-01/P1, PDF-08/P7) -----------------------------

  const registerPageNode = useCallback((pageNumber: number, node: HTMLDivElement | null) => {
    const observer = observerRef.current
    const previous = nodesRef.current.get(pageNumber)
    if (previous && observer) observer.unobserve(previous)

    if (node) {
      nodesRef.current.set(pageNumber, node)
      observer?.observe(node)
    } else {
      nodesRef.current.delete(pageNumber)
    }
  }, [])

  useEffect(() => {
    const viewer = viewerRef.current
    if (!viewer || pageCount === 0) return undefined

    const observer = new IntersectionObserver(
      (entries) => {
        setIntersectingPages((prev) => {
          const set = new Set(prev)
          for (const entry of entries) {
            const pageNumber = Number((entry.target as HTMLElement).dataset.page)
            if (entry.isIntersecting) {
              set.add(pageNumber)
            } else {
              set.delete(pageNumber)
            }
          }
          return [...set]
        })

        if (!isProgrammaticScrollRef.current) {
          const ratioEntries: PageRatioEntry[] = entries.map((entry) => ({
            pageNumber: Number((entry.target as HTMLElement).dataset.page),
            ratio: entry.intersectionRatio,
            isIntersecting: entry.isIntersecting,
          }))
          setVisibilityRatios((prev) => updateVisibilityRatios(prev, ratioEntries))
        }
      },
      { root: viewer, rootMargin: OBSERVER_ROOT_MARGIN, threshold: OBSERVER_THRESHOLDS },
    )
    observerRef.current = observer
    nodesRef.current.forEach((node) => observer.observe(node))

    return () => {
      observer.disconnect()
      observerRef.current = null
    }
  }, [pageCount])

  const activePages = useMemo(
    () =>
      includePendingJumpTarget(
        computeActivePageWindow(intersectingPages, pageCount, DEFAULT_OVERSCAN),
        pendingJumpPage,
        pageCount,
      ),
    [intersectingPages, pageCount, pendingJumpPage],
  )

  // Derives the "current page" indicator from whichever page has the
  // greatest on-screen visibility ratio right now (PDF-08) — skipped while a
  // programmatic jump is in flight so it doesn't fight that scroll.
  useEffect(() => {
    if (pendingJumpPage !== null) return
    const mostVisible = pickMostVisiblePage(visibilityRatios, currentPageRef.current)
    if (mostVisible !== currentPageRef.current) {
      setCurrentPage(mostVisible)
      setPageInput(String(mostVisible))
    }
  }, [visibilityRatios, pendingJumpPage])

  // ---- Keyboard shortcuts (PDF-03/P3 guards the page-number input) -------

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      const container = containerRef.current
      if (!container) return

      const isFocusInside =
        document.activeElement === document.body || container.contains(document.activeElement)
      if (!isFocusInside) return

      const isEditableTarget =
        e.target instanceof HTMLInputElement || e.target instanceof HTMLTextAreaElement

      if (e.ctrlKey || e.metaKey) {
        if (e.key === '=' || e.key === '+') {
          e.preventDefault()
          handleZoomIn()
        } else if (e.key === '-') {
          e.preventDefault()
          handleZoomOut()
        } else if (e.key === '0') {
          e.preventDefault()
          setZoomMode(DEFAULT_ZOOM)
        } else if (e.key.toLowerCase() === 'f') {
          e.preventDefault()
          if (findIsOpenRef.current) findClose()
          else findOpen()
        } else if (e.key.toLowerCase() === 'p') {
          e.preventDefault()
          void handlePrint()
        }
        return
      }

      // Everything below moves the document/cursor — must never fire while
      // the user is typing in the page-number input (PDF-03) or find bar.
      if (isEditableTarget) return

      if (e.key === 'ArrowDown' || e.key === 'PageDown' || (e.key === ' ' && !e.shiftKey)) {
        e.preventDefault()
        handleNextPage()
      } else if (e.key === 'ArrowUp' || e.key === 'PageUp' || (e.key === ' ' && e.shiftKey)) {
        e.preventDefault()
        handlePrevPage()
      } else if (e.key === 'Home') {
        e.preventDefault()
        scrollToPage(1)
      } else if (e.key === 'End') {
        e.preventDefault()
        scrollToPage(pageCountRef.current)
      }
    }

    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [handleNextPage, handlePrevPage, scrollToPage, handleZoomIn, handleZoomOut, handlePrint, findOpen, findClose])

  // Click outside to close the zoom menu.
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (isZoomMenuOpen && !(e.target as Element).closest('.pdf-viewer__zoom-menu-container')) {
        setIsZoomMenuOpen(false)
      }
    }
    document.addEventListener('mousedown', handleClickOutside)
    return () => document.removeEventListener('mousedown', handleClickOutside)
  }, [isZoomMenuOpen])

  // ---- Document load (PDF-14/P10 password prompt, PDF-15/P11 fake-worker) -

  useEffect(() => {
    setErrorMessage(null)
    setNavItems([])
    setStats(null)
    setPdfDoc(null)
    setPdfjs(null)
    setPageCount(0)
    setCurrentPage(1)
    setPageInput('1')
    setPasswordPrompt(null)
    cancelledPasswordRef.current = false

    if (file.kind !== 'binary') {
      setErrorMessage('PdfViewer expected a binary file.')
      return undefined
    }

    let cancelled = false

    void (async () => {
      try {
        const pdfjsModule = await import('pdfjs-dist/legacy/build/pdf.mjs')
        if (cancelled) return

        pdfjsModule.GlobalWorkerOptions.workerSrc = new URL(
          'pdfjs-dist/legacy/build/pdf.worker.min.mjs',
          import.meta.url,
        ).toString()

        const nextLoadingTask = pdfjsModule.getDocument({ data: file.content.slice(0) })
        // pdfjs-dist 6.x moved document teardown onto the loading task —
        // PDFDocumentProxy no longer exposes its own destroy() method.
        destroyLoadingTaskRef.current = () => nextLoadingTask.destroy()

        nextLoadingTask.onPassword = (updatePassword: (password: string) => void, reason: number) => {
          updatePasswordRef.current = updatePassword
          setPasswordPrompt({ isIncorrect: reason === pdfjsModule.PasswordResponses.INCORRECT_PASSWORD })
          setPasswordAttemptId((id) => id + 1)
        }

        const { result: pdf, usedFakeWorker } = await runDetectingFakeWorkerFallback(() =>
          nextLoadingTask.promise,
        )
        if (usedFakeWorker) {
          console.warn(
            '[PdfViewer] pdfjs-dist could not load its module worker and fell back to ' +
              'main-thread rendering. Large PDFs may block the UI while rendering — this ' +
              'usually means the worker script failed to resolve under the current build ' +
              '(e.g. a packaged file:// build serving it from an unexpected path).',
          )
        }

        if (cancelled) {
          await nextLoadingTask.destroy()
          return
        }

        setPasswordPrompt(null)
        setPdfjs(pdfjsModule)
        setPdfDoc(pdf)
        setPageCount(pdf.numPages)
        setStats({ kind: 'pdf', page: 1, pageCount: pdf.numPages })

        const outline = (await pdf.getOutline()) as ReadonlyArray<PdfOutlineNode> | null
        if (cancelled) return

        const navItems: NavItem[] =
          outline && outline.length > 0
            ? await buildOutlineNavItems(
                outline,
                scrollToPage,
                pdf.getDestination.bind(pdf),
                pdf.getPageIndex.bind(pdf),
              )
            : []

        if (cancelled) return
        setNavItems(navItems.length > 0 ? navItems : buildFallbackNavItems(pdf.numPages, scrollToPage))
      } catch (err) {
        if (!cancelled) {
          setNavItems([])
          setStats(null)
          if (!cancelledPasswordRef.current) {
            setErrorMessage(err instanceof Error ? err.message : 'Failed to render PDF.')
          }
        }
      }
    })()

    return () => {
      cancelled = true
      setPdfDoc(null)
    }
  }, [file, setNavItems, setStats, scrollToPage])

  useEffect(() => {
    if (pageCount > 0) {
      setStats({ kind: 'pdf', page: currentPage, pageCount })
    }
  }, [currentPage, pageCount, setStats])

  const handlePasswordSubmit = useCallback((password: string) => {
    updatePasswordRef.current?.(password)
  }, [])

  const handlePasswordCancel = useCallback(() => {
    updatePasswordRef.current = null
    cancelledPasswordRef.current = true
    setPasswordPrompt(null)
    setErrorMessage('This PDF is password protected.')
    void destroyLoadingTaskRef.current?.()
  }, [])

  const handleInternalLinkNavigate = useCallback((pageNumber: number) => scrollToPage(pageNumber), [scrollToPage])

  const actualZoomForDisplay =
    typeof zoomMode === 'number'
      ? clampZoom(zoomMode)
      : (() => {
          const geom = pageGeometry.get(currentPage) ?? DEFAULT_PAGE_GEOMETRY
          const isLandscape = rotation === 90 || rotation === 270
          const width = isLandscape ? geom.height : geom.width
          const height = isLandscape ? geom.width : geom.height
          return zoomMode === 'fit-width'
            ? clampZoom(containerSize.width / Math.max(1, width))
            : clampZoom(
                Math.min(containerSize.width / Math.max(1, width), containerSize.height / Math.max(1, height)),
              )
        })()

  if (errorMessage !== null) {
    return <div className="pdf-viewer__error">Failed to render PDF: {errorMessage}</div>
  }

  return (
    <div className="pdf-viewer-container" ref={containerRef} tabIndex={-1}>
      <PdfToolbar
        currentPage={currentPage}
        pageCount={pageCount}
        pageInput={pageInput}
        onPageInputChange={setPageInput}
        onPageInputKeyDown={handlePageInputKeyDown}
        onPageInputBlur={handlePageInputBlur}
        onPrevPage={handlePrevPage}
        onNextPage={handleNextPage}
        zoomMode={zoomMode}
        actualZoom={actualZoomForDisplay}
        isZoomMenuOpen={isZoomMenuOpen}
        onToggleZoomMenu={() => setIsZoomMenuOpen((open) => !open)}
        onSelectZoom={(mode) => {
          setZoomMode(mode)
          setIsZoomMenuOpen(false)
        }}
        onZoomIn={handleZoomIn}
        onZoomOut={handleZoomOut}
        onRotate={handleRotate}
        isThumbnailRailOpen={isThumbnailRailOpen}
        onToggleThumbnailRail={() => setIsThumbnailRailOpen((open) => !open)}
        onToggleFind={() => (find.isOpen ? find.close() : find.open())}
        onPrint={() => void handlePrint()}
      />

      <PdfFindBar
        isOpen={find.isOpen}
        query={find.query}
        matchCount={find.matches.length}
        currentMatchIndex={find.currentMatchIndex}
        isIndexing={find.isIndexing}
        indexedPageCount={find.indexedPageCount}
        totalPageCount={pageCount}
        onQueryChange={find.setQuery}
        onNext={find.next}
        onPrev={find.prev}
        onClose={find.close}
      />

      <div className="pdf-viewer__body">
        {isThumbnailRailOpen && pdfDoc && (
          <PdfThumbnailRail
            isOpen={isThumbnailRailOpen}
            pdfDoc={pdfDoc}
            pageCount={pageCount}
            currentPage={currentPage}
            rotation={rotation}
            onSelectPage={scrollToPage}
          />
        )}

        <div ref={viewerRef} className="pdf-viewer">
          {pdfDoc &&
            pdfjs &&
            pageCount > 0 &&
            Array.from({ length: pageCount }, (_, index) => {
              const pageNumber = index + 1
              const isActiveFindPage = activeMatch?.pageNumber === pageNumber
              return (
                <PdfPage
                  key={pageNumber}
                  pdfjs={pdfjs}
                  pdfDoc={pdfDoc}
                  pageNumber={pageNumber}
                  isActive={activePages.has(pageNumber)}
                  zoomMode={zoomMode}
                  rotation={rotation}
                  dpr={dpr}
                  containerWidth={containerSize.width}
                  containerHeight={containerSize.height}
                  geometry={pageGeometry.get(pageNumber)}
                  onGeometryResolved={reportGeometry}
                  registerNode={registerPageNode}
                  onInternalLinkNavigate={handleInternalLinkNavigate}
                  findQuery={find.query}
                  isActiveFindPage={isActiveFindPage}
                  activeMatchLocalIndex={isActiveFindPage ? activeMatchLocalIndex : -1}
                />
              )
            })}
        </div>
      </div>

      <div ref={printRootRef} className="pdf-viewer__print-root" aria-hidden="true" />
      {isPrinting && <div className="pdf-viewer__print-status" role="status">Preparing document for print…</div>}

      <PdfPasswordDialog
        key={passwordAttemptId}
        isOpen={passwordPrompt !== null}
        isIncorrect={passwordPrompt?.isIncorrect ?? false}
        onSubmit={handlePasswordSubmit}
        onCancel={handlePasswordCancel}
      />
    </div>
  )
}

export const PdfViewer = memo(PdfViewerBase)
