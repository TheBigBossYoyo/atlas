import { memo, useEffect, useRef, useState, useCallback } from 'react'
import { ChevronUp, ChevronDown, ZoomIn, ZoomOut, ChevronDown as ChevronDownSmall } from 'lucide-react'
import './__styles__/viewer-pdf.css'

import type { NavItem, ViewerProps } from '../formats/types'
import { useSetNavItems, useSetViewerStats } from './shared/useViewerContext'

import type { PDFDocumentProxy } from 'pdfjs-dist'

type PdfDocument = PDFDocumentProxy

type PdfDestination = string | readonly unknown[] | null

type PdfOutlineNode = {
  readonly title: string
  readonly dest: PdfDestination
  readonly items: ReadonlyArray<PdfOutlineNode>
}

type PdfRef = {
  readonly num: number
  readonly gen: number
}

function isPdfRef(value: unknown): value is PdfRef {
  if (typeof value !== 'object' || value === null) {
    return false
  }

  const maybeRef = value as Record<string, unknown>
  return typeof maybeRef.num === 'number' && typeof maybeRef.gen === 'number'
}

async function resolveDestinationPage(
  getDestination: (id: string) => Promise<readonly unknown[] | null>,
  getPageIndex: (ref: PdfRef) => Promise<number>,
  dest: PdfDestination,
): Promise<number | null> {
  const resolvedDest = typeof dest === 'string' ? await getDestination(dest) : dest

  if (!resolvedDest || resolvedDest.length === 0) {
    return null
  }

  const target = resolvedDest[0]

  if (typeof target === 'number' && Number.isFinite(target)) {
    return target + 1
  }

  if (isPdfRef(target)) {
    return (await getPageIndex(target)) + 1
  }

  return null
}

async function buildOutlineNavItems(
  outline: ReadonlyArray<PdfOutlineNode>,
  scrollToPage: (pageNumber: number) => void,
  getDestination: (id: string) => Promise<readonly unknown[] | null>,
  getPageIndex: (ref: PdfRef) => Promise<number>,
  level = 1,
): Promise<NavItem[]> {
  const items: NavItem[] = []

  for (const node of outline) {
    const pageNumber = await resolveDestinationPage(
      getDestination,
      getPageIndex,
      node.dest,
    )
    const label = node.title.trim() || 'Untitled'

    items.push({
      id: pageNumber === null ? `${level}-${label}` : `page-${pageNumber}-${label}`,
      label,
      level,
      onSelect: () => {
        if (pageNumber !== null) {
          scrollToPage(pageNumber)
        }
      },
    })

    if (node.items.length > 0) {
      items.push(
        ...(await buildOutlineNavItems(
          node.items,
          scrollToPage,
          getDestination,
          getPageIndex,
          level + 1,
        )),
      )
    }
  }

  return items
}

function buildFallbackNavItems(
  pageCount: number,
  scrollToPage: (pageNumber: number) => void,
): NavItem[] {
  return Array.from({ length: pageCount }, (_, index) => {
    const pageNumber = index + 1

    return {
      id: `page-${pageNumber}`,
      label: `Page ${pageNumber}`,
      level: 1,
      onSelect: () => scrollToPage(pageNumber),
    }
  })
}

const ZOOM_PRESETS = [0.5, 0.75, 1, 1.25, 1.5, 2]

function PdfViewerBase({ file }: ViewerProps) {
  const containerRef = useRef<HTMLDivElement | null>(null)
  const viewerRef = useRef<HTMLDivElement | null>(null)
  const [errorMessage, setErrorMessage] = useState<string | null>(null)
  
  const [pageCount, setPageCount] = useState(0)
  const [currentPage, setCurrentPage] = useState(1)
  const [pageInput, setPageInput] = useState('1')
  
  const [zoomMode, setZoomMode] = useState<number | 'fit-width' | 'fit-page'>(1)
  const [actualZoom, setActualZoom] = useState(1)
  const [isZoomMenuOpen, setIsZoomMenuOpen] = useState(false)
  
  const [pdfDoc, setPdfDoc] = useState<PdfDocument | null>(null)
  const renderTasksRef = useRef<Map<number, { cancel: () => void }>>(new Map())
  const intersectionObserverRef = useRef<IntersectionObserver | null>(null)
  const currentPageRef = useRef(1)
  const isProgrammaticScrollRef = useRef(false)
  const programmaticScrollTimeoutRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const scrollEndHandlerRef = useRef<(() => void) | null>(null)

  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()

  useEffect(() => {
    currentPageRef.current = currentPage
  }, [currentPage])

  // Click outside to close zoom menu
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (isZoomMenuOpen && !(e.target as Element).closest('.pdf-viewer__zoom-menu-container')) {
        setIsZoomMenuOpen(false)
      }
    }
    document.addEventListener('mousedown', handleClickOutside)
    return () => document.removeEventListener('mousedown', handleClickOutside)
  }, [isZoomMenuOpen])

  const finishProgrammaticScroll = useCallback(() => {
    const viewer = viewerRef.current

    isProgrammaticScrollRef.current = false

    if (programmaticScrollTimeoutRef.current !== null) {
      clearTimeout(programmaticScrollTimeoutRef.current)
      programmaticScrollTimeoutRef.current = null
    }

    if (viewer && scrollEndHandlerRef.current) {
      viewer.removeEventListener('scrollend', scrollEndHandlerRef.current)
      scrollEndHandlerRef.current = null
    }
  }, [])

  const scrollToPage = useCallback((pageNumber: number) => {
    const viewer = viewerRef.current
    if (!viewer) return

    const canvas = viewer.querySelector<HTMLCanvasElement>(`canvas[data-page="${pageNumber}"]`)
    if (canvas) {
      finishProgrammaticScroll()
      isProgrammaticScrollRef.current = true
      const handleScrollEnd = () => {
        finishProgrammaticScroll()
      }
      scrollEndHandlerRef.current = handleScrollEnd
      viewer.addEventListener('scrollend', handleScrollEnd, { once: true })
      programmaticScrollTimeoutRef.current = setTimeout(() => {
        finishProgrammaticScroll()
      }, 600)

      canvas.scrollIntoView({ behavior: 'smooth', block: 'start' })
      setCurrentPage(pageNumber)
      setPageInput(String(pageNumber))
    }
  }, [finishProgrammaticScroll])

  // Navigation handlers
  const handlePrevPage = useCallback(() => {
    if (currentPage > 1) scrollToPage(currentPage - 1)
  }, [currentPage, scrollToPage])

  const handleNextPage = useCallback(() => {
    if (currentPage < pageCount) scrollToPage(currentPage + 1)
  }, [currentPage, pageCount, scrollToPage])

  const handlePageInputKeyDown = useCallback((e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      let page = parseInt(pageInput, 10)
      if (isNaN(page)) page = currentPage
      if (page < 1) page = 1
      if (page > pageCount) page = pageCount
      scrollToPage(page)
      setPageInput(String(page))
    }
  }, [pageInput, currentPage, pageCount, scrollToPage])

  // Zoom handlers
  const handleZoomIn = useCallback(() => {
    setZoomMode(prev => {
      const current = typeof prev === 'number' ? prev : actualZoom
      return Math.min(3, current + 0.25)
    })
  }, [actualZoom])

  const handleZoomOut = useCallback(() => {
    setZoomMode(prev => {
      const current = typeof prev === 'number' ? prev : actualZoom
      return Math.max(0.25, current - 0.25)
    })
  }, [actualZoom])

  // Keyboard shortcuts
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (!viewerRef.current) return
      
      // Only handle if viewer is focused or body is focused
      if (document.activeElement !== document.body && 
          !viewerRef.current.contains(document.activeElement) && 
          !containerRef.current?.contains(document.activeElement)) {
        return
      }

      if (e.ctrlKey) {
        if (e.key === '=' || e.key === '+') {
          e.preventDefault()
          handleZoomIn()
        } else if (e.key === '-') {
          e.preventDefault()
          handleZoomOut()
        } else if (e.key === '0') {
          e.preventDefault()
          setZoomMode(1)
        }
      } else {
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
          scrollToPage(pageCount)
        }
      }
    }

    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [handleNextPage, handlePrevPage, scrollToPage, pageCount, handleZoomIn, handleZoomOut])

  // Initial load
  useEffect(() => {
    setErrorMessage(null)
    setNavItems([])
    setStats(null)

    if (file.kind !== 'binary') {
      setErrorMessage('PdfViewer expected a binary file.')
      return
    }

    let cancelled = false
    let destroyDocument: (() => Promise<void>) | null = null
    let destroyLoadingTask: (() => Promise<void>) | null = null

    void (async () => {
      try {
        const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')

        if (cancelled) return

        pdfjs.GlobalWorkerOptions.workerSrc = new URL(
          'pdfjs-dist/legacy/build/pdf.worker.min.mjs',
          import.meta.url,
        ).toString()

        const nextLoadingTask = pdfjs.getDocument({ data: file.content.slice(0) })
        destroyLoadingTask = () => nextLoadingTask.destroy()

        const pdf = await nextLoadingTask.promise
        destroyDocument = () => pdf.destroy()

        if (cancelled) {
          await pdf.destroy()
          return
        }

        setPdfDoc(pdf)
        setPageCount(pdf.numPages)
        setStats({ kind: 'pdf', page: 1, pageCount: pdf.numPages })

        const outline = (await pdf.getOutline()) as ReadonlyArray<PdfOutlineNode> | null

        if (cancelled) return

        const navItems =
          outline && outline.length > 0
            ? await buildOutlineNavItems(
                outline,
                scrollToPage,
                pdf.getDestination.bind(pdf),
                pdf.getPageIndex.bind(pdf),
              )
            : []

        if (cancelled) return

        setNavItems(
          navItems.length > 0
            ? navItems
            : buildFallbackNavItems(pdf.numPages, scrollToPage),
        )
      } catch (err) {
        if (!cancelled) {
          setNavItems([])
          setStats(null)
          setErrorMessage(
            err instanceof Error ? err.message : 'Failed to render PDF.',
          )
        }
      }
    })()

    return () => {
      cancelled = true
      setPdfDoc(null)
      if (destroyDocument) void destroyDocument()
      if (destroyLoadingTask) void destroyLoadingTask()
    }
  }, [file, setNavItems, setStats, scrollToPage])

  // Update stats on page change
  useEffect(() => {
    if (pageCount > 0) {
      setStats({ kind: 'pdf', page: currentPage, pageCount })
    }
  }, [currentPage, pageCount, setStats])

  // Render pages
  useEffect(() => {
    const viewer = viewerRef.current
    const tasksMap = renderTasksRef.current
    if (!pdfDoc || !viewer) return

    let cancelled = false

    const renderPages = async () => {
      // Clean up previous renders
      tasksMap.forEach(task => task.cancel())
      tasksMap.clear()
      viewer.innerHTML = ''
      
      // Setup intersection observer
      if (intersectionObserverRef.current) {
        intersectionObserverRef.current.disconnect()
      }
      
      intersectionObserverRef.current = new IntersectionObserver((entries) => {
        if (isProgrammaticScrollRef.current) {
          return
        }

        // Find the page taking up the most screen space
        let maxRatio = 0
        let visiblePage = currentPageRef.current
        
        entries.forEach(entry => {
          if (entry.isIntersecting && entry.intersectionRatio > maxRatio) {
            maxRatio = entry.intersectionRatio
            const pageNum = parseInt((entry.target as HTMLCanvasElement).dataset.page || '1', 10)
            visiblePage = pageNum
          }
        })
        
        if (maxRatio > 0 && visiblePage !== currentPageRef.current) {
          setCurrentPage(visiblePage)
          setPageInput(String(visiblePage))
        }
      }, { threshold: [0.1, 0.5, 0.9] })

      // Calculate base scale
      let scale = 1
      
      // Get a sample page to measure if we need fit-width/fit-page
      if (zoomMode === 'fit-width' || zoomMode === 'fit-page') {
        const page1 = await pdfDoc.getPage(1)
        const viewport1 = page1.getViewport({ scale: 1.0 })
        const containerWidth = viewer.clientWidth - 48 // 24px padding on sides
        const containerHeight = viewer.clientHeight - 48
        
        if (zoomMode === 'fit-width') {
          scale = containerWidth / viewport1.width
        } else {
          scale = Math.min(containerWidth / viewport1.width, containerHeight / viewport1.height)
        }
        
        // Don't go crazy big or small
        scale = Math.max(0.25, Math.min(3, scale))
        setActualZoom(scale)
      } else {
        scale = zoomMode as number
        setActualZoom(scale)
      }

      for (let pageNumber = 1; pageNumber <= pageCount; pageNumber += 1) {
        if (cancelled) return

        try {
          const page = await pdfDoc.getPage(pageNumber)
          if (cancelled) return

          const viewport = page.getViewport({ scale })
          const canvas = document.createElement('canvas')
          canvas.className = 'pdf-viewer__page'
          canvas.dataset.page = String(pageNumber)
          
          // Use devicePixelRatio for crisp rendering
          const dpr = window.devicePixelRatio || 1
          canvas.width = Math.ceil(viewport.width * dpr)
          canvas.height = Math.ceil(viewport.height * dpr)
          canvas.style.width = `${Math.ceil(viewport.width)}px`
          canvas.style.height = `${Math.ceil(viewport.height)}px`

          const canvasContext = canvas.getContext('2d')
          if (!canvasContext) {
            throw new Error(`Failed to create a 2D canvas context for page ${pageNumber}.`)
          }
          
          canvasContext.scale(dpr, dpr)

          viewer.appendChild(canvas)
          intersectionObserverRef.current.observe(canvas)
          
          const renderContext = { canvas, canvasContext, viewport }
          const renderTask = page.render(renderContext)
          tasksMap.set(pageNumber, renderTask as { cancel: () => void })
          
          await renderTask.promise
          page.cleanup()
        } catch (err: unknown) {
          if (err instanceof Error && err.name !== 'RenderingCancelledException') {
            console.error(`Error rendering page ${pageNumber}`, err)
          }
        }
      }
    }

    void renderPages()

    return () => {
      cancelled = true
      tasksMap.forEach(task => task.cancel())
      finishProgrammaticScroll()
      if (intersectionObserverRef.current) {
        intersectionObserverRef.current.disconnect()
      }
    }
  }, [pdfDoc, pageCount, zoomMode, finishProgrammaticScroll])

  // Handle ResizeObserver for fit modes
  useEffect(() => {
    if ((zoomMode !== 'fit-width' && zoomMode !== 'fit-page') || !viewerRef.current) return
    
    let timeoutId: ReturnType<typeof setTimeout>
    const observer = new ResizeObserver(() => {
      // Debounce re-render on resize
      clearTimeout(timeoutId)
      timeoutId = setTimeout(() => {
        // Trigger re-render by replacing state with same value to cause effect run
        setZoomMode(mode => mode === 'fit-width' ? 'fit-width' : 'fit-page')
      }, 200)
    })
    
    observer.observe(viewerRef.current)
    return () => {
      clearTimeout(timeoutId)
      observer.disconnect()
    }
  }, [zoomMode])

  if (errorMessage !== null) {
    return <div className="pdf-viewer__error">Failed to render PDF: {errorMessage}</div>
  }

  return (
    <div className="pdf-viewer-container" ref={containerRef} tabIndex={-1}>
      <div className="pdf-viewer__toolbar">
        <div className="pdf-viewer__toolbar-group">
          <button 
            className="pdf-viewer__toolbar-button" 
            onClick={handlePrevPage}
            disabled={currentPage <= 1}
            title="Previous Page (Up / PageUp)"
          >
            <ChevronUp size={18} />
          </button>
          <div className="pdf-viewer__page-indicator">
            <input 
              className="pdf-viewer__page-input"
              value={pageInput}
              onChange={e => setPageInput(e.target.value)}
              onKeyDown={handlePageInputKeyDown}
              onBlur={() => setPageInput(String(currentPage))}
            />
            <span>/ {pageCount || '?'}</span>
          </div>
          <button 
            className="pdf-viewer__toolbar-button" 
            onClick={handleNextPage}
            disabled={currentPage >= pageCount}
            title="Next Page (Down / PageDown)"
          >
            <ChevronDown size={18} />
          </button>
        </div>
        
        <div className="pdf-viewer__toolbar-group" style={{ marginLeft: 'auto' }}>
          <button 
            className="pdf-viewer__toolbar-button" 
            onClick={handleZoomOut}
            title="Zoom Out (Ctrl+-)"
          >
            <ZoomOut size={18} />
          </button>
          
          <div className="pdf-viewer__zoom-menu-container">
            <button 
              className="pdf-viewer__zoom-button"
              onClick={() => setIsZoomMenuOpen(!isZoomMenuOpen)}
            >
              {Math.round(actualZoom * 100)}%
              <ChevronDownSmall size={14} />
            </button>
            
            {isZoomMenuOpen && (
              <div className="pdf-viewer__zoom-menu">
                {ZOOM_PRESETS.map(preset => (
                  <button
                    key={preset}
                    className={`pdf-viewer__zoom-menu-item ${zoomMode === preset ? 'pdf-viewer__zoom-menu-item--active' : ''}`}
                    onClick={() => {
                      setZoomMode(preset)
                      setIsZoomMenuOpen(false)
                    }}
                  >
                    {Math.round(preset * 100)}%
                  </button>
                ))}
                <hr style={{ border: 'none', borderTop: '1px solid var(--border-primary)', margin: '4px 0' }} />
                <button
                  className={`pdf-viewer__zoom-menu-item ${zoomMode === 'fit-width' ? 'pdf-viewer__zoom-menu-item--active' : ''}`}
                  onClick={() => {
                    setZoomMode('fit-width')
                    setIsZoomMenuOpen(false)
                  }}
                >
                  Fit Width
                </button>
                <button
                  className={`pdf-viewer__zoom-menu-item ${zoomMode === 'fit-page' ? 'pdf-viewer__zoom-menu-item--active' : ''}`}
                  onClick={() => {
                    setZoomMode('fit-page')
                    setIsZoomMenuOpen(false)
                  }}
                >
                  Fit Page
                </button>
              </div>
            )}
          </div>

          <button 
            className="pdf-viewer__toolbar-button" 
            onClick={handleZoomIn}
            title="Zoom In (Ctrl++)"
          >
            <ZoomIn size={18} />
          </button>
        </div>
      </div>
      
      <div ref={viewerRef} className="pdf-viewer" />
    </div>
  )
}

export const PdfViewer = memo(PdfViewerBase)
