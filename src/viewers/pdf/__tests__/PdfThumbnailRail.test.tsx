/**
 * P4.7 — PdfThumbnailRail coverage sweep. The rail doesn't import
 * `pdfjs-dist` itself (only the `PdfDocument`/`PdfPage` *types*), so a
 * plain fake `pdfDoc` — matching just the `getPage().getViewport()/render()`
 * shape this component actually calls, same idea as `PdfViewer.test.tsx`'s
 * `makePageMock` — is enough; no module mock needed.
 *
 * `IntersectionObserver` is faked with a controllable stand-in (rather than
 * the always-immediately-visible fake in `PdfViewer.test.tsx`) so tests can
 * assert the lazy-rendering behavior itself: only pages within the
 * visible-window ± overscan should ever get a live canvas render.
 */
import { act, fireEvent, render, screen, waitFor } from '@testing-library/react'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { PdfThumbnailRail } from '../PdfThumbnailRail'
import type { PdfDocument } from '../types'

class ControllableIntersectionObserver implements IntersectionObserver {
  static instances: ControllableIntersectionObserver[] = []
  readonly root = null
  readonly rootMargin = ''
  readonly scrollMargin = ''
  readonly thresholds: ReadonlyArray<number> = []
  readonly observed = new Set<Element>()
  private readonly callback: IntersectionObserverCallback

  constructor(callback: IntersectionObserverCallback) {
    this.callback = callback
    ControllableIntersectionObserver.instances.push(this)
  }

  observe(target: Element): void {
    this.observed.add(target)
  }
  unobserve(target: Element): void {
    this.observed.delete(target)
  }
  disconnect(): void {
    this.observed.clear()
  }
  takeRecords(): IntersectionObserverEntry[] {
    return []
  }

  /** Test helper: deliver an intersection-change batch synchronously. */
  trigger(entries: ReadonlyArray<{ target: Element; isIntersecting: boolean }>): void {
    this.callback(
      entries.map(
        (entry) =>
          ({
            target: entry.target,
            isIntersecting: entry.isIntersecting,
            intersectionRatio: entry.isIntersecting ? 1 : 0,
          }) as IntersectionObserverEntry,
      ),
      this,
    )
  }
}

function makePdfDoc(renderedPages: number[], rotateOf: (pageNumber: number) => number = () => 0): PdfDocument {
  return {
    getPage: (pageNumber: number) =>
      Promise.resolve({
        rotate: rotateOf(pageNumber),
        getViewport: ({ scale = 1, rotation = 0 }: { scale?: number; rotation?: number } = {}) => {
          const sideways = rotation === 90 || rotation === 270
          return {
            width: (sideways ? 200 : 100) * scale,
            height: (sideways ? 100 : 200) * scale,
          }
        },
        render: ({ canvas }: { canvas: HTMLCanvasElement }) => {
          renderedPages.push(pageNumber)
          return { promise: Promise.resolve(), cancel: vi.fn(), canvas }
        },
      }),
  } as unknown as PdfDocument
}

function pageNode(pageNumber: number): HTMLElement {
  const el = document.querySelector(`[data-page="${pageNumber}"]`)
  if (!el) throw new Error(`no thumbnail node for page ${pageNumber}`)
  return el as HTMLElement
}

function latestObserver(): ControllableIntersectionObserver {
  const instance = ControllableIntersectionObserver.instances.at(-1)
  if (!instance) throw new Error('no IntersectionObserver instance created')
  return instance
}

describe('PdfThumbnailRail', () => {
  beforeEach(() => {
    ControllableIntersectionObserver.instances.length = 0
    vi.stubGlobal('IntersectionObserver', ControllableIntersectionObserver)
    HTMLCanvasElement.prototype.getContext = vi.fn(() => ({
      scale: vi.fn(),
      setTransform: vi.fn(),
    })) as unknown as typeof HTMLCanvasElement.prototype.getContext
  })

  afterEach(() => {
    vi.unstubAllGlobals()
    vi.restoreAllMocks()
  })

  it('renders nothing when closed', () => {
    const { container } = render(
      <PdfThumbnailRail
        isOpen={false}
        pdfDoc={makePdfDoc([])}
        pageCount={5}
        currentPage={1}
        rotation={0}
        onSelectPage={vi.fn()}
      />,
    )

    expect(container).toBeEmptyDOMElement()
  })

  it('renders one thumbnail button per page, labelled and marked current', () => {
    render(
      <PdfThumbnailRail
        isOpen
        pdfDoc={makePdfDoc([])}
        pageCount={5}
        currentPage={3}
        rotation={0}
        onSelectPage={vi.fn()}
      />,
    )

    const buttons = screen.getAllByRole('button')
    expect(buttons).toHaveLength(5)
    expect(screen.getByLabelText('Go to page 1')).toBeInTheDocument()
    expect(screen.getByLabelText('Go to page 5')).toBeInTheDocument()

    expect(screen.getByLabelText('Go to page 3')).toHaveAttribute('aria-current', 'true')
    expect(screen.getByLabelText('Go to page 1')).toHaveAttribute('aria-current', 'false')
  })

  it('calls onSelectPage with the clicked page number', () => {
    const onSelectPage = vi.fn()
    render(
      <PdfThumbnailRail
        isOpen
        pdfDoc={makePdfDoc([])}
        pageCount={5}
        currentPage={1}
        rotation={0}
        onSelectPage={onSelectPage}
      />,
    )

    fireEvent.click(screen.getByLabelText('Go to page 4'))
    expect(onSelectPage).toHaveBeenCalledWith(4)
    expect(onSelectPage).toHaveBeenCalledTimes(1)
  })

  it('only renders a live canvas for pages within the visible window (± overscan), not the whole document', async () => {
    // THUMBNAIL_OVERSCAN is 6 (see PdfThumbnailRail.tsx) — with only page 10
    // reported intersecting out of 20 total pages, the active window is
    // [4, 16]; pages 1-3 and 17-20 must never call pdfDoc.getPage()/render().
    const renderedPages: number[] = []
    render(
      <PdfThumbnailRail
        isOpen
        pdfDoc={makePdfDoc(renderedPages)}
        pageCount={20}
        currentPage={1}
        rotation={0}
        onSelectPage={vi.fn()}
      />,
    )

    await act(async () => {
      latestObserver().trigger([{ target: pageNode(10), isIntersecting: true }])
    })

    await waitFor(() => expect(renderedPages.length).toBe(13))
    expect(new Set(renderedPages)).toEqual(new Set(Array.from({ length: 13 }, (_, i) => i + 4)))

    // Canvases for rendered pages are actually sized; out-of-window ones
    // stay at their cleared 0x0 size.
    const canvas10 = pageNode(10).querySelector('canvas') as HTMLCanvasElement
    const canvas1 = pageNode(1).querySelector('canvas') as HTMLCanvasElement
    await waitFor(() => expect(canvas10.width).toBeGreaterThan(0))
    expect(canvas1.width).toBe(0)
  })

  it('releases a thumbnail canvas (resets it to 0x0) once it scrolls back out of the active window', async () => {
    const renderedPages: number[] = []
    render(
      <PdfThumbnailRail
        isOpen
        pdfDoc={makePdfDoc(renderedPages)}
        pageCount={20}
        currentPage={1}
        rotation={0}
        onSelectPage={vi.fn()}
      />,
    )

    await act(async () => {
      latestObserver().trigger([{ target: pageNode(10), isIntersecting: true }])
    })
    const canvas10 = pageNode(10).querySelector('canvas') as HTMLCanvasElement
    await waitFor(() => expect(canvas10.width).toBeGreaterThan(0))

    await act(async () => {
      latestObserver().trigger([{ target: pageNode(10), isIntersecting: false }])
    })

    await waitFor(() => expect(canvas10.width).toBe(0))
  })

  it("renders a thumbnail at the page's TOTAL rotation (intrinsic /Rotate plus the viewer's rotation state)", async () => {
    // Mirrors PdfViewer.test.tsx's equivalent regression test for the main
    // page canvas — the thumbnail rail has its own getViewport call site and
    // must combine the two rotations the same way (rotation.ts's
    // combineRotation), or a scanned/rotated page's thumbnail renders
    // upright while the real page renders sideways.
    const renderedPages: number[] = []
    render(
      <PdfThumbnailRail
        isOpen
        pdfDoc={makePdfDoc(renderedPages, () => 90)}
        pageCount={1}
        currentPage={1}
        rotation={0}
        onSelectPage={vi.fn()}
      />,
    )

    await act(async () => {
      latestObserver().trigger([{ target: pageNode(1), isIntersecting: true }])
    })

    const canvas1 = pageNode(1).querySelector('canvas') as HTMLCanvasElement
    // Raw page is 100x200 (portrait); at a combined 90deg rotation pdf.js's
    // viewport swaps width/height, so the rendered thumbnail must come out
    // landscape (wider than tall).
    await waitFor(() => expect(canvas1.width).toBeGreaterThan(0))
    expect(canvas1.width).toBeGreaterThan(canvas1.height)
  })

  it('unobserves a thumbnail node when the rail closes and stops observing it', () => {
    const { rerender } = render(
      <PdfThumbnailRail
        isOpen
        pdfDoc={makePdfDoc([])}
        pageCount={3}
        currentPage={1}
        rotation={0}
        onSelectPage={vi.fn()}
      />,
    )

    const observer = latestObserver()
    expect(observer.observed.size).toBe(3)

    rerender(
      <PdfThumbnailRail
        isOpen={false}
        pdfDoc={makePdfDoc([])}
        pageCount={3}
        currentPage={1}
        rotation={0}
        onSelectPage={vi.fn()}
      />,
    )

    // Closing unmounts the rail entirely, which disconnects its observer.
    expect(observer.observed.size).toBe(0)
  })
})
