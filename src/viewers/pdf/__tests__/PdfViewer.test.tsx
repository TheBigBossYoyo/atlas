import { act, render, screen, waitFor } from '@testing-library/react'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { ViewerProvider } from '../../shared/ViewerContext'
import { useNavItems } from '../../shared/useViewerContext'
import type { LoadedFile } from '../../../formats/types'
import { PdfViewer } from '../PdfViewer'

// Regression coverage for the Wave 3-P rewrite: a real pdfjs-dist runtime
// needs a Worker + canvas 2D rendering, neither of which jsdom provides, so
// the whole `pdfjs-dist/legacy/build/pdf.mjs` module is mocked with a
// minimal fake matching the shape PdfViewer/PdfPage actually call. This
// exercises the component's own orchestration (load -> render pages ->
// fallback nav items) rather than pdf.js internals, which the pure-module
// unit tests (outline/geometry/virtualization/search/annotations) already
// cover in isolation.

const { getDocumentMock, destroyMock, pageRotateState, renderCalls } = vi.hoisted(() => {
  // `sideways` mimics real pdf.js: getViewport({rotation}) swaps
  // width/height for a 90/270 TOTAL rotation. The page's raw (unrotated)
  // size here is a fixed 100x200 (portrait) — `pageRotateState.current`
  // lets individual tests simulate a page's own intrinsic /Rotate entry
  // (PDFPageProxy.rotate), independent of the viewer's own rotation state,
  // to prove the two get combined rather than one silently overriding the
  // other (see rotation.ts's `combineRotation`).
  function makeViewport({ scale = 1, rotation = 0 }: { scale?: number; rotation?: number } = {}) {
    const sideways = rotation === 90 || rotation === 270
    const width = (sideways ? 200 : 100) * scale
    const height = (sideways ? 100 : 200) * scale
    return {
      width,
      height,
      scale,
      rotation,
      rawDims: { pageWidth: 100, pageHeight: 200, pageX: 0, pageY: 0 },
      convertToViewportPoint: (x: number, y: number) => [x, y],
    }
  }

  const pageRotateState = { current: 0 }
  const renderCalls: Array<{ annotationMode?: number }> = []

  const makePageMock = vi.fn(() => ({
    rotate: pageRotateState.current,
    getViewport: (opts?: { scale?: number; rotation?: number }) => makeViewport(opts),
    render: (params: { annotationMode?: number }) => {
      renderCalls.push(params)
      return { promise: Promise.resolve(), cancel: vi.fn() }
    },
    getTextContent: () => Promise.resolve({ items: [], styles: {}, lang: null }),
    getAnnotations: () => Promise.resolve([]),
    cleanup: () => {},
  }))

  const destroyMock = vi.fn(() => Promise.resolve())

  const getDocumentMock = vi.fn(() => {
    const pdf = {
      numPages: 2,
      getPage: () => Promise.resolve(makePageMock()),
      getOutline: () => Promise.resolve(null),
      getDestination: () => Promise.resolve(null),
      getPageIndex: () => Promise.resolve(0),
      annotationStorage: { setValue: vi.fn() },
    }
    return {
      promise: Promise.resolve(pdf),
      destroy: destroyMock,
      onPassword: undefined as unknown,
    }
  })

  return { getDocumentMock, makePageMock, destroyMock, pageRotateState, renderCalls }
})

vi.mock('pdfjs-dist/legacy/build/pdf.mjs', () => ({
  GlobalWorkerOptions: { workerSrc: '' },
  PasswordResponses: { NEED_PASSWORD: 1, INCORRECT_PASSWORD: 2 },
  AnnotationMode: { DISABLE: 0, ENABLE: 1, ENABLE_FORMS: 2, ENABLE_STORAGE: 3 },
  TextLayer: class FakeTextLayer {
    render() {
      return Promise.resolve()
    }
    cancel() {}
  },
  getDocument: getDocumentMock,
}))

class FakeIntersectionObserver implements IntersectionObserver {
  readonly root = null
  readonly rootMargin = ''
  readonly scrollMargin = ''
  readonly thresholds: ReadonlyArray<number> = []
  private readonly callback: IntersectionObserverCallback

  constructor(callback: IntersectionObserverCallback) {
    this.callback = callback
  }

  observe(target: Element): void {
    // Simulate every observed page being immediately on-screen.
    void Promise.resolve().then(() => {
      this.callback(
        [
          {
            target,
            isIntersecting: true,
            intersectionRatio: 1,
          } as IntersectionObserverEntry,
        ],
        this,
      )
    })
  }
  unobserve(): void {}
  disconnect(): void {}
  takeRecords(): IntersectionObserverEntry[] {
    return []
  }
}

class FakeResizeObserver implements ResizeObserver {
  observe(): void {}
  unobserve(): void {}
  disconnect(): void {}
}

function binaryFile(): LoadedFile {
  return { kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: '/fixtures/sample.pdf', format: 'pdf' }
}

function NavItemCount() {
  const items = useNavItems()
  return <div data-testid="nav-item-count">{items.length}</div>
}

describe('PdfViewer', () => {
  beforeEach(() => {
    vi.stubGlobal('IntersectionObserver', FakeIntersectionObserver)
    vi.stubGlobal('ResizeObserver', FakeResizeObserver)
    vi.stubGlobal('matchMedia', () => ({
      matches: false,
      addEventListener: () => {},
      removeEventListener: () => {},
    }))
    HTMLCanvasElement.prototype.getContext = vi.fn(() => ({
      scale: vi.fn(),
      setTransform: vi.fn(),
    })) as unknown as typeof HTMLCanvasElement.prototype.getContext
    pageRotateState.current = 0
    renderCalls.length = 0
    // `vi.clearAllMocks()` below runs in `afterEach`, which — being
    // registered inside this `describe` block (innermost) — fires BEFORE
    // Testing Library's own auto-cleanup `afterEach` (registered at this
    // module's top-level import, outermost), per standard innermost-first
    // afterEach ordering. That means the PREVIOUS test's component actually
    // unmounts (and calls `destroyMock` via this branch's cleanup fix)
    // AFTER `vi.clearAllMocks()` already ran — so a leftover call can only
    // be cleared here, at the START of the next test.
    destroyMock.mockClear()
  })

  afterEach(() => {
    vi.unstubAllGlobals()
    vi.clearAllMocks()
  })

  it('loads a PDF, reports its page count, and falls back to a flat page-list nav when there is no outline', async () => {
    render(
      <ViewerProvider filePath="/fixtures/sample.pdf">
        <NavItemCount />
        <PdfViewer file={binaryFile()} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(screen.getByLabelText('Page number')).toHaveValue('1'))
    await waitFor(() => expect(screen.getByText('/ 2')).toBeInTheDocument())

    // No outline (getOutline resolves null) -> buildFallbackNavItems(2, ...).
    await waitFor(() => expect(screen.getByTestId('nav-item-count')).toHaveTextContent('2'))
  })

  it('renders exactly one page-wrapper per document page', async () => {
    render(
      <ViewerProvider filePath="/fixtures/sample.pdf">
        <PdfViewer file={binaryFile()} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(screen.getByText('/ 2')).toBeInTheDocument())
    await waitFor(() => {
      expect(document.querySelectorAll('.pdf-viewer__page-wrapper')).toHaveLength(2)
    })
  })

  it('shows a friendly error for a non-binary file instead of crashing', async () => {
    render(
      <ViewerProvider filePath="/fixtures/sample.md">
        <PdfViewer
          file={{ kind: 'text', content: '# hi', path: '/fixtures/sample.md', format: 'pdf' }}
        />
      </ViewerProvider>,
    )

    expect(await screen.findByText(/PdfViewer expected a binary file/)).toBeInTheDocument()
  })

  it('wraps a raw pdf.js load failure in a friendly message instead of showing it verbatim (RUN-14)', async () => {
    getDocumentMock.mockImplementationOnce(() => ({
      promise: Promise.reject(new Error('Invalid PDF structure')),
      destroy: destroyMock,
      onPassword: undefined as unknown,
    }))

    render(
      <ViewerProvider filePath="/fixtures/sample.pdf">
        <PdfViewer file={binaryFile()} />
      </ViewerProvider>,
    )

    expect(await screen.findByText(/the PDF file appears to be corrupted or malformed/)).toBeInTheDocument()
    expect(screen.queryByText('Invalid PDF structure')).not.toBeInTheDocument()
  })

  it('opens the find bar on Ctrl+F while the viewer is focused', async () => {
    render(
      <ViewerProvider filePath="/fixtures/sample.pdf">
        <PdfViewer file={binaryFile()} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(screen.getByText('/ 2')).toBeInTheDocument())

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'f', ctrlKey: true, bubbles: true }))
    })

    expect(await screen.findByPlaceholderText('Find in document…')).toBeInTheDocument()
  })

  it('destroys the pdfjs loading task (terminating its Worker) when the viewer unmounts', async () => {
    // Regression test: the document-load effect's cleanup previously only
    // called `setPdfDoc(null)` and never `destroyLoadingTaskRef.current()`,
    // so every unmount/file-switch leaked a full pdfjs-dist Worker thread
    // plus its retained document state for the life of the renderer process
    // — silently defeating the whole point of the P1 virtualization work.
    const { unmount } = render(
      <ViewerProvider filePath="/fixtures/sample.pdf">
        <PdfViewer file={binaryFile()} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(screen.getByText('/ 2')).toBeInTheDocument())
    expect(destroyMock).not.toHaveBeenCalled()

    unmount()

    await waitFor(() => expect(destroyMock).toHaveBeenCalledTimes(1))
  })

  it('destroys the previous document when the file changes, before loading the next one', async () => {
    const { rerender } = render(
      <ViewerProvider filePath="/fixtures/sample.pdf">
        <PdfViewer file={binaryFile()} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(screen.getByText('/ 2')).toBeInTheDocument())

    rerender(
      <ViewerProvider filePath="/fixtures/other.pdf">
        <PdfViewer file={{ ...binaryFile(), path: '/fixtures/other.pdf' }} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(destroyMock).toHaveBeenCalledTimes(1))
    // Let the new document's own load settle fully before the test ends —
    // otherwise a pending continuation of its async load chain can fire
    // after this test's `afterEach` has already un-stubbed the
    // IntersectionObserver/ResizeObserver globals.
    await waitFor(() => expect(screen.getByText('/ 2')).toBeInTheDocument())
  })

  it("renders a page at the TOTAL rotation (its own intrinsic /Rotate plus the viewer's rotation state)", async () => {
    // The page's own baked-in rotation (common for scanned/camera-captured
    // PDFs) must still take effect even at the viewer's default 0deg
    // rotation state — pdf.js's getViewport({rotation}) would otherwise
    // silently discard it once this viewer passes an explicit `rotation`
    // (see rotation.ts's `combineRotation`).
    pageRotateState.current = 90

    render(
      <ViewerProvider filePath="/fixtures/sample.pdf">
        <PdfViewer file={binaryFile()} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(screen.getByText('/ 2')).toBeInTheDocument())

    const canvas = await waitFor(() => {
      const el = document.querySelector(
        '.pdf-viewer__page-wrapper[data-page="1"] canvas.pdf-viewer__page',
      ) as HTMLCanvasElement | null
      if (!el || el.style.width === '') throw new Error('canvas not sized yet')
      return el
    })

    // Page's raw (unrotated) size is 100x200 (portrait). At a combined
    // total rotation of 90deg (0 from the viewer + 90 intrinsic), pdf.js's
    // viewport swaps width/height, so the rendered canvas must come out
    // landscape.
    expect(parseFloat(canvas.style.width)).toBeGreaterThan(parseFloat(canvas.style.height))
  })

  it('prints with annotationMode ENABLE_STORAGE so filled-in form values are included', async () => {
    // handlePrint previously called page.render() with no annotationMode
    // (pdf.js defaults to plain ENABLE), which paints annotations' static
    // appearance streams but never reads back the live values this
    // viewer's form-field overlay writes into pdfDoc.annotationStorage —
    // so a filled-in text field or a checked checkbox silently vanished
    // from the printed output.
    vi.stubGlobal('print', vi.fn())

    render(
      <ViewerProvider filePath="/fixtures/sample.pdf">
        <PdfViewer file={binaryFile()} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(screen.getByText('/ 2')).toBeInTheDocument())
    // Wait for both pages' own initial canvas render to finish before
    // resetting `renderCalls` — otherwise a still-in-flight main-view
    // render() call (annotationMode undefined) can land in the array right
    // alongside the print handler's own calls below and fail `every(...)`
    // for a reason that has nothing to do with print.
    await waitFor(() => {
      const canvases = document.querySelectorAll('.pdf-viewer__page-wrapper canvas.pdf-viewer__page')
      expect(canvases).toHaveLength(2)
      canvases.forEach((canvas) => expect((canvas as HTMLCanvasElement).style.width).not.toBe(''))
    })

    renderCalls.length = 0
    await act(async () => {
      screen.getByTitle('Print (Ctrl+P)').click()
    })

    await waitFor(() => expect(renderCalls.length).toBe(2))
    expect(renderCalls.every((call) => call.annotationMode === 3)).toBe(true)
  })
})
