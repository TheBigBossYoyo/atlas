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

const { getDocumentMock } = vi.hoisted(() => {
  function makeViewport({ scale = 1 }: { scale?: number; rotation?: number } = {}) {
    const width = 100 * scale
    const height = 200 * scale
    return {
      width,
      height,
      scale,
      rotation: 0,
      rawDims: { pageWidth: 100, pageHeight: 200, pageX: 0, pageY: 0 },
      convertToViewportPoint: (x: number, y: number) => [x, y],
    }
  }

  const makePageMock = vi.fn(() => ({
    getViewport: (opts?: { scale?: number; rotation?: number }) => makeViewport(opts),
    render: () => ({ promise: Promise.resolve(), cancel: vi.fn() }),
    getTextContent: () => Promise.resolve({ items: [], styles: {}, lang: null }),
    getAnnotations: () => Promise.resolve([]),
    cleanup: () => {},
  }))

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
      destroy: () => Promise.resolve(),
      onPassword: undefined as unknown,
    }
  })

  return { getDocumentMock, makePageMock }
})

vi.mock('pdfjs-dist/legacy/build/pdf.mjs', () => ({
  GlobalWorkerOptions: { workerSrc: '' },
  PasswordResponses: { NEED_PASSWORD: 1, INCORRECT_PASSWORD: 2 },
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
})
