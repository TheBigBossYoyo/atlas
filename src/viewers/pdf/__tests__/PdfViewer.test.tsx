import { act, render, screen, waitFor } from '@testing-library/react'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { PDFDocument } from 'pdf-lib'

import { ShortcutManagerProvider } from '../../../hooks/ShortcutManagerProvider'
import { ViewerProvider } from '../../shared/ViewerContext'
import { useNavItems, useViewerIsDirty, useViewerSave } from '../../shared/useViewerContext'
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

const { getDocumentMock, destroyMock, pageRotateState, renderCalls, annotationsState, saveDocumentMock, metadataState } = vi.hoisted(() => {
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
  // Page 1's annotations; empty unless a test puts form widgets there.
  const annotationsState: { current: unknown[] } = { current: [] }
  const saveDocumentMock = vi.fn(() => Promise.resolve(new Uint8Array([9, 9, 9])))
  // What `getMetadata().info` reports: encryption and signatures gate editing.
  const metadataState: { current: Record<string, unknown> } = { current: {} }

  const makePageMock = vi.fn(() => ({
    rotate: pageRotateState.current,
    getViewport: (opts?: { scale?: number; rotation?: number }) => makeViewport(opts),
    render: (params: { annotationMode?: number }) => {
      renderCalls.push(params)
      return { promise: Promise.resolve(), cancel: vi.fn() }
    },
    getTextContent: () => Promise.resolve({ items: [], styles: {}, lang: null }),
    getAnnotations: () => Promise.resolve(annotationsState.current),
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
      annotationStorage: { setValue: vi.fn(), getValue: (_key: string, fallback: unknown) => fallback },
      saveDocument: saveDocumentMock,
      getMetadata: () => Promise.resolve({ info: metadataState.current }),
    }
    return {
      promise: Promise.resolve(pdf),
      destroy: destroyMock,
      onPassword: undefined as unknown,
    }
  })

  return { getDocumentMock, makePageMock, destroyMock, pageRotateState, renderCalls, annotationsState, saveDocumentMock, metadataState }
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

/** Exposes the shell's view of the document session: dirty flag and Save. */
function SessionProbe() {
  const isDirty = useViewerIsDirty()
  const save = useViewerSave()
  return (
    <button type="button" data-testid="shell-save" data-dirty={String(isDirty)} onClick={() => void save()}>
      shell save
    </button>
  )
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
    annotationsState.current = []
    metadataState.current = {}
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
      <ShortcutManagerProvider>
        <ViewerProvider filePath="/fixtures/sample.pdf">
          <PdfViewer file={binaryFile()} />
        </ViewerProvider>
      </ShortcutManagerProvider>,
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

  it('a typed form value marks the document dirty, and Save writes the form through pdf.js to the same file', async () => {
    annotationsState.current = [
      { id: 'name', subtype: 'Widget', fieldType: 'Tx', fieldName: 'name', fieldValue: '', rect: [0, 0, 50, 10] },
    ]
    const saveBinaryFile = vi.fn(() => Promise.resolve({ saved: true, path: '/fixtures/sample.pdf' }))
    vi.stubGlobal('electronAPI', { saveBinaryFile })

    render(
      <ViewerProvider filePath="/fixtures/sample.pdf">
        <SessionProbe />
        <PdfViewer file={binaryFile()} />
      </ViewerProvider>,
    )
    const field = (await screen.findAllByRole('textbox', { name: 'name' }))[0]
    expect(screen.getByTestId('shell-save')).toHaveAttribute('data-dirty', 'false')

    await act(async () => {
      ;(field as HTMLInputElement).value = 'Youssef'
      field.dispatchEvent(new Event('input', { bubbles: true }))
    })
    await waitFor(() => expect(screen.getByTestId('shell-save')).toHaveAttribute('data-dirty', 'true'))

    await act(async () => {
      screen.getByTestId('shell-save').click()
    })
    await waitFor(() => expect(saveBinaryFile).toHaveBeenCalledTimes(1))
    expect(saveDocumentMock).toHaveBeenCalledTimes(1)
    expect(saveBinaryFile).toHaveBeenCalledWith(
      expect.objectContaining({
        content: new Uint8Array([9, 9, 9]),
        existingPath: '/fixtures/sample.pdf',
        filters: [{ name: 'PDF documents', extensions: ['pdf'] }],
      }),
    )
    await waitFor(() => expect(screen.getByTestId('shell-save')).toHaveAttribute('data-dirty', 'false'))
  })

  it('an unedited PDF saves its original bytes without asking pdf.js to rewrite it', async () => {
    const saveBinaryFile = vi.fn(() => Promise.resolve({ saved: true, path: '/fixtures/sample.pdf' }))
    vi.stubGlobal('electronAPI', { saveBinaryFile })
    render(
      <ViewerProvider filePath="/fixtures/sample.pdf">
        <SessionProbe />
        <PdfViewer file={binaryFile()} />
      </ViewerProvider>,
    )
    await waitFor(() => expect(screen.getByText('/ 2')).toBeInTheDocument())
    await act(async () => {
      screen.getByTestId('shell-save').click()
    })
    await waitFor(() => expect(saveBinaryFile).toHaveBeenCalledTimes(1))
    expect(saveDocumentMock).not.toHaveBeenCalled()
    expect(saveBinaryFile).toHaveBeenCalledWith(expect.objectContaining({ content: new Uint8Array([1, 2, 3]) }))
  })

  it('deleting a page in edit mode removes it from the view, the page count and the sidebar; undo brings it back', async () => {
    render(
      <ShortcutManagerProvider>
        <ViewerProvider filePath="/fixtures/sample.pdf">
          <NavItemCount />
          <SessionProbe />
          <PdfViewer file={binaryFile()} />
        </ViewerProvider>
      </ShortcutManagerProvider>,
    )
    await waitFor(() => expect(screen.getByText('/ 2')).toBeInTheDocument())

    await act(async () => {
      screen.getByRole('button', { name: 'Edit' }).click()
    })
    await act(async () => {
      screen.getByRole('button', { name: 'Delete this page' }).click()
    })

    await waitFor(() => expect(screen.getByText('/ 1')).toBeInTheDocument())
    expect(document.querySelectorAll('.pdf-viewer__page-wrapper')).toHaveLength(1)
    await waitFor(() => expect(screen.getByTestId('nav-item-count')).toHaveTextContent('1'))
    expect(screen.getByTestId('shell-save')).toHaveAttribute('data-dirty', 'true')
    // The last remaining page cannot be deleted.
    expect(screen.getByRole('button', { name: 'Delete this page' })).toBeDisabled()

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true }))
    })
    await waitFor(() => expect(screen.getByText('/ 2')).toBeInTheDocument())
    expect(screen.getByTestId('shell-save')).toHaveAttribute('data-dirty', 'false')
  })

  describe('edit mode', () => {
    /**
     * A real two-page file to match the mocked pdf.js document: a save with
     * page or annotation edits rewrites these bytes with the real pdf-lib.
     */
    async function twoPagePdf(): Promise<LoadedFile> {
      const doc = await PDFDocument.create()
      doc.addPage([100, 200])
      doc.addPage([100, 200])
      const bytes = await doc.save()
      return {
        kind: 'binary',
        content: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
        path: '/fixtures/sample.pdf',
        format: 'pdf',
      }
    }

    async function openEditMode() {
      const file = await twoPagePdf()
      render(
        <ShortcutManagerProvider>
          <ViewerProvider filePath="/fixtures/sample.pdf">
            <SessionProbe />
            <PdfViewer file={file} />
          </ViewerProvider>
        </ShortcutManagerProvider>,
      )
      await waitFor(() => expect(screen.getByText('/ 2')).toBeInTheDocument())
      await act(async () => {
        screen.getByRole('button', { name: 'Edit' }).click()
      })
    }

    /** The annotation tools wait for pdf-lib's font metrics to load. */
    async function waitForTools() {
      await waitFor(() => expect(screen.getByRole('radio', { name: 'Add a text box' })).toBeEnabled(), { timeout: 10_000 })
    }

    it('switches tools, shows each tool its own controls, and closes again', async () => {
      await openEditMode()
      await waitForTools()
      expect(screen.getByRole('radio', { name: 'Select, move and delete edits' })).toHaveAttribute('aria-checked', 'true')
      expect(screen.getByText('Size')).toBeInTheDocument()

      await act(async () => {
        screen.getByRole('radio', { name: 'Draw' }).click()
      })
      expect(screen.getByText('Width')).toBeInTheDocument()
      expect(screen.getByRole('radio', { name: 'Red' })).toHaveAttribute('aria-checked', 'true')
      await act(async () => {
        screen.getByRole('radio', { name: 'Blue' }).click()
      })
      expect(screen.getByRole('radio', { name: 'Blue' })).toHaveAttribute('aria-checked', 'true')

      await act(async () => {
        screen.getByRole('radio', { name: /^Edit text/ }).click()
      })
      // Colour, size and typeface apply to text added on empty space.
      expect(screen.getByRole('radio', { name: 'Blue' })).toBeInTheDocument()
      expect(screen.getByText('Font')).toBeInTheDocument()
      expect(screen.getByText(/Click a paragraph to edit it in place/)).toBeInTheDocument()

      await act(async () => {
        screen.getByRole('button', { name: 'Edit' }).click()
      })
      expect(screen.queryByRole('toolbar', { name: 'PDF editing' })).toBeNull()
    })

    it('Save As always asks where, and a page rotation turns the page on screen', async () => {
      const saveBinaryFile = vi.fn(() => Promise.resolve({ saved: false }))
      vi.stubGlobal('electronAPI', { saveBinaryFile })
      await openEditMode()

      const sizeOf = () => {
        const wrapper = document.querySelector('.pdf-viewer__page-wrapper[data-page="1"]') as HTMLElement
        return { width: parseFloat(wrapper.style.width), height: parseFloat(wrapper.style.height) }
      }
      await waitFor(() => expect(sizeOf().height).toBeGreaterThan(sizeOf().width))
      await act(async () => {
        screen.getByRole('button', { name: 'Rotate this page right (saved)' }).click()
      })
      await waitFor(() => expect(sizeOf().width).toBeGreaterThan(sizeOf().height))

      await act(async () => {
        screen.getByRole('button', { name: 'Save as a new PDF (Ctrl+Shift+S)' }).click()
      })
      await waitFor(() => expect(saveBinaryFile).toHaveBeenCalledTimes(1))
      const request = (saveBinaryFile.mock.calls[0] as unknown[])[0] as Record<string, unknown>
      expect(request).not.toHaveProperty('existingPath')
      expect(request.suggestedName).toBe('sample.pdf')
      // A cancelled dialog leaves the document dirty and says nothing.
      expect(screen.getByTestId('shell-save')).toHaveAttribute('data-dirty', 'true')
      expect(screen.queryByRole('alert')).toBeNull()
    })

    it('shows the write error when the save fails, and keeps the changes', async () => {
      vi.stubGlobal('electronAPI', {
        saveBinaryFile: vi.fn(() => Promise.resolve({ saved: false, error: 'The file is open in another program.' })),
      })
      await openEditMode()
      await act(async () => {
        screen.getByRole('button', { name: 'Move this page down' }).click()
      })
      await act(async () => {
        screen.getByRole('button', { name: /^Save$/ }).click()
      })
      expect(await screen.findByRole('alert')).toHaveTextContent('The file is open in another program.')
      expect(screen.getByTestId('shell-save')).toHaveAttribute('data-dirty', 'true')
    })

    it('an encrypted PDF offers form filling only: tools and page operations are disabled, with the reason shown', async () => {
      metadataState.current = { EncryptFilterName: 'Standard' }
      await openEditMode()
      await waitFor(() => expect(screen.getByText(/This PDF is encrypted/)).toBeInTheDocument())
      expect(screen.getByRole('radio', { name: 'Add a text box' })).toBeDisabled()
      expect(screen.getByRole('button', { name: 'Delete this page' })).toBeDisabled()
      expect(screen.getByRole('button', { name: 'Rotate this page right (saved)' })).toBeDisabled()
    })

    it('warns before editing a digitally signed PDF', async () => {
      metadataState.current = { IsSignaturesPresent: true }
      await openEditMode()
      expect(await screen.findByText(/digitally signed/)).toBeInTheDocument()
    })
  })
})
