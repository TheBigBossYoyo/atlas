import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import type { Document as DocxDocument, RunChild } from '../../docx/model'
import { ViewerProvider } from '../shared/ViewerContext'
import { useViewerIsDirty, useViewerSave } from '../shared/useViewerContext'
import { DocxViewer } from '../DocxViewer'

// ---------------------------------------------------------------------------
// Test-only helpers
// ---------------------------------------------------------------------------

/** Renders the active viewer's shared-context dirty flag as text for assertions. */
function ViewerDirtyProbe() {
  const isDirty = useViewerIsDirty()
  return <span data-testid="viewer-dirty">{String(isDirty)}</span>
}

/** Renders a button that calls the shared context's registered save(). */
function ViewerSaveProbe() {
  const save = useViewerSave()
  return (
    <button type="button" onClick={() => void save()}>
      Context Save
    </button>
  )
}

function collectRunChildren(document: DocxDocument): ReadonlyArray<RunChild> {
  const children: RunChild[] = []
  for (const section of document.sections) {
    for (const block of section.blocks) {
      if (block.kind !== 'paragraph') continue
      for (const child of block.children) {
        if (child.kind === 'run') {
          children.push(...child.children)
        } else if (child.kind === 'hyperlink') {
          for (const grandchild of child.children) {
            if (grandchild.kind === 'run') {
              children.push(...grandchild.children)
            }
          }
        }
      }
    }
  }
  return children
}

function collectText(document: DocxDocument): string {
  return collectRunChildren(document)
    .filter((child): child is Extract<RunChild, { kind: 'text' }> => child.kind === 'text')
    .map((child) => child.value)
    .join('')
}

function hasDrawing(document: DocxDocument): boolean {
  return collectRunChildren(document).some((child) => child.kind === 'drawing')
}

/** Types into Find, then Replace, and clicks "Replace" — a real edit via the
 * real (unmocked) find/replace/command pipeline, without needing a genuine
 * DOM selection (the mocked PageStack doesn't render real paragraph text). */
function replaceFirstMatch(query: string, replacement: string): void {
  const findInput = document.querySelector<HTMLInputElement>('.docx-find__input[aria-label="Find"]')
  const replaceInput = document.querySelector<HTMLInputElement>('.docx-find__input[aria-label="Replace"]')
  if (!findInput || !replaceInput) {
    throw new Error('Find/Replace inputs not found')
  }
  fireEvent.change(findInput, { target: { value: query } })
  fireEvent.change(replaceInput, { target: { value: replacement } })
  fireEvent.click(screen.getByRole('button', { name: 'Replace' }))
}

const { loadDocxMock, saveDocxMock, paginateMock } = vi.hoisted(() => ({
  loadDocxMock: vi.fn(),
  saveDocxMock: vi.fn(),
  paginateMock: vi.fn(),
}))

vi.mock('../../docx', () => ({
  loadDocx: loadDocxMock,
  saveDocx: saveDocxMock,
}))

vi.mock('../../docx/layout', () => ({
  paginate: paginateMock,
}))

vi.mock('../../docx/render', () => ({
  PageStack: ({ document }: { document: { sections: ReadonlyArray<{ blocks: ReadonlyArray<unknown> }> } }) => (
    <div className="docx-page-stack">
      <div className="docx-page">
        <div className="docx-page__line" data-paragraph-path="0,0">
          <span data-run-index="0">{document.sections.length}</span>
        </div>
      </div>
    </div>
  ),
  MediaContext: { Provider: ({ children }: { children: React.ReactNode }) => <>{children}</> },
}))

function createBundle() {
  return {
    document: {
      kind: 'document' as const,
      sections: [
        {
          kind: 'section' as const,
          props: {},
          blocks: [
            {
              kind: 'paragraph' as const,
              props: {},
              children: [
                {
                  kind: 'run' as const,
                  props: {},
                  children: [{ kind: 'text' as const, value: 'Hello DOCX' }],
                },
              ],
            },
          ],
        },
      ],
      styles: new Map(),
      numbering: new Map(),
      headers: new Map(),
      footers: new Map(),
      comments: new Map(),
      footnotes: new Map(),
      endnotes: new Map(),
    },
    rawArchive: new Map(),
  }
}

describe('DocxViewer editor', () => {
  beforeEach(() => {
    loadDocxMock.mockResolvedValue(createBundle())
    saveDocxMock.mockResolvedValue(new Uint8Array([1, 2, 3]))
    paginateMock.mockResolvedValue([
      {
        sectionIndex: 0,
        pageIndex: 0,
        sizePt: { width: 400, height: 300 },
        marginsPt: { top: 0, right: 0, bottom: 0, left: 0, header: 0, footer: 0, gutter: 0 },
        columns: [],
        headerLines: [],
        footerLines: [],
      },
    ])

    window.electronAPI = {
      getInitialFile: vi.fn(),
      openFileDialog: vi.fn(),
      openFileByPath: vi.fn(),
      saveFile: vi.fn(),
      onFileOpened: vi.fn(),
      setTheme: vi.fn(),
      openFileBinary: vi.fn(),
      readBinaryByPath: vi.fn(),
      onFileOpenedPath: vi.fn(),
      onSpellCheckMenu: vi.fn(() => () => {}),
      replaceMisspelling: vi.fn(),
      addWordToDictionary: vi.fn(),
      image: { pick: vi.fn() },
      // `saveBinaryFile` isn't part of the typed `ElectronAPI` (DocxViewer reads
      // it through its own local `BinaryElectronAPI` extension) — cast through
      // `unknown` rather than typing this test double as `any`.
      saveBinaryFile: vi.fn().mockResolvedValue({ saved: true, path: 'C:/docs/sample.docx' }),
    } as unknown as typeof window.electronAPI
  })

  afterEach(() => {
    vi.clearAllMocks()
  })

  it('renders editor chrome and saves through saveDocx', async () => {
    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{
            kind: 'binary',
            content: new Uint8Array([1, 2, 3]).buffer,
            path: 'C:/docs/sample.docx',
            format: 'docx',
          }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))

    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })
  })

  it('exposes a Print button that invokes window.print', async () => {
    const printSpy = vi.spyOn(window, 'print').mockImplementation(() => {})

    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{
            kind: 'binary',
            content: new Uint8Array([1, 2, 3]).buffer,
            path: 'C:/docs/sample.docx',
            format: 'docx',
          }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: /print document/i })).toBeInTheDocument()
    })

    fireEvent.click(screen.getByRole('button', { name: /print document/i }))

    expect(printSpy).toHaveBeenCalledTimes(1)
    expect(document.body.classList.contains('atlas-printing')).toBe(false)

    printSpy.mockRestore()
  })

  it('handles MS Word keyboard shortcuts (F.4)', async () => {
    const printSpy = vi.spyOn(window, 'print').mockImplementation(() => {})

    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{
            kind: 'binary',
            content: new Uint8Array([1, 2, 3]).buffer,
            path: 'C:/docs/sample.docx',
            format: 'docx',
          }}
        />
      </ViewerProvider>,
    )

    const editor = await screen.findByRole('textbox', { name: 'Document editor' })

    // Ctrl+P → print
    fireEvent.keyDown(editor, { key: 'p', ctrlKey: true })
    expect(printSpy).toHaveBeenCalledTimes(1)

    // Ctrl+F → opens find panel
    fireEvent.keyDown(editor, { key: 'f', ctrlKey: true })
    await waitFor(() => {
      expect(document.querySelector('.docx-find__input[aria-label="Find"]')).not.toBeNull()
    })

    // Ctrl+H → opens find/replace panel (idempotent: panel already open)
    fireEvent.keyDown(editor, { key: 'h', ctrlKey: true })
    expect(document.querySelector('.docx-find__input[aria-label="Replace"]')).not.toBeNull()

    // Ctrl+S — DXE-07/RUN-03: must actually save (used to swallow the key and
    // do nothing at all).
    const ctrlS = new KeyboardEvent('keydown', { key: 's', ctrlKey: true, bubbles: true, cancelable: true })
    fireEvent(editor, ctrlS)
    expect(ctrlS.defaultPrevented).toBe(true)
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    // Ctrl+L alignment shortcut — must call preventDefault
    const ctrlL = new KeyboardEvent('keydown', { key: 'l', ctrlKey: true, bubbles: true, cancelable: true })
    fireEvent(editor, ctrlL)
    expect(ctrlL.defaultPrevented).toBe(true)

    // F7 → spell check toggle — must call preventDefault
    const f7 = new KeyboardEvent('keydown', { key: 'F7', bubbles: true, cancelable: true })
    fireEvent(editor, f7)
    expect(f7.defaultPrevented).toBe(true)

    // Ctrl+K → hyperlink — must call preventDefault
    const ctrlK = new KeyboardEvent('keydown', { key: 'k', ctrlKey: true, bubbles: true, cancelable: true })
    fireEvent(editor, ctrlK)
    expect(ctrlK.defaultPrevented).toBe(true)

    // Ctrl+1 → line spacing — must call preventDefault
    const ctrl1 = new KeyboardEvent('keydown', { key: '1', ctrlKey: true, bubbles: true, cancelable: true })
    fireEvent(editor, ctrl1)
    expect(ctrl1.defaultPrevented).toBe(true)

    printSpy.mockRestore()
  })

  // ---------------------------------------------------------------------------
  // P1.1/SHELL-03/DXE-09 — document-session capability contract
  // ---------------------------------------------------------------------------

  it('P1.1: reports dirty via the shared ViewerContext after an edit, and clears it after a successful save', async () => {
    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <ViewerDirtyProbe />
        <DocxViewer
          file={{
            kind: 'binary',
            content: new Uint8Array([1, 2, 3]).buffer,
            path: 'C:/docs/sample.docx',
            format: 'docx',
          }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    expect(screen.getByTestId('viewer-dirty')).toHaveTextContent('false')

    replaceFirstMatch('Hello', 'Howdy')

    await waitFor(() => {
      expect(screen.getByTestId('viewer-dirty')).toHaveTextContent('true')
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))

    await waitFor(() => {
      expect(screen.getByTestId('viewer-dirty')).toHaveTextContent('false')
    })
  })

  it('P1.1/DXE-07/RUN-03: the shared context save() reaches this viewer\'s own save implementation', async () => {
    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <ViewerSaveProbe />
        <DocxViewer
          file={{
            kind: 'binary',
            content: new Uint8Array([1, 2, 3]).buffer,
            path: 'C:/docs/sample.docx',
            format: 'docx',
          }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    fireEvent.click(screen.getByRole('button', { name: 'Context Save' }))

    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })
  })

  // ---------------------------------------------------------------------------
  // P1.6/DXE-01 — image insert must operate on the live documentModel
  // ---------------------------------------------------------------------------

  it('P1.6/DXE-01: inserting an image preserves edits made since the file was loaded', async () => {
    window.electronAPI!.image!.pick = vi.fn().mockResolvedValue({
      cancelled: false,
      bytes: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
      mime: 'image/png',
      suggestedName: 'photo.png',
    })

    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{
            kind: 'binary',
            content: new Uint8Array([1, 2, 3]).buffer,
            path: 'C:/docs/sample.docx',
            format: 'docx',
          }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    // Edit the document first (via the real find/replace pipeline) — this
    // also leaves `range` pointing at a real, valid position, which
    // handleInsertImage needs as its insertion cursor.
    replaceFirstMatch('Hello', 'Howdy')

    // Insert an image. Before the P1.6 fix, insertImageIntoBundle ran against
    // the stale `bundle.document` (still "Hello DOCX"), silently reverting
    // the "Howdy" edit above.
    fireEvent.click(screen.getByRole('tab', { name: 'Insert' }))
    fireEvent.click(screen.getByTitle('Image'))

    await waitFor(() => {
      expect(saveDocxMock).not.toHaveBeenCalled() // sanity: haven't saved yet
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))

    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    const savedDocument = saveDocxMock.mock.calls[0][0].document as DocxDocument
    expect(collectText(savedDocument)).toContain('Howdy')
    expect(hasDrawing(savedDocument)).toBe(true)
  })
})
