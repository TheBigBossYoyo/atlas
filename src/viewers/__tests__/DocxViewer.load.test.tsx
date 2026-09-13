import { act, render, screen, waitFor } from '@testing-library/react'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { ViewerProvider } from '../shared/ViewerContext'
import { DocxViewer } from '../DocxViewer'
import type { LoadedFile } from '../../formats/types'

// Regression coverage for the file-change effect in `DocxViewerBase` (see
// DocxViewer.tsx) — reworked to fix `react-hooks/set-state-in-effect` while
// preserving its reset/load/cancel behavior exactly. These tests exercise the
// behaviors that refactor had to preserve: showing a load error, clearing
// state on file switch, and ignoring a stale in-flight load's outcome after
// the user has already navigated away from the file that started it.

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
  PageStack: ({ document }: { document: { sections: ReadonlyArray<unknown> } }) => (
    <div className="docx-page-stack">
      <div className="docx-page">
        <div className="docx-page__line" data-paragraph-path="0,0">
          <span data-testid="section-count">{document.sections.length}</span>
        </div>
      </div>
    </div>
  ),
  MediaContext: { Provider: ({ children }: { children: React.ReactNode }) => <>{children}</> },
}))

function createBundle(sectionCount: number) {
  return {
    document: {
      kind: 'document' as const,
      sections: Array.from({ length: sectionCount }, () => ({
        kind: 'section' as const,
        props: {},
        blocks: [],
      })),
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

function binaryFile(path: string): LoadedFile {
  return {
    kind: 'binary',
    content: new Uint8Array([1, 2, 3]).buffer,
    path,
    format: 'docx',
  }
}

describe('DocxViewer file-change effect', () => {
  beforeEach(() => {
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
      saveBinaryFile: vi.fn(),
      onFileOpened: vi.fn(),
      setTheme: vi.fn(),
      openFileBinary: vi.fn(),
      readBinaryByPath: vi.fn(),
      onFileOpenedPath: vi.fn(),
      getPathForFile: vi.fn(),
      registerDroppedPath: vi.fn(),
      requestOpenRecent: vi.fn(),
      revealInFolder: vi.fn(),
      spellcheck: {
        onContextMenu: vi.fn(() => () => {}),
        replaceMisspelling: vi.fn(),
        addWord: vi.fn(),
        getLanguages: vi.fn(),
        setLanguages: vi.fn(),
      },
    }
  })

  afterEach(() => {
    vi.clearAllMocks()
  })

  it('shows a load error when loadDocx rejects', async () => {
    loadDocxMock.mockRejectedValueOnce(new Error('corrupt archive'))

    render(
      <ViewerProvider filePath="C:/docs/bad.docx">
        <DocxViewer file={binaryFile('C:/docs/bad.docx')} />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByText('Failed to render DOCX: corrupt archive')).toBeInTheDocument()
    })
  })

  it('clears a previous load error and renders the new file after switching files', async () => {
    loadDocxMock.mockRejectedValueOnce(new Error('corrupt archive'))

    const { rerender } = render(
      <ViewerProvider filePath="C:/docs/bad.docx">
        <DocxViewer file={binaryFile('C:/docs/bad.docx')} />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByText('Failed to render DOCX: corrupt archive')).toBeInTheDocument()
    })

    loadDocxMock.mockResolvedValueOnce(createBundle(1))

    rerender(
      <ViewerProvider filePath="C:/docs/good.docx">
        <DocxViewer file={binaryFile('C:/docs/good.docx')} />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.queryByText(/Failed to render DOCX/)).not.toBeInTheDocument()
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })
  })

  it('ignores a stale load failure after switching to a file that loads successfully', async () => {
    // Regression test for the `cancelled` guard on the effect's catch branch:
    // without it, a slow load for a file the user has already navigated away
    // from can still call `setErrorMessage` after the fact and clobber the
    // correctly-loaded replacement file with a spurious error screen.
    let rejectFirstLoad: (error: Error) => void = () => {}
    const firstLoad = new Promise<ReturnType<typeof createBundle>>((_resolve, reject) => {
      rejectFirstLoad = reject
    })
    loadDocxMock.mockImplementationOnce(() => firstLoad)
    loadDocxMock.mockImplementationOnce(() => Promise.resolve(createBundle(2)))

    const { rerender } = render(
      <ViewerProvider filePath="C:/docs/first.docx">
        <DocxViewer file={binaryFile('C:/docs/first.docx')} />
      </ViewerProvider>,
    )

    // Switch to a second file before the first file's load settles.
    rerender(
      <ViewerProvider filePath="C:/docs/second.docx">
        <DocxViewer file={binaryFile('C:/docs/second.docx')} />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    // Reject the stale first load. It must not overwrite the second file's
    // already-rendered editor with an error screen.
    await act(async () => {
      rejectFirstLoad(new Error('first file corrupt'))
      await Promise.resolve()
      await Promise.resolve()
    })

    expect(screen.queryByText(/Failed to render DOCX/)).not.toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
  })

  it('does not throw when a stale successful load resolves after switching files', async () => {
    let resolveFirstLoad: (bundle: ReturnType<typeof createBundle>) => void = () => {}
    const firstLoad = new Promise<ReturnType<typeof createBundle>>(resolve => {
      resolveFirstLoad = resolve
    })

    loadDocxMock.mockImplementationOnce(() => firstLoad)
    loadDocxMock.mockImplementationOnce(() => Promise.resolve(createBundle(2)))

    const { rerender } = render(
      <ViewerProvider filePath="C:/docs/first.docx">
        <DocxViewer file={binaryFile('C:/docs/first.docx')} />
      </ViewerProvider>,
    )

    rerender(
      <ViewerProvider filePath="C:/docs/second.docx">
        <DocxViewer file={binaryFile('C:/docs/second.docx')} />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByTestId('section-count')).toHaveTextContent('2')
    })

    await act(async () => {
      resolveFirstLoad(createBundle(1))
      await Promise.resolve()
      await Promise.resolve()
    })

    // The second file's content must still be what's displayed.
    expect(screen.getByTestId('section-count')).toHaveTextContent('2')
  })
})
