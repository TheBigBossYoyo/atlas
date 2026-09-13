import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { ViewerProvider } from '../shared/ViewerContext'
import { DocxViewer } from '../DocxViewer'

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
    }
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

    // Ctrl+S → swallowed (no throw, no save dialog from browser)
    const ctrlS = new KeyboardEvent('keydown', { key: 's', ctrlKey: true, bubbles: true, cancelable: true })
    fireEvent(editor, ctrlS)
    expect(ctrlS.defaultPrevented).toBe(true)

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
})
