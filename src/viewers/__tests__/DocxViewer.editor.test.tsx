import { fireEvent, render, screen, waitFor, within } from '@testing-library/react'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import type { Document as DocxDocument, RunChild } from '../../docx/model'
import { extractCommentText } from '../../docx/editor/comments'
import { ShortcutManagerProvider } from '../../hooks/ShortcutManagerProvider'
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

function hasPageBreak(document: DocxDocument): boolean {
  return collectRunChildren(document).some(
    (child) => child.kind === 'break' && child.breakType === 'page',
  )
}

/** Every top-level `ins-revision` paragraph child's own flattened run text —
 * used to confirm a paste landed as a tracked insertion rather than plain
 * text (regression coverage for DXE-11 + DXE-19 composing correctly). */
function insRevisionTexts(document: DocxDocument): ReadonlyArray<string> {
  const texts: string[] = []
  for (const section of document.sections) {
    for (const block of section.blocks) {
      if (block.kind !== 'paragraph') continue
      for (const child of block.children as ReadonlyArray<{
        kind: string
        children?: ReadonlyArray<{ kind: string; children?: ReadonlyArray<{ kind: string; value?: string }> }>
      }>) {
        if (child.kind !== 'ins-revision') continue
        for (const run of child.children ?? []) {
          if (run.kind !== 'run') continue
          for (const grandchild of run.children ?? []) {
            if (grandchild.kind === 'text' && typeof grandchild.value === 'string') {
              texts.push(grandchild.value)
            }
          }
        }
      }
    }
  }
  return texts
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

/** DXE-14 — a document whose only top-level block is a 1x1 table, so a
 * selection at paragraph path `[0, 0, 0, 0, 0]` (section, table block, row,
 * cell, paragraph) sits inside its one cell. */
function createTableBundle() {
  return {
    document: {
      kind: 'document' as const,
      sections: [
        {
          kind: 'section' as const,
          props: {},
          blocks: [
            {
              kind: 'table' as const,
              rows: [
                {
                  kind: 'table-row' as const,
                  cells: [
                    {
                      kind: 'table-cell' as const,
                      blocks: [
                        {
                          kind: 'paragraph' as const,
                          props: {},
                          children: [
                            {
                              kind: 'run' as const,
                              props: {},
                              children: [{ kind: 'text' as const, value: 'Cell text' }],
                            },
                          ],
                        },
                      ],
                    },
                  ],
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

/**
 * DXE-14 — the mocked `PageStack` above always renders one fixed line
 * regardless of document content, so a real "cursor inside a table cell"
 * selection is built directly against the DOM instead: append a line with
 * the deep paragraph path a table cell would have, then point the browser's
 * real Selection at its run span — exactly what `getSelectionFromDom`
 * (`DocxViewer.tsx`) reads.
 */
function placeSelectionInsideSurfaceTableCell(): void {
  const surface = document.querySelector('.docx-viewer__surface') as HTMLElement
  const line = document.createElement('div')
  line.className = 'docx-page__line'
  line.setAttribute('data-paragraph-path', '0,0,0,0,0')
  const span = document.createElement('span')
  span.setAttribute('data-run-index', '0')
  span.textContent = 'Cell text'
  line.appendChild(span)
  surface.appendChild(line)

  const textNode = span.firstChild as Text
  const domRange = document.createRange()
  domRange.setStart(textNode, 0)
  domRange.setEnd(textNode, 0)
  const selection = window.getSelection()
  selection?.removeAllRanges()
  selection?.addRange(domRange)
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
      saveBinaryFile: vi.fn().mockResolvedValue({ saved: true, path: 'C:/docs/sample.docx', name: 'sample.docx' }),
      onFileOpened: vi.fn(),
      setTheme: vi.fn(),
      openFileBinary: vi.fn(),
      newDocument: vi.fn(),
      readBinaryByPath: vi.fn(),
      onFileOpenedPath: vi.fn(),
      getPathForFile: vi.fn(),
      registerDroppedPath: vi.fn().mockResolvedValue({ ok: true }),
      requestOpenRecent: vi.fn().mockResolvedValue({ ok: true }),
      revealInFolder: vi.fn().mockResolvedValue({ ok: true }),
      // Overridden per-test via `window.electronAPI!.image!.pick = ...` (see
      // the P1.6/DXE-01 image-insert test below), so it must exist here.
      image: { pick: vi.fn() },
      spellcheck: {
        onContextMenu: vi.fn(() => () => {}),
        replaceMisspelling: vi.fn(),
        addWord: vi.fn(),
        getLanguages: vi.fn().mockResolvedValue({ available: [], enabled: [] }),
        setLanguages: vi.fn().mockResolvedValue({ ok: true }),
      },
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

    // Ctrl+P → print. The shortcut is registered by an effect that runs once
    // the document has finished loading, which lags the editor element itself
    // on a slow machine (this raced on CI) — retry the key until it lands.
    await waitFor(() => {
      fireEvent.keyDown(editor, { key: 'p', ctrlKey: true })
      expect(printSpy).toHaveBeenCalled()
    })

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

  it('P1.1: an edit made while a save is still in flight is still reported dirty once that save resolves', async () => {
    // Hold saveDocx open so we can land a second edit *during* the save,
    // simulating a fast typist outrunning a slow save() round-trip. Kept on
    // an object (rather than a bare `let`) so TS doesn't narrow the
    // not-yet-assigned property to a bare `null` at the read site below.
    const pending: { resolveSave: ((bytes: Uint8Array) => void) | null } = { resolveSave: null }
    saveDocxMock.mockImplementation(() => {
      return new Promise<Uint8Array>((resolve) => {
        pending.resolveSave = resolve
      })
    })

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

    replaceFirstMatch('Hello', 'Howdy')
    await waitFor(() => {
      expect(screen.getByTestId('viewer-dirty')).toHaveTextContent('true')
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => expect(saveDocxMock).toHaveBeenCalledTimes(1))

    // A second edit lands while the first save's saveDocx()/saveBinaryFile()
    // round-trip is still pending — this content is NOT part of what's about
    // to be written to disk.
    replaceFirstMatch('Howdy', 'Howdy again')

    // Now let the in-flight save resolve with the (now-stale) bytes it
    // captured before the second edit.
    pending.resolveSave?.(new Uint8Array([1, 2, 3]))

    // Before the fix, handleSave unconditionally called setDirty(false) using
    // the documentModel it captured when the save *started*, permanently
    // clearing the dirty flag even though the second edit above was never
    // written to disk — silently telling the user there was nothing left to
    // save when there was.
    await waitFor(() => {
      expect(screen.getByTestId('viewer-dirty')).toHaveTextContent('true')
    })
  })

  // ---------------------------------------------------------------------------
  // DIRTY-1 — undo/redo back to the saved content must clear/re-set dirty.
  // Before the fix, dirty was `documentModel !== lastSavedDocument` by
  // reference; `History.undo`/`redo` always rebuild the document via
  // `applyCommand`, so even undoing back to byte-for-byte what was on disk
  // produced a document that was never `===` to the saved snapshot, and the
  // unsaved-changes dot (and close-prompt, driven by this same
  // `useViewerIsDirty` flag) never cleared.
  // ---------------------------------------------------------------------------

  it('DIRTY-1: save, edit, then undo back to the saved content clears dirty', async () => {
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

    const editor = await screen.findByRole('textbox', { name: 'Document editor' })
    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })
    expect(screen.getByTestId('viewer-dirty')).toHaveTextContent('false')

    // Save the document as-is, establishing "Hello DOCX" as the on-disk
    // baseline.
    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })
    expect(screen.getByTestId('viewer-dirty')).toHaveTextContent('false')

    // One more edit past the save point.
    replaceFirstMatch('Hello', 'Howdy')
    await waitFor(() => {
      expect(screen.getByTestId('viewer-dirty')).toHaveTextContent('true')
    })

    // Undo it — back to exactly "Hello DOCX", what's already on disk.
    const ctrlZ = new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true, cancelable: true })
    fireEvent(editor, ctrlZ)

    await waitFor(() => {
      expect(screen.getByTestId('viewer-dirty')).toHaveTextContent('false')
    })
  })

  it('DIRTY-1: save, edit, undo, then redo is dirty again', async () => {
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

    const editor = await screen.findByRole('textbox', { name: 'Document editor' })
    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    replaceFirstMatch('Hello', 'Howdy')
    await waitFor(() => {
      expect(screen.getByTestId('viewer-dirty')).toHaveTextContent('true')
    })

    const ctrlZ = new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true, cancelable: true })
    fireEvent(editor, ctrlZ)
    await waitFor(() => {
      expect(screen.getByTestId('viewer-dirty')).toHaveTextContent('false')
    })

    // Redo the edit back in — content diverges from disk again.
    const ctrlShiftZ = new KeyboardEvent('keydown', {
      key: 'Z',
      ctrlKey: true,
      shiftKey: true,
      bubbles: true,
      cancelable: true,
    })
    fireEvent(editor, ctrlShiftZ)

    await waitFor(() => {
      expect(screen.getByTestId('viewer-dirty')).toHaveTextContent('true')
    })
  })

  it('DIRTY-1: edit, save, then undo past the save point is dirty (content no longer matches disk)', async () => {
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

    const editor = await screen.findByRole('textbox', { name: 'Document editor' })
    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    // Edit, then save — "Howdy DOCX" is now the on-disk baseline.
    replaceFirstMatch('Hello', 'Howdy')
    await waitFor(() => {
      expect(screen.getByTestId('viewer-dirty')).toHaveTextContent('true')
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })
    await waitFor(() => {
      expect(screen.getByTestId('viewer-dirty')).toHaveTextContent('false')
    })

    // Undo the edit that was just saved — back to "Hello DOCX", which no
    // longer matches what's on disk ("Howdy DOCX").
    const ctrlZ = new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true, cancelable: true })
    fireEvent(editor, ctrlZ)

    await waitFor(() => {
      expect(screen.getByTestId('viewer-dirty')).toHaveTextContent('true')
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

  it('DXE-18: Ctrl+Z after inserting an image removes it again', async () => {
    window.electronAPI!.image!.pick = vi.fn().mockResolvedValue({
      cancelled: false,
      bytes: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
      mime: 'image/png',
      suggestedName: 'photo.png',
    })

    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    replaceFirstMatch('Hello', 'Howdy')

    fireEvent.click(screen.getByRole('tab', { name: 'Insert' }))
    fireEvent.click(screen.getByLabelText('Image'))

    // Image insertion is async (awaits the picker, then
    // decodeImageNaturalSizePt); give it a tick to land — the same pattern
    // the existing P1.6/DXE-01 test above already relies on — before undoing
    // it, so Ctrl+Z targets the image-insert's inverse rather than firing
    // before it was ever pushed to History.
    await waitFor(() => {
      expect(saveDocxMock).not.toHaveBeenCalled()
    })

    const editor = await screen.findByRole('textbox', { name: 'Document editor' })
    const ctrlZ = new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true, cancelable: true })
    fireEvent(editor, ctrlZ)

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    const savedDocument = saveDocxMock.mock.calls[0][0].document as DocxDocument
    expect(hasDrawing(savedDocument)).toBe(false)
    expect(collectText(savedDocument)).toContain('Howdy')
  })

  it('the paragraph stays editable after inserting an image into it (no getEditableRuns poisoning)', async () => {
    window.electronAPI!.image!.pick = vi.fn().mockResolvedValue({
      cancelled: false,
      bytes: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
      mime: 'image/png',
      suggestedName: 'photo.png',
    })

    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    replaceFirstMatch('Hello', 'Howdy')

    fireEvent.click(screen.getByRole('tab', { name: 'Insert' }))
    fireEvent.click(screen.getByLabelText('Image'))

    await waitFor(() => {
      expect(saveDocxMock).not.toHaveBeenCalled()
    })

    // The image landed inline in the same (only) paragraph. A subsequent,
    // unrelated edit to that same paragraph must still work — before the
    // getEditableRuns fix, any run with a non-text child (the inserted
    // drawing) made the *whole paragraph* reject every later command.
    replaceFirstMatch('DOCX', 'Editor')

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    const savedDocument = saveDocxMock.mock.calls[0][0].document as DocxDocument
    expect(hasDrawing(savedDocument)).toBe(true)
    expect(collectText(savedDocument)).toContain('Howdy')
    expect(collectText(savedDocument)).toContain('Editor')
  })

  it('DXE-24: Save As omits existingPath so a save dialog is shown', async () => {
    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save As' })).toBeInTheDocument()
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save As' }))

    await waitFor(() => {
      expect(window.electronAPI!.saveBinaryFile).toHaveBeenCalledTimes(1)
    })

    const request = (window.electronAPI!.saveBinaryFile as ReturnType<typeof vi.fn>).mock.calls[0][0]
    expect(request.existingPath).toBeUndefined()
  })

  it('DXE-25: a save failure stays visible until dismissed or a save succeeds, not cleared by the next edit', async () => {
    window.electronAPI!.saveBinaryFile = vi.fn().mockResolvedValue({ saved: false, error: 'Disk is full' })

    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(screen.getByRole('alert')).toHaveTextContent('Disk is full')
    })

    // A further edit must NOT clear the banner.
    replaceFirstMatch('Hello', 'Howdy')
    expect(screen.getByRole('alert')).toHaveTextContent('Disk is full')

    fireEvent.click(screen.getByRole('button', { name: 'Dismiss error' }))
    expect(screen.queryByRole('alert')).not.toBeInTheDocument()
  })

  // ---------------------------------------------------------------------------
  // D15/DXE-08 — spellcheck accept fixes the model, not just the DOM
  // ---------------------------------------------------------------------------

  it('accepting a spellcheck suggestion fixes the word in the saved document', async () => {
    const listeners = new Set<(payload: { word: string; suggestions: string[]; x: number; y: number }) => void>()
    window.electronAPI!.spellcheck = {
      onContextMenu: vi.fn((cb) => {
        listeners.add(cb)
        return () => listeners.delete(cb)
      }),
      replaceMisspelling: vi.fn(),
      addWord: vi.fn(),
      getLanguages: vi.fn().mockResolvedValue({ available: [], enabled: [] }),
      setLanguages: vi.fn().mockResolvedValue({ ok: true }),
    }

    loadDocxMock.mockResolvedValue({
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
                children: [{ kind: 'run' as const, props: {}, children: [{ kind: 'text' as const, value: 'I went ot the store' }] }],
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
    })

    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    fireEvent.mouseUp(await screen.findByRole('textbox', { name: 'Document editor' }))

    listeners.forEach((cb) => cb({ word: 'ot', suggestions: ['to'], x: 5, y: 5 }))
    await waitFor(() => {
      expect(screen.getByText('to')).toBeInTheDocument()
    })
    fireEvent.click(screen.getByText('to'))

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    const savedDocument = saveDocxMock.mock.calls[0][0].document as DocxDocument
    expect(collectText(savedDocument)).toBe('I went to the store')
    // The old behavior only called the native bridge — the fix must now come
    // from the model-level command pipeline instead.
    expect(window.electronAPI!.spellcheck.replaceMisspelling).not.toHaveBeenCalled()
  })

  // ---------------------------------------------------------------------------
  // D18 — hyperlink insertion and list toggling
  // ---------------------------------------------------------------------------

  // F1 — Electron doesn't implement `window.prompt` (it throws "prompt() is
  // not supported"), so Insert Hyperlink used to do nothing at all in the
  // packaged app even though these tests previously passed: they mocked
  // `window.prompt` directly, which jsdom is happy to let a test stub out,
  // masking the real-Electron failure entirely. These now drive the actual
  // DocxPromptDialog (see DocxPromptDialog.tsx) instead, and assert
  // `window.prompt` is never called — the regression this batch fixes.
  it('inserts a hyperlink via the dialog, prefilled with https://, and persists it on save', async () => {
    const promptSpy = vi.spyOn(window, 'prompt')

    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    replaceFirstMatch('Hello', 'Howdy')

    fireEvent.click(screen.getByRole('tab', { name: 'Insert' }))
    fireEvent.click(screen.getByLabelText('Hyperlink'))

    const dialog = screen.getByRole('dialog', { name: 'Insert hyperlink' })
    const urlInput = within(dialog).getByLabelText('URL')
    expect(urlInput).toHaveValue('https://')
    expect(urlInput).toHaveFocus()

    fireEvent.change(urlInput, { target: { value: 'https://example.com' } })
    fireEvent.click(within(dialog).getByRole('button', { name: 'Insert' }))

    expect(dialog).not.toBeInTheDocument()
    expect(promptSpy).not.toHaveBeenCalled()

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    const savedBundle = saveDocxMock.mock.calls[0][0] as { document: DocxDocument; relationships?: ReadonlyArray<{ target: string; targetMode?: string }> }
    expect(savedBundle.relationships?.some((rel) => rel.target === 'https://example.com' && rel.targetMode === 'External')).toBe(true)

    promptSpy.mockRestore()
  })

  it('does not insert a hyperlink when the dialog is cancelled', async () => {
    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    replaceFirstMatch('Hello', 'Howdy')
    fireEvent.click(screen.getByRole('tab', { name: 'Insert' }))
    fireEvent.click(screen.getByLabelText('Hyperlink'))

    const dialog = screen.getByRole('dialog', { name: 'Insert hyperlink' })
    fireEvent.change(within(dialog).getByLabelText('URL'), { target: { value: 'https://example.com' } })
    fireEvent.click(within(dialog).getByRole('button', { name: 'Cancel' }))

    expect(screen.queryByRole('dialog', { name: 'Insert hyperlink' })).not.toBeInTheDocument()

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    const savedBundle = saveDocxMock.mock.calls[0][0] as { relationships?: ReadonlyArray<unknown> }
    expect(savedBundle.relationships ?? []).toHaveLength(0)
  })

  it('does not insert a hyperlink when the dialog is closed with Escape', async () => {
    // DocxPromptDialog's Escape-close registers through the shared shortcut
    // dispatcher (matching UnsavedChangesDialog), which needs a real
    // <ShortcutManagerProvider> ancestor to exercise for real — App.tsx
    // always provides one in production (every other test above renders
    // without it and only exercises the direct onKeyDown-driven document
    // commands, which don't go through the dispatcher).
    render(
      <ShortcutManagerProvider>
        <ViewerProvider filePath="C:/docs/sample.docx">
          <DocxViewer
            file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
          />
        </ViewerProvider>
      </ShortcutManagerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    replaceFirstMatch('Hello', 'Howdy')
    fireEvent.click(screen.getByRole('tab', { name: 'Insert' }))
    fireEvent.click(screen.getByLabelText('Hyperlink'))

    const dialog = screen.getByRole('dialog', { name: 'Insert hyperlink' })
    const urlInput = within(dialog).getByLabelText('URL')
    fireEvent.change(urlInput, { target: { value: 'https://example.com' } })
    fireEvent.keyDown(window, { key: 'Escape' })

    expect(screen.queryByRole('dialog', { name: 'Insert hyperlink' })).not.toBeInTheDocument()

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    const savedBundle = saveDocxMock.mock.calls[0][0] as { relationships?: ReadonlyArray<unknown> }
    expect(savedBundle.relationships ?? []).toHaveLength(0)
  })

  // ---------------------------------------------------------------------------
  // F1 — Add Comment / Reply, also previously silent no-ops via
  // `window.prompt` (see the hyperlink block above for the full context).
  // ---------------------------------------------------------------------------

  it('adds a comment via the dialog and replies to it via the dialog; both persist on save', async () => {
    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    replaceFirstMatch('Hello', 'Howdy')

    fireEvent.click(screen.getByRole('tab', { name: 'Insert' }))
    fireEvent.click(screen.getByLabelText('Comment'))

    const commentDialog = screen.getByRole('dialog', { name: 'Add comment' })
    const commentInput = within(commentDialog).getByLabelText('Comment')
    expect(commentInput).toHaveFocus()
    fireEvent.change(commentInput, { target: { value: 'First comment' } })
    fireEvent.click(within(commentDialog).getByRole('button', { name: 'Add comment' }))
    expect(commentDialog).not.toBeInTheDocument()

    await waitFor(() => {
      expect(screen.getByText('First comment')).toBeInTheDocument()
    })

    // The comments pane's own "Reply" trigger — unambiguous here since the
    // reply dialog isn't open yet.
    fireEvent.click(screen.getByRole('button', { name: 'Reply' }))

    const replyDialog = screen.getByRole('dialog', { name: 'Reply to comment' })
    const replyInput = within(replyDialog).getByLabelText('Reply')
    expect(replyInput).toHaveFocus()
    fireEvent.change(replyInput, { target: { value: 'A reply' } })
    // Scoped to the dialog: the pane's own "Reply" trigger button is still
    // in the DOM behind it with the same accessible name.
    fireEvent.click(within(replyDialog).getByRole('button', { name: 'Reply' }))
    expect(replyDialog).not.toBeInTheDocument()

    await waitFor(() => {
      expect(screen.getByText('A reply')).toBeInTheDocument()
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    const savedBundle = saveDocxMock.mock.calls[0][0] as { document: DocxDocument }
    expect(savedBundle.document.comments.size).toBe(2)
    const texts = Array.from(savedBundle.document.comments.values()).map(extractCommentText)
    expect(texts).toContain('First comment')
    expect(texts).toContain('A reply')
  })

  it('does not add a comment when the dialog is cancelled', async () => {
    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    replaceFirstMatch('Hello', 'Howdy')
    fireEvent.click(screen.getByRole('tab', { name: 'Insert' }))
    fireEvent.click(screen.getByLabelText('Comment'))

    const commentDialog = screen.getByRole('dialog', { name: 'Add comment' })
    fireEvent.change(within(commentDialog).getByLabelText('Comment'), { target: { value: 'Should not be saved' } })
    fireEvent.click(within(commentDialog).getByRole('button', { name: 'Cancel' }))

    expect(screen.queryByRole('dialog', { name: 'Add comment' })).not.toBeInTheDocument()
    expect(screen.queryByText('Should not be saved')).not.toBeInTheDocument()

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    const savedBundle = saveDocxMock.mock.calls[0][0] as { document: DocxDocument }
    expect(savedBundle.document.comments.size).toBe(0)
  })

  it('does not add a reply when the reply dialog is cancelled', async () => {
    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    replaceFirstMatch('Hello', 'Howdy')
    fireEvent.click(screen.getByRole('tab', { name: 'Insert' }))
    fireEvent.click(screen.getByLabelText('Comment'))

    const commentDialog = screen.getByRole('dialog', { name: 'Add comment' })
    fireEvent.change(within(commentDialog).getByLabelText('Comment'), { target: { value: 'Root comment' } })
    fireEvent.click(within(commentDialog).getByRole('button', { name: 'Add comment' }))

    await waitFor(() => {
      expect(screen.getByText('Root comment')).toBeInTheDocument()
    })

    fireEvent.click(screen.getByRole('button', { name: 'Reply' }))
    const replyDialog = screen.getByRole('dialog', { name: 'Reply to comment' })
    fireEvent.change(within(replyDialog).getByLabelText('Reply'), { target: { value: 'Should not be saved' } })
    fireEvent.click(within(replyDialog).getByRole('button', { name: 'Cancel' }))

    expect(screen.queryByRole('dialog', { name: 'Reply to comment' })).not.toBeInTheDocument()
    expect(screen.queryByText('Should not be saved')).not.toBeInTheDocument()

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    // Only the root comment — no reply was added.
    const savedBundle = saveDocxMock.mock.calls[0][0] as { document: DocxDocument }
    expect(savedBundle.document.comments.size).toBe(1)
  })

  it('toggling a bulleted list sets numPr and creates a numbering definition that survives save', async () => {
    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    replaceFirstMatch('Hello', 'Howdy')

    fireEvent.click(screen.getByLabelText('Bullet List'))

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    const savedBundle = saveDocxMock.mock.calls[0][0] as {
      document: DocxDocument
      numberingPart?: { nums: ReadonlyMap<string, unknown> }
    }
    const paragraph = savedBundle.document.sections[0].blocks[0] as { props?: { numPr?: { numId?: string } } }
    expect(paragraph.props?.numPr?.numId).toBe('1')
    expect(savedBundle.numberingPart?.nums.has('1')).toBe(true)
  })

  it('inserts a page break at the caret via the Insert tab and it survives save/undo', async () => {
    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    // Establishes a real, valid caret position (see the image-insert tests
    // above for why this matters) before inserting the break there.
    replaceFirstMatch('Hello', 'Howdy')

    fireEvent.click(screen.getByRole('tab', { name: 'Insert' }))
    fireEvent.click(screen.getByLabelText('Page Break'))

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    const savedDocument = saveDocxMock.mock.calls[0][0].document as DocxDocument
    expect(hasPageBreak(savedDocument)).toBe(true)
    expect(collectText(savedDocument)).toContain('Howdy')
  })

  it('Ctrl+Z after inserting a page break removes it again', async () => {
    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    replaceFirstMatch('Hello', 'Howdy')

    fireEvent.click(screen.getByRole('tab', { name: 'Insert' }))
    fireEvent.click(screen.getByLabelText('Page Break'))

    const editor = await screen.findByRole('textbox', { name: 'Document editor' })
    const ctrlZ = new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true, cancelable: true })
    fireEvent(editor, ctrlZ)

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    const savedDocument = saveDocxMock.mock.calls[0][0].document as DocxDocument
    expect(hasPageBreak(savedDocument)).toBe(false)
    expect(collectText(savedDocument)).toContain('Howdy')
  })

  // ---------------------------------------------------------------------------
  // D30/DXE-25 review-tab toggles
  // ---------------------------------------------------------------------------

  it('toggle-spell-check flips the editor surface spellCheck attribute', async () => {
    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
        />
      </ViewerProvider>,
    )

    const editor = await screen.findByRole('textbox', { name: 'Document editor' })
    expect(editor).toHaveAttribute('spellcheck', 'true')

    fireEvent.click(screen.getByRole('tab', { name: 'Review' }))
    fireEvent.click(screen.getByLabelText('Toggle spell check'))

    expect(editor).toHaveAttribute('spellcheck', 'false')
  })

  // ---------------------------------------------------------------------------
  // D17/DXE-11 track-changes setting persistence
  // ---------------------------------------------------------------------------

  it('toggling track changes persists bundle.settings.trackChanges into the saved document', async () => {
    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    fireEvent.click(screen.getByRole('tab', { name: 'Review' }))
    fireEvent.click(screen.getByLabelText('Toggle track changes'))

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    const savedBundle = saveDocxMock.mock.calls[0][0] as { settings?: { trackChanges: boolean } }
    expect(savedBundle.settings?.trackChanges).toBe(true)
  })

  // ---------------------------------------------------------------------------
  // P2.1 — active-viewer combo registration must not swallow native
  // text-editing shortcuts in the viewer's own plain <input> fields
  // ---------------------------------------------------------------------------

  it(
    'P2.1: reserves Ctrl+B for the document while focus is on a non-field element (e.g. after clicking a ' +
      'Toolbar button), but leaves native shortcuts alone inside the Find/Replace panel\'s own <input>s',
    async () => {
      render(
        <ShortcutManagerProvider>
          <ViewerProvider filePath="C:/docs/sample.docx">
            <DocxViewer
              file={{
                kind: 'binary',
                content: new Uint8Array([1, 2, 3]).buffer,
                path: 'C:/docs/sample.docx',
                format: 'docx',
              }}
            />
          </ViewerProvider>
        </ShortcutManagerProvider>,
      )

      const editor = await screen.findByRole('textbox', { name: 'Document editor' })

      // Open Find (also proves the viewer-tier registration is reachable
      // through a real window-level dispatch when wrapped in the real
      // ShortcutManagerProvider, not just via the direct onKeyDown prop).
      const ctrlF = new KeyboardEvent('keydown', { key: 'f', ctrlKey: true, bubbles: true, cancelable: true })
      fireEvent(editor, ctrlF)
      const findInput = await screen.findByLabelText('Find')

      // Focus is now on a genuine <input>, not the contentEditable — Ctrl+A
      // (select all text typed into the search box) must NOT be swallowed by
      // the viewer's own reserved-combo detector.
      findInput.focus()
      const ctrlA = new KeyboardEvent('keydown', { key: 'a', ctrlKey: true, bubbles: true, cancelable: true })
      fireEvent(findInput, ctrlA)
      expect(ctrlA.defaultPrevented).toBe(false)

      // Same for Ctrl+Z (undo typed text) and Ctrl+B (would otherwise bold
      // the document) while focus is in the search box.
      const ctrlZ = new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true, cancelable: true })
      fireEvent(findInput, ctrlZ)
      expect(ctrlZ.defaultPrevented).toBe(false)

      const ctrlBInField = new KeyboardEvent('keydown', { key: 'b', ctrlKey: true, bubbles: true, cancelable: true })
      fireEvent(findInput, ctrlBInField)
      expect(ctrlBInField.defaultPrevented).toBe(false)

      // But away from any plain field — e.g. focus on a toolbar button,
      // mirroring "clicked Bold, then pressed Ctrl+B again" — Ctrl+B is
      // still reserved for the document (SHELL-08's own scenario), not left
      // to fall through to a shell-global handler.
      const boldButton = screen.getByRole('button', { name: /bold/i })
      boldButton.focus()
      const ctrlBOnButton = new KeyboardEvent('keydown', { key: 'b', ctrlKey: true, bubbles: true, cancelable: true })
      fireEvent(boldButton, ctrlBOnButton)
      expect(ctrlBOnButton.defaultPrevented).toBe(true)
    },
  )

  // ---------------------------------------------------------------------------
  // DXE-14 — right-click table-editing context menu
  // ---------------------------------------------------------------------------

  it('right-clicking inside a table cell opens the table menu, and picking an entry applies its command', async () => {
    loadDocxMock.mockResolvedValue(createTableBundle())
    window.getSelection()?.removeAllRanges()

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

    placeSelectionInsideSurfaceTableCell()
    const surface = document.querySelector('.docx-viewer__surface') as HTMLElement
    const contextMenuEvent = new MouseEvent('contextmenu', {
      bubbles: true,
      cancelable: true,
      clientX: 40,
      clientY: 60,
    })
    fireEvent(surface, contextMenuEvent)
    expect(contextMenuEvent.defaultPrevented).toBe(true)

    expect(await screen.findByRole('menu', { name: 'Table editing' })).toBeInTheDocument()

    fireEvent.click(screen.getByText('Delete Table'))

    await waitFor(() => {
      expect(screen.getByTestId('viewer-dirty')).toHaveTextContent('true')
    })
    expect(screen.queryByRole('menu', { name: 'Table editing' })).not.toBeInTheDocument()
  })

  it('right-clicking outside a table leaves the native context menu alone', async () => {
    window.getSelection()?.removeAllRanges()

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

    const surface = document.querySelector('.docx-viewer__surface') as HTMLElement
    const contextMenuEvent = new MouseEvent('contextmenu', {
      bubbles: true,
      cancelable: true,
      clientX: 10,
      clientY: 10,
    })
    fireEvent(surface, contextMenuEvent)

    expect(contextMenuEvent.defaultPrevented).toBe(false)
    expect(screen.queryByRole('menu', { name: 'Table editing' })).not.toBeInTheDocument()
  })

  // ---------------------------------------------------------------------------
  // DXE-19 — rich HTML paste (tables/hyperlinks/formatting) wired into the
  // real `onPaste` handler, going through the async bundle-patch path.
  // ---------------------------------------------------------------------------

  it('pastes rich HTML as formatted text and records an external hyperlink relationship', async () => {
    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
        />
        <ViewerDirtyProbe />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    // Establishes a real, non-null cursor via a genuine find/replace command
    // (see `replaceFirstMatch`'s own doc comment above) rather than
    // fabricating a DOM selection the mocked PageStack can't render — paste
    // needs `range?.focus` populated just like hyperlink insertion does.
    replaceFirstMatch('Hello', 'Howdy')

    const editor = await screen.findByRole('textbox', { name: 'Document editor' })
    fireEvent.paste(editor, {
      clipboardData: {
        getData: (format: string) =>
          format === 'text/html'
            ? '<p>Pasted <b>bold</b> and <a href="https://example.com">a link</a></p>'
            : '',
      },
    })

    await waitFor(() => {
      expect(screen.getByTestId('viewer-dirty')).toHaveTextContent('true')
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    const savedBundle = saveDocxMock.mock.calls[0][0] as {
      document: DocxDocument
      relationships?: ReadonlyArray<{ target: string; targetMode?: string }>
    }
    expect(collectText(savedBundle.document)).toBe('HowdyPasted bold and a link DOCX')
    expect(
      savedBundle.relationships?.some(
        (rel) => rel.target === 'https://example.com' && rel.targetMode === 'External',
      ),
    ).toBe(true)
  })

  it('falls back to plain-text paste when the clipboard carries no parseable HTML blocks', async () => {
    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
        />
        <ViewerDirtyProbe />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    replaceFirstMatch('Hello DOCX', 'Howdy ')

    const editor = await screen.findByRole('textbox', { name: 'Document editor' })
    fireEvent.paste(editor, {
      clipboardData: {
        getData: (format: string) => (format === 'text/plain' ? 'plain paste' : ''),
      },
    })

    await waitFor(() => {
      expect(screen.getByTestId('viewer-dirty')).toHaveTextContent('true')
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    const savedDocument = saveDocxMock.mock.calls[0][0].document as DocxDocument
    expect(collectText(savedDocument)).toBe('Howdy plain paste')
  })

  // ---------------------------------------------------------------------------
  // Regression — DXE-11 (track changes) + DXE-19 (paste) composing correctly.
  // `applyEditorCommands`/`handleRichPaste` used to apply their command batch
  // with no `trackChanges` argument at all, so pasting while Track Changes
  // was on silently produced a plain, untracked edit instead of a `w:ins` —
  // exactly the edit a reviewer relying on Track Changes most needs to see.
  // ---------------------------------------------------------------------------

  it('records a plain-text paste as a tracked insertion when Track Changes is on', async () => {
    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
        />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    fireEvent.click(screen.getByRole('tab', { name: 'Review' }))
    fireEvent.click(screen.getByLabelText('Toggle track changes'))

    replaceFirstMatch('Hello DOCX', 'Howdy ')

    const editor = await screen.findByRole('textbox', { name: 'Document editor' })
    fireEvent.paste(editor, {
      clipboardData: {
        getData: (format: string) => (format === 'text/plain' ? 'plain paste' : ''),
      },
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    const savedDocument = saveDocxMock.mock.calls[0][0].document as DocxDocument
    expect(insRevisionTexts(savedDocument)).toContain('plain paste')
  })

  it('records rich HTML paste as a tracked insertion when Track Changes is on', async () => {
    render(
      <ViewerProvider filePath="C:/docs/sample.docx">
        <DocxViewer
          file={{ kind: 'binary', content: new Uint8Array([1, 2, 3]).buffer, path: 'C:/docs/sample.docx', format: 'docx' }}
        />
        <ViewerDirtyProbe />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument()
    })

    fireEvent.click(screen.getByRole('tab', { name: 'Review' }))
    fireEvent.click(screen.getByLabelText('Toggle track changes'))

    replaceFirstMatch('Hello', 'Howdy')

    const editor = await screen.findByRole('textbox', { name: 'Document editor' })
    fireEvent.paste(editor, {
      clipboardData: {
        getData: (format: string) => (format === 'text/html' ? '<p>Pasted text</p>' : ''),
      },
    })

    await waitFor(() => {
      expect(screen.getByTestId('viewer-dirty')).toHaveTextContent('true')
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })

    const savedDocument = saveDocxMock.mock.calls[0][0].document as DocxDocument
    expect(insRevisionTexts(savedDocument)).toContain('Pasted text')
  })
})
