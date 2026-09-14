/**
 * DEFER-5 / DXS-20 — "Update Fields" / "Update TOC" buttons wired into
 * DocxViewer's topbar. Mirrors DocxViewer.editor.test.tsx's mock setup
 * (mocking only `docx`'s load/save and `docx/layout`'s `paginate` + the
 * `PageStack` renderer) so the real `docx/fields` module and `Document`
 * model are exercised end to end.
 */
import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import type { Document as DocxDocument } from '../../docx/model'
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
  PageStack: () => <div className="docx-page-stack" />,
  MediaContext: { Provider: ({ children }: { children: React.ReactNode }) => <>{children}</> },
}))

function authorFieldDocument(): DocxDocument {
  return {
    kind: 'document',
    sections: [
      {
        kind: 'section',
        props: {},
        blocks: [
          {
            kind: 'paragraph',
            children: [
              {
                kind: 'field',
                fieldType: 'AUTHOR',
                instruction: 'AUTHOR',
                result: [{ kind: 'run', children: [{ kind: 'text', value: 'Stale Author' }] }],
                raw: '<w:fldSimple w:instr="AUTHOR"><w:r><w:t>Stale Author</w:t></w:r></w:fldSimple>',
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
  }
}

function documentWithNoFields(): DocxDocument {
  return {
    kind: 'document',
    sections: [
      {
        kind: 'section',
        props: {},
        blocks: [{ kind: 'paragraph', children: [{ kind: 'run', children: [{ kind: 'text', value: 'Plain text' }] }] }],
      },
    ],
    styles: new Map(),
    numbering: new Map(),
    headers: new Map(),
    footers: new Map(),
    comments: new Map(),
    footnotes: new Map(),
    endnotes: new Map(),
  }
}

const CORE_XML =
  '<?xml version="1.0"?><cp:coreProperties xmlns:dc="a" xmlns:cp="b"><dc:creator>Real Author</dc:creator></cp:coreProperties>'

function coreXmlBytes(): Uint8Array {
  return new TextEncoder().encode(CORE_XML)
}

function renderViewer(): void {
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
}

describe('DocxViewer field commands (DEFER-5 / DXS-20)', () => {
  beforeEach(() => {
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
  })

  afterEach(() => {
    vi.clearAllMocks()
  })

  it('renders "Update Fields" and "Update TOC" buttons', async () => {
    loadDocxMock.mockResolvedValue({ document: documentWithNoFields(), rawArchive: new Map() })
    renderViewer()

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Update fields' })).toBeInTheDocument()
    })
    expect(screen.getByRole('button', { name: 'Update table of contents' })).toBeInTheDocument()
  })

  it('recalculates an AUTHOR field from docProps/core.xml\'s dc:creator and shows a status message', async () => {
    loadDocxMock.mockResolvedValue({
      document: authorFieldDocument(),
      rawArchive: new Map([['docProps/core.xml', coreXmlBytes()]]),
    })
    renderViewer()

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Update fields' })).toBeInTheDocument()
    })

    fireEvent.click(screen.getByRole('button', { name: 'Update fields' }))

    await waitFor(() => {
      expect(screen.getByRole('status')).toHaveTextContent('Updated 1 field.')
    })

    // Saving now serializes the RECALCULATED value, not the stale one.
    fireEvent.click(screen.getByRole('button', { name: 'Save' }))
    await waitFor(() => {
      expect(saveDocxMock).toHaveBeenCalledTimes(1)
    })
    const savedBundle = saveDocxMock.mock.calls[0]?.[0]
    const block = savedBundle?.document.sections[0].blocks[0]
    const field = block?.kind === 'paragraph' ? block.children[0] : undefined
    expect(field?.kind === 'field' && field.result).toEqual([
      { kind: 'run', children: [{ kind: 'text', value: 'Real Author' }] },
    ])
  })

  it('shows a status message when there are no fields to update', async () => {
    loadDocxMock.mockResolvedValue({ document: documentWithNoFields(), rawArchive: new Map() })
    renderViewer()

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Update fields' })).toBeInTheDocument()
    })

    fireEvent.click(screen.getByRole('button', { name: 'Update fields' }))

    await waitFor(() => {
      expect(screen.getByRole('status')).toHaveTextContent('No fields needed updating.')
    })
  })

  it('shows a status message when there is no table of contents to update', async () => {
    loadDocxMock.mockResolvedValue({ document: documentWithNoFields(), rawArchive: new Map() })
    renderViewer()

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Update table of contents' })).toBeInTheDocument()
    })

    fireEvent.click(screen.getByRole('button', { name: 'Update table of contents' }))

    await waitFor(() => {
      expect(screen.getByRole('status')).toHaveTextContent('No table of contents found to update.')
    })
  })

  it('dismisses the status message when its close button is clicked', async () => {
    loadDocxMock.mockResolvedValue({ document: documentWithNoFields(), rawArchive: new Map() })
    renderViewer()

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Update fields' })).toBeInTheDocument()
    })
    fireEvent.click(screen.getByRole('button', { name: 'Update fields' }))
    await waitFor(() => {
      expect(screen.getByRole('status')).toBeInTheDocument()
    })

    fireEvent.click(screen.getByRole('button', { name: 'Dismiss message' }))
    expect(screen.queryByRole('status')).not.toBeInTheDocument()
  })
})
