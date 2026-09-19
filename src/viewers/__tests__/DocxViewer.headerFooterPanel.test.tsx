/**
 * DocxViewer — header/footer panel keyboard/focus affordances (UX).
 *
 * Unlike every other toggleable dialog/panel in the shell (ShortcutsModal,
 * UnsavedChangesDialog, ExportMenu/NewDocumentMenu/ThemeMenu, FindReplace),
 * this panel had none: opening it left focus on the toolbar toggle button,
 * Escape did nothing while focused inside it, and closing it (by any means)
 * never returned focus to the toggle. Mirrors DocxViewer.fields.test.tsx's
 * mock setup (mocking only `docx`'s load/save, `docx/layout`'s `paginate`,
 * and the `PageStack` renderer) so the real header/footer editing module is
 * exercised end to end.
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

function documentWithHeaderFooter(): DocxDocument {
  return {
    kind: 'document',
    sections: [
      {
        kind: 'section',
        props: {
          headerReference: [{ id: 'rId4', type: 'default' }],
          footerReference: [{ id: 'rId5', type: 'default' }],
        },
        blocks: [{ kind: 'paragraph', children: [{ kind: 'run', children: [{ kind: 'text', value: 'Body text' }] }] }],
      },
    ],
    styles: new Map(),
    numbering: new Map(),
    comments: new Map(),
    footnotes: new Map(),
    endnotes: new Map(),
    headers: new Map([
      ['rId4', { kind: 'header', id: 'rId4', blocks: [{ kind: 'paragraph', children: [{ kind: 'run', children: [{ kind: 'text', value: 'Quarterly report' }] }] }] }],
    ]),
    footers: new Map([
      ['rId5', { kind: 'footer', id: 'rId5', blocks: [{ kind: 'paragraph', children: [{ kind: 'run', children: [{ kind: 'text', value: 'Page 1' }] }] }] }],
    ]),
  } as unknown as DocxDocument
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

describe('DocxViewer header/footer panel', () => {
  beforeEach(() => {
    loadDocxMock.mockResolvedValue({ document: documentWithHeaderFooter(), rawArchive: new Map() })
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

  async function openPanel(): Promise<HTMLElement> {
    renderViewer()
    const toggle = await waitFor(() => screen.getByRole('button', { name: 'Header and footer' }))
    toggle.focus()
    fireEvent.click(toggle)
    return screen.getByRole('group', { name: 'Header and footer' })
  }

  it('moves focus into the panel when opened', async () => {
    await openPanel()
    expect(screen.getByRole('button', { name: 'Close header and footer' })).toHaveFocus()
  })

  it('closes on Escape while focus is inside the panel', async () => {
    const panel = await openPanel()
    expect(panel).toBeInTheDocument()

    fireEvent.keyDown(panel, { key: 'Escape' })

    expect(screen.queryByRole('group', { name: 'Header and footer' })).not.toBeInTheDocument()
  })

  it('restores focus to the toolbar toggle once the panel closes', async () => {
    await openPanel()
    const toggle = screen.getByRole('button', { name: 'Header and footer' })

    fireEvent.click(screen.getByRole('button', { name: 'Close header and footer' }))

    expect(screen.queryByRole('group', { name: 'Header and footer' })).not.toBeInTheDocument()
    expect(toggle).toHaveFocus()
  })
})
