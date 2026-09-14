/**
 * T11 (DAT-16) — OdtViewer render/parse fixture coverage, plus T7 (tracked
 * changes visible + page count) and T9 (size cap) regression coverage.
 */
import { readFileSync } from 'node:fs'
import path from 'node:path'
import { renderHook, screen, waitFor } from '@testing-library/react'
import { describe, expect, it } from 'vitest'

import type { LoadedFile } from '../../formats/types'
import { ViewerProvider } from '../shared/ViewerContext'
import { useViewerStats } from '../shared/useViewerContext'
import { DOCUMENT_SIZE_CAP_BYTES } from '../shared/documentGuard'
import { OdtViewer } from '../OdtViewer'

const FIXTURES_DIR = path.resolve(process.cwd(), 'src/viewers/__fixtures__')
const CORPUS_DIR = path.join(FIXTURES_DIR, 'odt-corpus')

// A generous timeout for assertions that wait on OdtViewer's real (non-mocked)
// dynamic import of odf-kit/reader + dompurify — under full-suite worker
// contention that transform can occasionally take longer than the default
// 1000ms waitFor/findBy* timeout.
const IMPORT_TIMEOUT = { timeout: 10_000 }

function readArrayBuffer(filePath: string): ArrayBuffer {
  const buf = readFileSync(filePath)
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer
}

/** Mounts OdtViewer and returns a live `useViewerStats()` reading alongside it. */
function renderOdtWithStats(file: LoadedFile) {
  return renderHook(() => useViewerStats(), {
    wrapper: ({ children }) => (
      <ViewerProvider filePath={file.path}>
        <OdtViewer file={file} />
        {children}
      </ViewerProvider>
    ),
  })
}

describe('OdtViewer — fixture render/parse (T11)', () => {
  it('renders a real .odt fixture and reports word/page stats', async () => {
    const file: LoadedFile = {
      kind: 'binary',
      content: readArrayBuffer(path.join(FIXTURES_DIR, 'sample.odt')),
      path: '/tmp/sample.odt',
      format: 'odt',
    }

    const { result } = renderOdtWithStats(file)

    await screen.findByText(/Atlas ODT fixture/i, {}, IMPORT_TIMEOUT)
    await waitFor(() => expect(result.current).toMatchObject({ kind: 'document' }), IMPORT_TIMEOUT)
    expect(result.current).toMatchObject({ pages: 1 })
  })

  it('does not show the tracked-changes banner for a document with no tracked changes', async () => {
    const file: LoadedFile = {
      kind: 'binary',
      content: readArrayBuffer(path.join(FIXTURES_DIR, 'sample.odt')),
      path: '/tmp/sample.odt',
      format: 'odt',
    }

    renderOdtWithStats(file)
    await screen.findByText(/Atlas ODT fixture/i, {}, IMPORT_TIMEOUT)

    expect(screen.queryByRole('status')).not.toBeInTheDocument()
  })
})

describe('OdtViewer — tracked changes visible by default (T7/DAT-18)', () => {
  it('renders <ins>/<del> markup and a tracked-changes banner for the corpus fixture', async () => {
    const file: LoadedFile = {
      kind: 'binary',
      content: readArrayBuffer(path.join(CORPUS_DIR, 'tracked-changes.odt')),
      path: '/tmp/tracked-changes.odt',
      format: 'odt',
    }

    renderOdtWithStats(file)

    await screen.findByRole('status', {}, IMPORT_TIMEOUT)
    expect(document.querySelector('.odt-viewer__body ins')).not.toBeNull()
    expect(document.querySelector('.odt-viewer__body del')).not.toBeNull()
    expect(screen.getByText(/This paragraph was inserted by a reviewer\./)).toBeInTheDocument()
    expect(screen.getByText(/This paragraph was deleted\./)).toBeInTheDocument()
  })

  it('renders a table and lists from the corpus fixture', async () => {
    const file: LoadedFile = {
      kind: 'binary',
      content: readArrayBuffer(path.join(CORPUS_DIR, 'tables-and-lists.odt')),
      path: '/tmp/tables-and-lists.odt',
      format: 'odt',
    }

    renderOdtWithStats(file)
    await screen.findByText(/First bullet/i, {}, IMPORT_TIMEOUT)

    expect(document.querySelector('.odt-viewer__body table')).not.toBeNull()
    expect(document.querySelector('.odt-viewer__body ul')).not.toBeNull()
    expect(document.querySelector('.odt-viewer__body ol')).not.toBeNull()
  })
})

describe('OdtViewer — size cap (T9/DAT-20)', () => {
  it('refuses an oversized file with a friendly message instead of attempting to parse it', async () => {
    const oversized = new ArrayBuffer(DOCUMENT_SIZE_CAP_BYTES + 1)
    const file: LoadedFile = {
      kind: 'binary',
      content: oversized,
      path: '/tmp/huge.odt',
      format: 'odt',
    }

    renderOdtWithStats(file)

    expect(await screen.findByText(/too large to preview/i)).toBeInTheDocument()
  })
})
