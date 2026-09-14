/**
 * T11 (DAT-16) — RtfViewer render/parse fixture coverage, plus T7 (page
 * count) and T9 (size cap) regression coverage.
 */
import { readFileSync } from 'node:fs'
import path from 'node:path'
import { render, renderHook, waitFor } from '@testing-library/react'
import { describe, expect, it } from 'vitest'

import type { LoadedFile } from '../../formats/types'
import { ViewerProvider } from '../shared/ViewerContext'
import { useViewerStats } from '../shared/useViewerContext'
import { DOCUMENT_SIZE_CAP_BYTES } from '../shared/documentGuard'
import { RtfViewer } from '../RtfViewer'

const FIXTURES_DIR = path.resolve(process.cwd(), 'src/viewers/__fixtures__')

function readFixtureArrayBuffer(name: string): ArrayBuffer {
  const buf = readFileSync(path.join(FIXTURES_DIR, name))
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer
}

describe('RtfViewer — fixture render/parse (T11)', () => {
  it('renders a real .rtf fixture and reports word/page stats', async () => {
    const file: LoadedFile = {
      kind: 'binary',
      content: readFixtureArrayBuffer('sample.rtf'),
      path: '/tmp/sample.rtf',
      format: 'rtf',
    }

    const { result } = renderHook(() => useViewerStats(), {
      wrapper: ({ children }) => (
        <ViewerProvider filePath={file.path}>
          <RtfViewer file={file} />
          {children}
        </ViewerProvider>
      ),
    })

    await waitFor(() => expect(result.current).toMatchObject({ kind: 'document' }))
    // jsdom lays out nothing (scrollHeight is always 0), so the page-count
    // fallback in pageEstimate.ts always resolves to 1 here — the actual
    // height-driven multi-page math is unit-tested directly in
    // pageEstimate.test.ts instead of depending on jsdom layout.
    expect(result.current).toMatchObject({ pages: 1 })
  })
})

describe('RtfViewer — size cap (T9/DAT-20)', () => {
  it('refuses an oversized file with a friendly message instead of attempting to parse it', async () => {
    const oversized = new ArrayBuffer(DOCUMENT_SIZE_CAP_BYTES + 1)
    const file: LoadedFile = {
      kind: 'binary',
      content: oversized,
      path: '/tmp/huge.rtf',
      format: 'rtf',
    }

    const { findByText } = render(
      <ViewerProvider filePath={file.path}>
        <RtfViewer file={file} />
      </ViewerProvider>,
    )

    expect(await findByText(/too large to preview/i)).toBeInTheDocument()
  })
})
