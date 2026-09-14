/**
 * T2/DAT-07 — off-main-thread xlsx/ods parsing.
 *
 * jsdom (this project's vitest environment) has no `Worker` global, so these
 * tests exercise `useSpreadsheetWorkbook`'s synchronous fallback path for
 * real, and separately verify — via a minimal mock `Worker` — that a
 * large-enough buffer is dispatched to a Worker instead of being parsed
 * inline, which is what actually keeps the main thread unblocked in a real
 * browser/Electron renderer.
 */
import { renderHook, waitFor } from '@testing-library/react'
import { afterEach, describe, expect, it, vi } from 'vitest'
import * as XLSX from 'xlsx'

import { useSpreadsheetWorkbook } from '../useSpreadsheetWorkbook'
import { XLSX_WORKER_BYTE_THRESHOLD } from '../sizeThresholds'

function buildWorkbookBuffer(): ArrayBuffer {
  const ws = XLSX.utils.aoa_to_sheet([['Name', 'Score'], ['Alice', 10]])
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')
  return XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer
}

describe('useSpreadsheetWorkbook', () => {
  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('starts in the loading state before a buffer is provided', () => {
    const { result } = renderHook(() => useSpreadsheetWorkbook(null))
    expect(result.current.status).toBe('loading')
  })

  it('parses a small buffer synchronously (jsdom has no Worker global)', async () => {
    expect(typeof Worker).toBe('undefined')

    const buffer = buildWorkbookBuffer()
    const { result } = renderHook(() => useSpreadsheetWorkbook(buffer))

    await waitFor(() => expect(result.current.status).toBe('ready'))
    if (result.current.status !== 'ready') throw new Error('expected ready')
    expect(result.current.sheets).toHaveLength(1)
    expect(result.current.sheets[0].grid.rows[0]).toEqual(['Name', 'Score'])
  })

  it('reports an error state when the buffer cannot be parsed', async () => {
    // SheetJS is very permissive about plain garbage bytes (it happily
    // treats them as a one-cell text sheet), so this needs bytes that
    // actually fail parsing: a ZIP-magic ("PK") header followed by nonsense,
    // which XLSX.read recognizes as a corrupt xlsx/ods archive and throws on.
    const corruptZip = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10])
    const { result } = renderHook(() => useSpreadsheetWorkbook(corruptZip.buffer as ArrayBuffer))

    await waitFor(() => expect(result.current.status).toBe('error'))
  })

  it('dispatches a large buffer to a Worker instead of parsing it inline', async () => {
    const postedMessages: unknown[] = []
    const captured: { onMessage: ((event: MessageEvent) => void) | null } = { onMessage: null }

    class FakeWorker {
      onmessage: ((event: MessageEvent) => void) | null = null
      onerror: ((event: ErrorEvent) => void) | null = null
      postMessage(message: unknown) {
        postedMessages.push(message)
        captured.onMessage = (event) => this.onmessage?.(event)
      }
      terminate() {}
    }

    vi.stubGlobal('Worker', FakeWorker)

    // Pad the buffer past the worker threshold without needing a real 1MB
    // workbook fixture.
    const padded = new ArrayBuffer(XLSX_WORKER_BYTE_THRESHOLD + 1)

    const { result } = renderHook(() => useSpreadsheetWorkbook(padded))

    await waitFor(() => expect(postedMessages).toHaveLength(1))
    expect(result.current.status).toBe('loading')

    // Simulate the worker replying — the hook should transition to ready.
    captured.onMessage?.({ data: { ok: true, sheets: [] } } as MessageEvent)
    await waitFor(() => expect(result.current.status).toBe('ready'))
  })
})
