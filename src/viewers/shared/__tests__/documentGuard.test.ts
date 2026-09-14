/**
 * T9 (DAT-20) — RTF/ODT size cap + render-timeout guard.
 */
import { describe, expect, it, vi } from 'vitest'

import {
  DOCUMENT_SIZE_CAP_BYTES,
  DocumentRenderTimeoutError,
  formatSizeCapMessage,
  withRenderTimeout,
} from '../documentGuard'

describe('formatSizeCapMessage', () => {
  it('reports the file size and the cap in MB, mentioning the format label', () => {
    const message = formatSizeCapMessage(30 * 1024 * 1024, 'rtf')
    expect(message).toMatch(/too large to preview/i)
    expect(message).toContain('30.0 MB')
    expect(message).toContain('RTF')
  })
})

describe('withRenderTimeout', () => {
  it('resolves with the wrapped promise value when it settles before the timeout', async () => {
    const result = await withRenderTimeout(Promise.resolve('done'), 1000, 'too slow')
    expect(result).toBe('done')
  })

  it('rejects with the original error when the wrapped promise rejects first', async () => {
    await expect(withRenderTimeout(Promise.reject(new Error('boom')), 1000, 'too slow')).rejects.toThrow('boom')
  })

  it('rejects with a DocumentRenderTimeoutError once the budget elapses', async () => {
    vi.useFakeTimers()
    try {
      const neverResolves = new Promise<void>(() => {})
      const race = withRenderTimeout(neverResolves, 50, 'took too long')

      const assertion = expect(race).rejects.toBeInstanceOf(DocumentRenderTimeoutError)
      await vi.advanceTimersByTimeAsync(50)
      await assertion
    } finally {
      vi.useRealTimers()
    }
  })

  it('does not leave a dangling timer once the promise settles first', async () => {
    vi.useFakeTimers()
    try {
      await withRenderTimeout(Promise.resolve('fast'), 50, 'too slow')
      // If the timer were still pending, advancing past it and awaiting a
      // microtask flush would let a leftover rejection surface here.
      await vi.advanceTimersByTimeAsync(100)
    } finally {
      vi.useRealTimers()
    }
  })
})

describe('DOCUMENT_SIZE_CAP_BYTES', () => {
  it('is a sane, positive size (tens of MB)', () => {
    expect(DOCUMENT_SIZE_CAP_BYTES).toBeGreaterThan(1024 * 1024)
  })
})
