/**
 * VERSIONS-2 — the periodic version capture.
 *
 * VERSIONS-1 captured markdown only, from a string already sitting in App. Every
 * other format has to be re-serialised — a `.docx`/`.xlsx`/`.pptx` is re-zipped
 * from the in-memory model — which turned a synchronous, free call into an
 * asynchronous, expensive one. That is what these tests pin:
 *
 *   - an async source is awaited, not stored as a pending Promise;
 *   - two ticks never overlap, because one re-zip of a large document can
 *     outlast the two-minute interval and the user is typing on that thread;
 *   - a source that fails does not kill the timer for the rest of the session.
 *
 * All three are invisible when they regress: the history just quietly gets
 * thinner, or the app quietly gets slower.
 */
import { renderHook } from '@testing-library/react'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { AUTO_CAPTURE_INTERVAL_MS, useAutoCapture, type VersionHistoryState } from '../useVersionHistory'

function makeHistory(overrides: Partial<VersionHistoryState> = {}): VersionHistoryState {
  return {
    available: true,
    versions: [],
    loading: false,
    error: null,
    refresh: vi.fn(async () => {}),
    captureNow: vi.fn(async () => true),
    readVersion: vi.fn(async () => new Uint8Array()),
    clear: vi.fn(async () => {}),
    ...overrides,
  }
}

/** Advances past `count` capture intervals and lets every awaited step settle. */
async function tick(count = 1): Promise<void> {
  for (let i = 0; i < count; i += 1) {
    await vi.advanceTimersByTimeAsync(AUTO_CAPTURE_INTERVAL_MS)
  }
}

beforeEach(() => {
  vi.useFakeTimers()
})

afterEach(() => {
  vi.useRealTimers()
})

describe('useAutoCapture (VERSIONS-2)', () => {
  it('does not call the source on render, only when a capture is due', async () => {
    const getBytes = vi.fn(() => new Uint8Array([1]))
    const history = makeHistory()
    renderHook(() => useAutoCapture(history, true, getBytes))

    // Serialising is the expensive part; a timer that usually has nothing to do
    // must not pay for it on every render.
    expect(getBytes).not.toHaveBeenCalled()

    await tick()
    expect(getBytes).toHaveBeenCalledTimes(1)
    expect(history.captureNow).toHaveBeenCalledTimes(1)
  })

  it('awaits an async source and stores the resolved bytes', async () => {
    const bytes = new Uint8Array([0x50, 0x4b, 0x03, 0x04])
    const history = makeHistory()
    renderHook(() => useAutoCapture(history, true, async () => bytes))

    await tick()

    expect(history.captureNow).toHaveBeenCalledWith(bytes)
  })

  it('skips a tick while the previous capture is still serialising', async () => {
    // The case this guard exists for: re-zipping a large deck takes longer than
    // the interval. Overlapping ticks would stack several full re-serialisations
    // of the same document, all competing with the typing that triggered them.
    let release: (() => void) | null = null
    const getBytes = vi.fn(
      () =>
        new Promise<Uint8Array>((resolve) => {
          release = () => resolve(new Uint8Array([1]))
        }),
    )
    const history = makeHistory()
    renderHook(() => useAutoCapture(history, true, getBytes))

    await tick(3)
    // Three intervals passed, but the first capture never finished.
    expect(getBytes).toHaveBeenCalledTimes(1)
    expect(history.captureNow).not.toHaveBeenCalled()

    release!()
    await vi.advanceTimersByTimeAsync(0)
    expect(history.captureNow).toHaveBeenCalledTimes(1)

    // And once it is free, the next tick is taken normally.
    await tick()
    expect(getBytes).toHaveBeenCalledTimes(2)
  })

  it('keeps capturing after a source throws', async () => {
    const getBytes = vi
      .fn<() => Promise<Uint8Array>>()
      .mockRejectedValueOnce(new Error('the package could not be written'))
      .mockResolvedValue(new Uint8Array([2]))
    const history = makeHistory()
    renderHook(() => useAutoCapture(history, true, getBytes))

    await tick()
    expect(history.captureNow).not.toHaveBeenCalled()

    await tick()
    expect(history.captureNow).toHaveBeenCalledTimes(1)
  })

  it('stores nothing for a source with no bytes to give', async () => {
    // A viewer that registered no capture resolves `null`; an empty document
    // resolves zero bytes. Neither is a version worth recording.
    const history = makeHistory()
    const { unmount } = renderHook(() => useAutoCapture(history, true, async () => null))
    await tick()
    expect(history.captureNow).not.toHaveBeenCalled()
    unmount()

    const empty = makeHistory()
    renderHook(() => useAutoCapture(empty, true, async () => new Uint8Array()))
    await tick()
    expect(empty.captureNow).not.toHaveBeenCalled()
  })

  it('does nothing while disabled or while the store is unavailable', async () => {
    const getBytes = vi.fn(() => new Uint8Array([1]))

    const { unmount } = renderHook(() => useAutoCapture(makeHistory(), false, getBytes))
    await tick(2)
    expect(getBytes).not.toHaveBeenCalled()
    unmount()

    renderHook(() => useAutoCapture(makeHistory({ available: false }), true, getBytes))
    await tick(2)
    expect(getBytes).not.toHaveBeenCalled()
  })

  it('stops on unmount', async () => {
    const getBytes = vi.fn(() => new Uint8Array([1]))
    const { unmount } = renderHook(() => useAutoCapture(makeHistory(), true, getBytes))

    unmount()
    await tick(2)

    expect(getBytes).not.toHaveBeenCalled()
  })

  it('calls the latest source, not the one from the render that started the timer', async () => {
    // The timer effect deliberately does not depend on `getBytes` (re-creating
    // it on every keystroke would mean a capture never came due while someone
    // was typing), so the source is read through a ref. That only works if the
    // ref is actually kept current.
    const first = vi.fn(() => new Uint8Array([1]))
    const second = vi.fn(() => new Uint8Array([2]))
    const history = makeHistory()

    const { rerender } = renderHook(
      ({ getBytes }: { getBytes: () => Uint8Array }) => useAutoCapture(history, true, getBytes),
      { initialProps: { getBytes: first } },
    )

    rerender({ getBytes: second })
    await tick()

    expect(first).not.toHaveBeenCalled()
    expect(second).toHaveBeenCalledTimes(1)
  })
})
