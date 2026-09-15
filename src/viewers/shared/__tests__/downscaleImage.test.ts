/**
 * S18/SLD-19 — thumbnail image downscaling.
 *
 * jsdom (this project's Vitest environment) implements neither
 * `createImageBitmap` nor `OffscreenCanvas`, so every case here exercises
 * the feature-detected fallback path — the real resize path only runs in a
 * genuine browser/Electron renderer. The fallback itself (never rejecting,
 * always resolving to a usable src, caching per key) is exactly the
 * contract callers depend on, so it is worth covering directly.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'

import { getDownscaledImage } from '../downscaleImage'

const TINY_PNG_DATA_URL =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='

afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
})

describe('getDownscaledImage', () => {
  it('resolves to the original src unchanged when createImageBitmap/OffscreenCanvas are unavailable (jsdom)', async () => {
    expect(typeof globalThis.createImageBitmap).not.toBe('function')

    const result = await getDownscaledImage(TINY_PNG_DATA_URL, 32)
    expect(result).toBe(TINY_PNG_DATA_URL)
  })

  it('resolves to the original src unchanged for a non-data: URL (nothing to safely re-fetch/resize)', async () => {
    const result = await getDownscaledImage('https://example.com/pic.png', 32)
    expect(result).toBe('https://example.com/pic.png')
  })

  it('returns the same promise for a repeated (src, size) request instead of recomputing', async () => {
    const first = getDownscaledImage(TINY_PNG_DATA_URL, 64)
    const second = getDownscaledImage(TINY_PNG_DATA_URL, 64)
    expect(first).toBe(second)
    await expect(first).resolves.toBe(TINY_PNG_DATA_URL)
  })

  it('treats different target sizes for the same src as separate cache entries', async () => {
    const at32 = getDownscaledImage(TINY_PNG_DATA_URL, 32)
    const at64 = getDownscaledImage(TINY_PNG_DATA_URL, 64)
    expect(at32).not.toBe(at64)
  })

  it('never rejects even if something downstream throws', async () => {
    // Simulate a runtime that reports support but actually fails, proving
    // the promise still resolves (to the original src) instead of
    // propagating the error to the caller.
    vi.stubGlobal('createImageBitmap', vi.fn().mockRejectedValue(new Error('decode failed')))
    vi.stubGlobal('OffscreenCanvas', class {} as unknown as typeof OffscreenCanvas)
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue({ blob: () => Promise.resolve(new Blob()) }))

    await expect(getDownscaledImage('data:image/png;base64,xx', 16)).resolves.toBe(
      'data:image/png;base64,xx',
    )
  })
})
