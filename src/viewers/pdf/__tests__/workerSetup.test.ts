import { afterEach, describe, expect, it, vi } from 'vitest'

import { isFakeWorkerWarning, runDetectingFakeWorkerFallback } from '../workerSetup'

describe('isFakeWorkerWarning', () => {
  it('recognizes pdfjs-dist\'s exact fake-worker warning text', () => {
    expect(isFakeWorkerWarning('Warning: Setting up fake worker.')).toBe(true)
  })

  it('recognizes the warning even alongside other text', () => {
    expect(isFakeWorkerWarning('prefix Warning: Setting up fake worker. suffix')).toBe(true)
  })

  it('rejects unrelated strings and non-strings', () => {
    expect(isFakeWorkerWarning('Warning: something else.')).toBe(false)
    expect(isFakeWorkerWarning(42)).toBe(false)
    expect(isFakeWorkerWarning(null)).toBe(false)
    expect(isFakeWorkerWarning(undefined)).toBe(false)
  })
})

describe('runDetectingFakeWorkerFallback', () => {
  const originalWarn = console.warn

  afterEach(() => {
    console.warn = originalWarn
  })

  it('reports usedFakeWorker: false when the task never warns', async () => {
    const { result, usedFakeWorker } = await runDetectingFakeWorkerFallback(async () => 'ok')
    expect(result).toBe('ok')
    expect(usedFakeWorker).toBe(false)
  })

  it('detects the fallback warning logged during the task', async () => {
    const { usedFakeWorker } = await runDetectingFakeWorkerFallback(async () => {
      console.warn('Warning: Setting up fake worker.')
      return null
    })
    expect(usedFakeWorker).toBe(true)
  })

  it('always restores the original console.warn, even when the task throws', async () => {
    await expect(
      runDetectingFakeWorkerFallback(async () => {
        throw new Error('boom')
      }),
    ).rejects.toThrow('boom')

    expect(console.warn).toBe(originalWarn)
  })

  it('still forwards warnings to the original console.warn', async () => {
    const spy = vi.fn()
    console.warn = spy

    await runDetectingFakeWorkerFallback(async () => {
      console.warn('Warning: Setting up fake worker.')
    })

    expect(spy).toHaveBeenCalledWith('Warning: Setting up fake worker.')
  })
})
