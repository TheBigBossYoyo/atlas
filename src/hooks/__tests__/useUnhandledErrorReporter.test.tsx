/**
 * Renderer-side last-resort error reporting — see the hook's own header for why
 * the renderer needed this when `electron/main.cjs` has had process-level
 * `uncaughtException`/`unhandledRejection` handlers all along.
 *
 * The case worth guarding hardest is the false positive: markdown regularly
 * references remote images (the production CSP deliberately still allows
 * `https:` for exactly that reason), and a dead image URL fires an `error`
 * event at the window. If that raised a toast, opening a README with one stale
 * badge would accuse Atlas of a failure it did not have.
 */
import { act, render } from '@testing-library/react'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { useUnhandledErrorReporter } from '../useUnhandledErrorReporter'

function Harness({
  showToast,
  describe: describeDetail = (detail: string) => `Something went wrong: ${detail}`,
}: {
  showToast: (message: string, variant?: 'error' | 'success' | 'info') => void
  describe?: (detail: string) => string
}) {
  useUnhandledErrorReporter(showToast, describeDetail)
  return null
}

/**
 * jsdom does not implement `PromiseRejectionEvent`, and a real unhandled
 * rejection in a test would fail the run rather than reach a listener — so the
 * event the browser would dispatch is dispatched directly, which is exactly the
 * contract the hook is written against.
 */
function rejectWith(reason: unknown): void {
  const event = new Event('unhandledrejection') as Event & { reason?: unknown }
  event.reason = reason
  window.dispatchEvent(event)
}

function throwWith(error: unknown, target: EventTarget | null = null): void {
  const event = new Event('error') as Event & { error?: unknown; message?: string }
  event.error = error
  event.message = error instanceof Error ? error.message : String(error)
  if (target !== null) {
    target.dispatchEvent(event)
    return
  }
  window.dispatchEvent(event)
}

let consoleError: ReturnType<typeof vi.spyOn>

beforeEach(() => {
  consoleError = vi.spyOn(console, 'error').mockImplementation(() => {})
})

afterEach(() => {
  vi.restoreAllMocks()
  vi.useRealTimers()
})

describe('useUnhandledErrorReporter', () => {
  it('reports an unhandled rejection through the toast queue and the console', () => {
    const showToast = vi.fn()
    render(<Harness showToast={showToast} />)

    act(() => rejectWith(new Error('writeOfficePackage failed')))

    expect(showToast).toHaveBeenCalledWith('Something went wrong: writeOfficePackage failed', 'error')
    expect(consoleError).toHaveBeenCalled()
  })

  it('reports an uncaught error the same way', () => {
    const showToast = vi.fn()
    render(<Harness showToast={showToast} />)

    act(() => throwWith(new Error('Cannot read properties of null')))

    expect(showToast).toHaveBeenCalledWith('Something went wrong: Cannot read properties of null', 'error')
  })

  it('formats a rejection that carries something other than an Error', () => {
    const showToast = vi.fn()
    render(<Harness showToast={showToast} />)

    act(() => rejectWith('plain string reason'))
    act(() => rejectWith({ code: 500 }))

    expect(showToast).toHaveBeenNthCalledWith(1, 'Something went wrong: plain string reason', 'error')
    expect(showToast).toHaveBeenNthCalledWith(2, 'Something went wrong: {"code":500}', 'error')
  })

  it('survives a reason that cannot be stringified as JSON', () => {
    const showToast = vi.fn()
    render(<Harness showToast={showToast} />)

    const circular: Record<string, unknown> = {}
    circular.self = circular

    act(() => rejectWith(circular))

    expect(showToast).toHaveBeenCalledTimes(1)
    expect(showToast.mock.calls[0][0]).toContain('Something went wrong:')
  })

  it('ignores a failed subresource load, so a dead image in a document raises nothing', () => {
    const showToast = vi.fn()
    render(<Harness showToast={showToast} />)

    const img = document.createElement('img')
    document.body.appendChild(img)
    // What a broken <img src="https://..."> actually dispatches: an `error`
    // event on the ELEMENT, with no `error` property on it.
    act(() => {
      const event = new Event('error', { bubbles: false })
      img.dispatchEvent(event)
    })

    expect(showToast).not.toHaveBeenCalled()
    img.remove()
  })

  it('still reports a real throw that happens to originate at an element', () => {
    const showToast = vi.fn()
    render(<Harness showToast={showToast} />)

    const div = document.createElement('div')
    document.body.appendChild(div)
    act(() => throwWith(new Error('handler blew up'), div))

    expect(showToast).toHaveBeenCalledWith('Something went wrong: handler blew up', 'error')
    div.remove()
  })

  it('suppresses a repeat of the same message inside the dedupe window, but keeps logging it', () => {
    const showToast = vi.fn()
    render(<Harness showToast={showToast} />)

    act(() => rejectWith(new Error('same failure')))
    act(() => rejectWith(new Error('same failure')))
    act(() => rejectWith(new Error('same failure')))

    // One toast for the user; every occurrence still in the console.
    expect(showToast).toHaveBeenCalledTimes(1)
    expect(consoleError).toHaveBeenCalledTimes(3)
  })

  it('reports the same message again once the dedupe window has passed', () => {
    vi.useFakeTimers()
    const showToast = vi.fn()
    render(<Harness showToast={showToast} />)

    act(() => rejectWith(new Error('recurring failure')))
    act(() => {
      vi.advanceTimersByTime(6000)
    })
    act(() => rejectWith(new Error('recurring failure')))

    expect(showToast).toHaveBeenCalledTimes(2)
  })

  it('reports distinct failures separately', () => {
    const showToast = vi.fn()
    render(<Harness showToast={showToast} />)

    act(() => rejectWith(new Error('first')))
    act(() => rejectWith(new Error('second')))

    expect(showToast).toHaveBeenCalledTimes(2)
  })

  it('caps how many toasts one session can produce, without capping the log', () => {
    const showToast = vi.fn()
    render(<Harness showToast={showToast} />)

    for (let i = 0; i < 15; i += 1) {
      act(() => rejectWith(new Error(`distinct failure ${i}`)))
    }

    expect(showToast).toHaveBeenCalledTimes(10)
    expect(consoleError).toHaveBeenCalledTimes(15)
  })

  it('stops listening once unmounted', () => {
    const showToast = vi.fn()
    const { unmount } = render(<Harness showToast={showToast} />)

    unmount()
    act(() => rejectWith(new Error('after unmount')))

    expect(showToast).not.toHaveBeenCalled()
  })

  it('uses the caller-supplied formatter, so the message comes from the catalogue', () => {
    const showToast = vi.fn()
    render(<Harness showToast={showToast} describe={(detail) => `Une erreur est survenue : ${detail}`} />)

    act(() => rejectWith(new Error('échec')))

    expect(showToast).toHaveBeenCalledWith('Une erreur est survenue : échec', 'error')
  })
})
