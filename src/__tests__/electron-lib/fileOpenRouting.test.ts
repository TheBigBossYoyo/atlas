import { describe, expect, it } from 'vitest'

import { decideFileOpenAction } from '../../../electron/lib/fileOpenRouting.cjs'

describe('decideFileOpenAction (P5.2/ELEC-07)', () => {
  it('sends immediately once a window exists and the renderer has signalled it is listening', () => {
    expect(decideFileOpenAction({ hasWindow: true, rendererReady: true })).toBe('send')
  })

  it('queues when there is no window yet at all (the original ELEC-07 case)', () => {
    expect(decideFileOpenAction({ hasWindow: false, rendererReady: false })).toBe('queue')
  })

  it('queues when the window exists but the renderer has not mounted its listener yet', () => {
    // This is the narrower race the original ELEC-07 fix missed: `mainWindow`
    // is non-null as soon as `new BrowserWindow()` returns, well before the
    // renderer has run its boot-subscriptions effect.
    expect(decideFileOpenAction({ hasWindow: true, rendererReady: false })).toBe('queue')
  })

  it('queues (never sends) when there is somehow no window but the renderer flag is stale-true', () => {
    // Defensive case: `rendererReady` alone must never be trusted without a window.
    expect(decideFileOpenAction({ hasWindow: false, rendererReady: true })).toBe('queue')
  })
})
