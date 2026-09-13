import { act, renderHook } from '@testing-library/react'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import type { SpellCheckContextMenuPayload } from '../../../electron'
import { useSpellCheck } from '../useSpellCheck'

type Listener = (payload: SpellCheckContextMenuPayload) => void

interface MockBridge {
  readonly onSpellCheckMenu: (cb: Listener) => () => void
  readonly replaceMisspelling: (word: string) => Promise<{ replaced: boolean }>
  readonly addWordToDictionary: (word: string) => Promise<{ added: boolean }>
}

interface MockState {
  listeners: Set<Listener>
  unsubscribed: number
  replaceMisspelling: ReturnType<typeof vi.fn>
  addWord: ReturnType<typeof vi.fn>
}

function installMockBridge(): MockState {
  const state: MockState = {
    listeners: new Set(),
    unsubscribed: 0,
    replaceMisspelling: vi.fn(async (_word: string) => ({ replaced: true })),
    addWord: vi.fn(async (_word: string) => ({ added: true })),
  }
  const bridge: MockBridge = {
    onSpellCheckMenu: (cb) => {
      state.listeners.add(cb)
      return () => {
        state.listeners.delete(cb)
        state.unsubscribed += 1
      }
    },
    replaceMisspelling: state.replaceMisspelling as unknown as MockBridge['replaceMisspelling'],
    addWordToDictionary: state.addWord as unknown as MockBridge['addWordToDictionary'],
  }
  ;(window as unknown as { electronAPI?: MockBridge }).electronAPI = {
    ...bridge,
  }
  return state
}

describe('useSpellCheck', () => {
  let mock: MockState

  beforeEach(() => {
    mock = installMockBridge()
  })

  afterEach(() => {
    delete (window as unknown as { electronAPI?: unknown }).electronAPI
  })

  it('starts with null state', () => {
    const { result } = renderHook(() => useSpellCheck())
    expect(result.current.state).toBeNull()
  })

  it('captures context-menu payloads from the bridge', () => {
    const { result } = renderHook(() => useSpellCheck())
    act(() => {
      mock.listeners.forEach((cb) =>
        cb({ word: 'teh', suggestions: ['the', 'tech'], x: 10, y: 20 }),
      )
    })
    expect(result.current.state).toEqual({
      word: 'teh',
      suggestions: ['the', 'tech'],
      x: 10,
      y: 20,
    })
  })

  it('dismiss clears state', () => {
    const { result } = renderHook(() => useSpellCheck())
    act(() => {
      mock.listeners.forEach((cb) => cb({ word: 'a', suggestions: [], x: 0, y: 0 }))
    })
    act(() => result.current.dismiss())
    expect(result.current.state).toBeNull()
  })

  it('addToDictionary forwards to the bridge and returns true on success', async () => {
    const { result } = renderHook(() => useSpellCheck())
    const ok = await result.current.addToDictionary('atlas')
    expect(ok).toBe(true)
    expect(mock.addWord).toHaveBeenCalledWith('atlas')
  })

  it('replaceMisspelling forwards to the bridge and clears active state', async () => {
    const { result } = renderHook(() => useSpellCheck())
    act(() => {
      mock.listeners.forEach((cb) => cb({ word: 'teh', suggestions: ['the'], x: 1, y: 2 }))
    })
    let ok = false
    await act(async () => {
      ok = await result.current.replaceMisspelling('the')
    })
    expect(ok).toBe(true)
    expect(mock.replaceMisspelling).toHaveBeenCalledWith('the')
    expect(result.current.state).toBeNull()
  })

  it('unsubscribes on unmount', () => {
    const { unmount } = renderHook(() => useSpellCheck())
    unmount()
    expect(mock.unsubscribed).toBe(1)
  })

  it('is inert when electronAPI is undefined', async () => {
    delete (window as unknown as { electronAPI?: unknown }).electronAPI
    const { result } = renderHook(() => useSpellCheck())
    expect(result.current.state).toBeNull()
    const ok = await result.current.addToDictionary('x')
    expect(ok).toBe(false)
  })
})
