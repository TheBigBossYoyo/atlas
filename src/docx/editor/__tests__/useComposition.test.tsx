import { act, renderHook } from '@testing-library/react'
import { describe, expect, it } from 'vitest'

import { useComposition } from '../useComposition'

type CompositionEventShape = Pick<CompositionEvent, 'data' | 'type'>
type InputEventShape = Pick<InputEvent, 'data' | 'inputType' | 'type'>

function makeCompositionEvent(type: CompositionEvent['type'], data: string): CompositionEvent {
  const event: CompositionEventShape = { type, data }
  // jsdom may not expose a constructible CompositionEvent, so these tests use a typed plain-object cast.
  return event as unknown as CompositionEvent
}

function makeInputEvent(inputType: string, data: string | null): InputEvent {
  const event: InputEventShape = { type: 'beforeinput', inputType, data }
  return event as unknown as InputEvent
}

describe('useComposition', () => {
  it('starts inactive with beforeinput swallowing disabled', () => {
    const { result } = renderHook(() => useComposition())

    expect(result.current.state).toEqual({ active: false, text: '', startedAt: 0, appliedText: null })
    expect(result.current.shouldSwallow(makeInputEvent('insertCompositionText', 'é'))).toBe(false)
  })

  it('tracks composition updates and swallows matching beforeinput events', () => {
    const { result } = renderHook(() => useComposition())

    act(() => {
      result.current.handlers.onCompositionStart(makeCompositionEvent('compositionstart', ''))
      result.current.handlers.onCompositionUpdate(makeCompositionEvent('compositionupdate', '漢字'))
    })

    expect(result.current.state.active).toBe(true)
    expect(result.current.state.text).toBe('漢字')
    expect(result.current.shouldSwallow(makeInputEvent('insertCompositionText', '漢字'))).toBe(true)
    expect(result.current.shouldSwallow(makeInputEvent('insertText', '漢字'))).toBe(true)
  })

  it('returns commit text on composition end and resets state', () => {
    const { result } = renderHook(() => useComposition())
    let commitText: string | null = null

    act(() => {
      result.current.handlers.onCompositionStart(makeCompositionEvent('compositionstart', ''))
      result.current.handlers.onCompositionUpdate(makeCompositionEvent('compositionupdate', 'ç'))
      commitText = result.current.handlers.onCompositionEnd(makeCompositionEvent('compositionend', 'ç'))
    })

    expect(commitText).toBe('ç')
    expect(result.current.state).toEqual({ active: false, text: '', startedAt: 0, appliedText: null })
    expect(result.current.shouldSwallow(makeInputEvent('insertText', 'ç'))).toBe(false)
  })

  it('returns null for empty composition end commits', () => {
    const { result } = renderHook(() => useComposition())
    let commitText: string | null = 'pending'

    act(() => {
      result.current.handlers.onCompositionStart(makeCompositionEvent('compositionstart', ''))
      commitText = result.current.handlers.onCompositionEnd(makeCompositionEvent('compositionend', ''))
    })

    expect(commitText).toBeNull()
    expect(result.current.state).toEqual({ active: false, text: '', startedAt: 0, appliedText: null })
  })

  it('markApplied records the applied text so compositionend does not double-insert it (DXE-20)', () => {
    const { result } = renderHook(() => useComposition())
    let commitText: string | null = 'pending'

    act(() => {
      result.current.handlers.onCompositionStart(makeCompositionEvent('compositionstart', ''))
      result.current.handlers.onCompositionUpdate(makeCompositionEvent('compositionupdate', '漢字'))
      // Simulates a beforeinput(insertText) firing with the final composed
      // text before compositionend — the caller applies it to the model and
      // marks it applied.
      result.current.markApplied('漢字')
      commitText = result.current.handlers.onCompositionEnd(makeCompositionEvent('compositionend', '漢字'))
    })

    expect(commitText).toBeNull()
  })
})
