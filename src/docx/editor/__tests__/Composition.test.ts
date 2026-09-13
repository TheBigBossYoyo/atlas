import { afterEach, describe, expect, it, vi } from 'vitest'

import {
  createCompositionState,
  onCompositionEnd,
  onCompositionStart,
  onCompositionUpdate,
  shouldSwallowBeforeInput,
} from '../Composition'

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

describe('Composition', () => {
  afterEach(() => {
    vi.restoreAllMocks()
  })

  it('createCompositionState returns an inactive empty state', () => {
    expect(createCompositionState()).toEqual({
      active: false,
      text: '',
      startedAt: 0,
    })
  })

  it('onCompositionStart sets active state with empty text and current timestamp', () => {
    vi.spyOn(Date, 'now').mockReturnValue(1234)

    const nextState = onCompositionStart(
      { active: false, text: 'stale', startedAt: 10 },
      makeCompositionEvent('compositionstart', 'é'),
    )

    expect(nextState).toEqual({
      active: true,
      text: '',
      startedAt: 1234,
    })
  })

  it('onCompositionUpdate stores the latest accumulated text from event.data', () => {
    const started = { active: true, text: '', startedAt: 22 }
    const updated = onCompositionUpdate(started, makeCompositionEvent('compositionupdate', 'é'))
    const accumulated = onCompositionUpdate(updated, makeCompositionEvent('compositionupdate', 'été'))

    expect(updated).toEqual({ active: true, text: 'é', startedAt: 22 })
    expect(accumulated).toEqual({ active: true, text: 'été', startedAt: 22 })
  })

  it('onCompositionEnd returns commitText and resets state', () => {
    const result = onCompositionEnd(
      { active: true, text: 'é', startedAt: 77 },
      makeCompositionEvent('compositionend', 'été'),
    )

    expect(result).toEqual({
      state: { active: false, text: '', startedAt: 0 },
      commitText: 'été',
    })
  })

  it('shouldSwallowBeforeInput returns true for insertCompositionText while active', () => {
    expect(
      shouldSwallowBeforeInput(
        { active: true, text: 'é', startedAt: 1 },
        makeInputEvent('insertCompositionText', 'é'),
      ),
    ).toBe(true)
  })

  it('shouldSwallowBeforeInput returns true for insertCompositionText variants while active', () => {
    expect(
      shouldSwallowBeforeInput(
        { active: true, text: '漢', startedAt: 1 },
        makeInputEvent('insertCompositionTextReplacement', '漢'),
      ),
    ).toBe(true)
  })

  it('shouldSwallowBeforeInput returns true for insertText matching the composition buffer', () => {
    expect(
      shouldSwallowBeforeInput(
        { active: true, text: 'ç', startedAt: 1 },
        makeInputEvent('insertText', 'ç'),
      ),
    ).toBe(true)
  })

  it('shouldSwallowBeforeInput returns false when composition is inactive', () => {
    expect(
      shouldSwallowBeforeInput(
        { active: false, text: 'é', startedAt: 1 },
        makeInputEvent('insertCompositionText', 'é'),
      ),
    ).toBe(false)
  })

  it('shouldSwallowBeforeInput returns false for non-text input types', () => {
    expect(
      shouldSwallowBeforeInput(
        { active: true, text: 'é', startedAt: 1 },
        makeInputEvent('deleteContentBackward', null),
      ),
    ).toBe(false)
  })

  it('onCompositionEnd with empty data returns an empty commit and inactive state', () => {
    const result = onCompositionEnd(
      { active: true, text: '', startedAt: 55 },
      makeCompositionEvent('compositionend', ''),
    )

    expect(result).toEqual({
      state: { active: false, text: '', startedAt: 0 },
      commitText: '',
    })
  })

  it('composition helpers do not mutate original state objects', () => {
    const original = { active: false, text: 'seed', startedAt: 9 }
    const started = onCompositionStart(original, makeCompositionEvent('compositionstart', 'e'))
    const active = { active: true, text: 'e', startedAt: 10 }
    const updated = onCompositionUpdate(active, makeCompositionEvent('compositionupdate', 'é'))
    const ended = onCompositionEnd(active, makeCompositionEvent('compositionend', 'é'))

    expect(original).toEqual({ active: false, text: 'seed', startedAt: 9 })
    expect(active).toEqual({ active: true, text: 'e', startedAt: 10 })
    expect(started).not.toBe(original)
    expect(updated).not.toBe(active)
    expect(ended.state).not.toBe(active)
  })
})
