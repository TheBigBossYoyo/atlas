import { afterEach, describe, expect, it, vi } from 'vitest'

import {
  createCompositionState,
  markCompositionTextApplied,
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
      appliedText: null,
    })
  })

  it('onCompositionStart sets active state with empty text and current timestamp', () => {
    vi.spyOn(Date, 'now').mockReturnValue(1234)

    const nextState = onCompositionStart(
      { active: false, text: 'stale', startedAt: 10, appliedText: null },
      makeCompositionEvent('compositionstart', 'é'),
    )

    expect(nextState).toEqual({
      active: true,
      text: '',
      startedAt: 1234,
      appliedText: null,
    })
  })

  it('onCompositionUpdate stores the latest accumulated text from event.data', () => {
    const started = { active: true, text: '', startedAt: 22, appliedText: null }
    const updated = onCompositionUpdate(started, makeCompositionEvent('compositionupdate', 'é'))
    const accumulated = onCompositionUpdate(updated, makeCompositionEvent('compositionupdate', 'été'))

    expect(updated).toEqual({ active: true, text: 'é', startedAt: 22, appliedText: null })
    expect(accumulated).toEqual({ active: true, text: 'été', startedAt: 22, appliedText: null })
  })

  it('onCompositionEnd returns commitText and resets state', () => {
    const result = onCompositionEnd(
      { active: true, text: 'é', startedAt: 77, appliedText: null },
      makeCompositionEvent('compositionend', 'été'),
    )

    expect(result).toEqual({
      state: { active: false, text: '', startedAt: 0, appliedText: null },
      commitText: 'été',
    })
  })

  it('shouldSwallowBeforeInput returns true for insertCompositionText while active', () => {
    expect(
      shouldSwallowBeforeInput(
        { active: true, text: 'é', startedAt: 1, appliedText: null },
        makeInputEvent('insertCompositionText', 'é'),
      ),
    ).toBe(true)
  })

  it('shouldSwallowBeforeInput returns true for insertCompositionText variants while active', () => {
    expect(
      shouldSwallowBeforeInput(
        { active: true, text: '漢', startedAt: 1, appliedText: null },
        makeInputEvent('insertCompositionTextReplacement', '漢'),
      ),
    ).toBe(true)
  })

  it('shouldSwallowBeforeInput returns true for insertText matching the composition buffer', () => {
    expect(
      shouldSwallowBeforeInput(
        { active: true, text: 'ç', startedAt: 1, appliedText: null },
        makeInputEvent('insertText', 'ç'),
      ),
    ).toBe(true)
  })

  it('shouldSwallowBeforeInput returns false when composition is inactive', () => {
    expect(
      shouldSwallowBeforeInput(
        { active: false, text: 'é', startedAt: 1, appliedText: null },
        makeInputEvent('insertCompositionText', 'é'),
      ),
    ).toBe(false)
  })

  it('shouldSwallowBeforeInput returns false for non-text input types', () => {
    expect(
      shouldSwallowBeforeInput(
        { active: true, text: 'é', startedAt: 1, appliedText: null },
        makeInputEvent('deleteContentBackward', null),
      ),
    ).toBe(false)
  })

  it('onCompositionEnd with empty data returns an empty commit and inactive state', () => {
    const result = onCompositionEnd(
      { active: true, text: '', startedAt: 55, appliedText: null },
      makeCompositionEvent('compositionend', ''),
    )

    expect(result).toEqual({
      state: { active: false, text: '', startedAt: 0, appliedText: null },
      commitText: '',
    })
  })

  it('composition helpers do not mutate original state objects', () => {
    const original = { active: false, text: 'seed', startedAt: 9, appliedText: null }
    const started = onCompositionStart(original, makeCompositionEvent('compositionstart', 'e'))
    const active = { active: true, text: 'e', startedAt: 10, appliedText: null }
    const updated = onCompositionUpdate(active, makeCompositionEvent('compositionupdate', 'é'))
    const ended = onCompositionEnd(active, makeCompositionEvent('compositionend', 'é'))

    expect(original).toEqual({ active: false, text: 'seed', startedAt: 9, appliedText: null })
    expect(active).toEqual({ active: true, text: 'e', startedAt: 10, appliedText: null })
    expect(started).not.toBe(original)
    expect(updated).not.toBe(active)
    expect(ended.state).not.toBe(active)
  })

  // ─── DXE-20 — IME double-insert on candidate cycling ────────────────────────

  it('markCompositionTextApplied records applied text only while composition is active', () => {
    const active = { active: true, text: '漢字', startedAt: 1, appliedText: null }
    const marked = markCompositionTextApplied(active, '漢字')
    expect(marked).toEqual({ active: true, text: '漢字', startedAt: 1, appliedText: '漢字' })

    const inactive = { active: false, text: '', startedAt: 0, appliedText: null }
    expect(markCompositionTextApplied(inactive, 'x')).toBe(inactive)
  })

  it('onCompositionEnd skips re-committing text a beforeinput already applied', () => {
    const state = { active: true, text: '漢字', startedAt: 1, appliedText: '漢字' }
    const result = onCompositionEnd(state, makeCompositionEvent('compositionend', '漢字'))

    // The text already landed in the model via the earlier beforeinput —
    // committing it again here would double-insert it.
    expect(result.commitText).toBe('')
    expect(result.state).toEqual(createCompositionState())
  })

  it('onCompositionEnd still commits when the final text differs from what was already applied', () => {
    const state = { active: true, text: '漢字', startedAt: 1, appliedText: '漢' }
    const result = onCompositionEnd(state, makeCompositionEvent('compositionend', '漢字'))

    expect(result.commitText).toBe('漢字')
  })
})
