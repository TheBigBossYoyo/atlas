export interface CompositionState {
  active: boolean
  text: string
  startedAt: number
  /**
   * DXE-20 — some browsers fire a `beforeinput(insertText)` carrying the
   * fully-composed text right before `compositionend`, rather than (or in
   * addition to) an `insertCompositionText` event. When that beforeinput's
   * text doesn't match the last `compositionupdate` snapshot in `text`,
   * `shouldSwallowBeforeInput` lets it through and the caller applies it to
   * the model — so by the time `compositionend` fires, the text has already
   * landed once. `appliedText` records exactly what was applied that way so
   * `onCompositionEnd` can recognize the duplicate and skip re-committing it,
   * instead of inserting the same candidate a second time.
   */
  appliedText: string | null
}

export function createCompositionState(): CompositionState {
  return {
    active: false,
    text: '',
    startedAt: 0,
    appliedText: null,
  }
}

export function onCompositionStart(
  state: CompositionState,
  event: CompositionEvent,
): CompositionState {
  void state
  void event

  return {
    active: true,
    text: '',
    startedAt: Date.now(),
    appliedText: null,
  }
}

export function onCompositionUpdate(
  state: CompositionState,
  event: CompositionEvent,
): CompositionState {
  return {
    ...state,
    text: event.data ?? '',
  }
}

export function onCompositionEnd(
  state: CompositionState,
  event: CompositionEvent,
): { state: CompositionState; commitText: string } {
  const finalText = event.data ?? state.text
  const alreadyApplied = state.appliedText !== null && state.appliedText === finalText

  return {
    state: createCompositionState(),
    commitText: alreadyApplied ? '' : finalText,
  }
}

/**
 * Marks `text` as already applied to the model via a non-swallowed
 * `beforeinput(insertText)` that fired while composition was still active
 * (see `appliedText` above). A no-op when composition isn't active, so a
 * stray call outside composition can't leave stale state behind.
 */
export function markCompositionTextApplied(
  state: CompositionState,
  text: string,
): CompositionState {
  if (!state.active) {
    return state
  }

  return { ...state, appliedText: text }
}

export function shouldSwallowBeforeInput(
  state: CompositionState,
  event: InputEvent,
): boolean {
  if (!state.active) {
    return false
  }

  if (event.inputType.startsWith('insertCompositionText')) {
    return true
  }

  if (event.inputType !== 'insertText') {
    return false
  }

  return state.text.length > 0 && event.data === state.text
}
