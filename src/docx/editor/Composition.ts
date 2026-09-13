export interface CompositionState {
  active: boolean
  text: string
  startedAt: number
}

export function createCompositionState(): CompositionState {
  return {
    active: false,
    text: '',
    startedAt: 0,
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
  const commitText = event.data ?? state.text

  return {
    state: createCompositionState(),
    commitText,
  }
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
