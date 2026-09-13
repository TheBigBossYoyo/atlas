import { useCallback, useRef, useState } from 'react'

import {
  createCompositionState,
  onCompositionEnd as endComposition,
  onCompositionStart as startComposition,
  onCompositionUpdate as updateComposition,
  shouldSwallowBeforeInput,
} from './Composition'
import type { CompositionState } from './Composition'

type CompositionHandlers = Readonly<{
  onCompositionStart: (event: CompositionEvent) => void
  onCompositionUpdate: (event: CompositionEvent) => void
  onCompositionEnd: (event: CompositionEvent) => string | null
}>

export type UseCompositionResult = Readonly<{
  state: CompositionState
  handlers: CompositionHandlers
  shouldSwallow: (event: InputEvent) => boolean
}>

export function useComposition(): UseCompositionResult {
  const [state, setState] = useState<CompositionState>(() => createCompositionState())
  const stateRef = useRef(state)

  const onCompositionStart = useCallback((event: CompositionEvent) => {
    const nextState = startComposition(stateRef.current, event)
    stateRef.current = nextState
    setState(nextState)
  }, [])

  const onCompositionUpdate = useCallback((event: CompositionEvent) => {
    const nextState = updateComposition(stateRef.current, event)
    stateRef.current = nextState
    setState(nextState)
  }, [])

  const onCompositionEnd = useCallback((event: CompositionEvent) => {
    const result = endComposition(stateRef.current, event)
    stateRef.current = result.state
    setState(result.state)

    return result.commitText.length > 0 ? result.commitText : null
  }, [])

  const shouldSwallow = useCallback((event: InputEvent) => {
    return shouldSwallowBeforeInput(stateRef.current, event)
  }, [])

  return {
    state,
    handlers: {
      onCompositionStart,
      onCompositionUpdate,
      onCompositionEnd,
    },
    shouldSwallow,
  }
}
