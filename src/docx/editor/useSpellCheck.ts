import { useCallback, useEffect, useMemo, useState } from 'react'

import type { SpellCheckContextMenuPayload } from '../../electron'

export interface SpellCheckState {
  readonly word: string
  readonly suggestions: ReadonlyArray<string>
  readonly x: number
  readonly y: number
}

export interface UseSpellCheckResult {
  readonly state: SpellCheckState | null
  readonly dismiss: () => void
  readonly replaceMisspelling: (word: string) => Promise<boolean>
  readonly addToDictionary: (word: string) => Promise<boolean>
}

function getSpellcheckBridge(): NonNullable<Window['electronAPI']>['spellcheck'] | undefined {
  if (typeof window === 'undefined') {
    return undefined
  }
  return window.electronAPI?.spellcheck
}

export function useSpellCheck(): UseSpellCheckResult {
  const [state, setState] = useState<SpellCheckState | null>(null)

  useEffect(() => {
    const bridge = getSpellcheckBridge()
    if (bridge === undefined || typeof bridge.onContextMenu !== 'function') {
      return
    }

    const unsubscribe = bridge.onContextMenu((payload: SpellCheckContextMenuPayload) => {
      setState({
        word: payload.word,
        suggestions: payload.suggestions,
        x: payload.x,
        y: payload.y,
      })
    })

    return unsubscribe
  }, [])

  const dismiss = useCallback(() => {
    setState(null)
  }, [])

  const replaceMisspelling = useCallback(async (word: string): Promise<boolean> => {
    const bridge = getSpellcheckBridge()
    if (bridge === undefined || typeof bridge.replaceMisspelling !== 'function') {
      return false
    }

    try {
      const result = await bridge.replaceMisspelling(word)
      setState(null)
      return result.replaced === true
    } catch {
      return false
    }
  }, [])

  const addToDictionary = useCallback(async (word: string): Promise<boolean> => {
    const bridge = getSpellcheckBridge()
    if (bridge === undefined || typeof bridge.addWord !== 'function') {
      return false
    }
    try {
      const result = await bridge.addWord(word)
      setState(null)
      return result.added === true
    } catch {
      return false
    }
  }, [])

  return useMemo(
    () => ({ state, dismiss, replaceMisspelling, addToDictionary }),
    [addToDictionary, dismiss, replaceMisspelling, state],
  )
}
