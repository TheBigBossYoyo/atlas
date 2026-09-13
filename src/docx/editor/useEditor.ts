import { useCallback, useRef, useState } from 'react'

import type { Document } from '../model'
import type { Range } from './commandTypes'
import { History } from './History'
import { handleBeforeInput, handleKeyDown } from './Input'

export type UseEditorResult = Readonly<{
  document: Document
  range: Range | null
  setRange: (r: Range | null) => void
  onBeforeInput: (e: InputEvent) => void
  onKeyDown: (e: KeyboardEvent) => void
}>

export function useEditor(initialDocument: Document): UseEditorResult {
  const [document, setDocument] = useState<Document>(initialDocument)
  const [range, setRange] = useState<Range | null>(null)
  const historyRef = useRef<History>(new History())

  const onBeforeInput = useCallback(
    (e: InputEvent) => {
      const result = handleBeforeInput(e, {
        document,
        range,
        history: historyRef.current,
      })
      if (result !== null) {
        e.preventDefault()
        setDocument(result.document)
        setRange(result.range)
      }
    },
    [document, range],
  )

  const onKeyDown = useCallback(
    (e: KeyboardEvent) => {
      const result = handleKeyDown(e, {
        document,
        range,
        history: historyRef.current,
      })
      if (result !== null) {
        e.preventDefault()
        setDocument(result.document)
        setRange(result.range)
      }
    },
    [document, range],
  )

  return {
    document,
    range,
    setRange,
    onBeforeInput,
    onKeyDown,
  }
}
