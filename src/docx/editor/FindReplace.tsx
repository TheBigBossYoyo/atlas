import { useCallback, useEffect, useRef, useState } from 'react'

import { ChevronDown, ChevronUp, Replace, Search, X } from 'lucide-react'

import type { FindOptions } from './Find'
import './__styles__/find-replace.css'

// ---------------------------------------------------------------------------
// Props
// ---------------------------------------------------------------------------

export interface FindReplaceProps {
  readonly open: boolean
  readonly onClose: () => void
  readonly onFind: (query: string, opts: FindOptions) => void
  readonly onFindNext: () => void
  readonly onFindPrev: () => void
  readonly onReplace: () => void
  readonly onReplaceAll: () => void
  readonly matchCount: number
  readonly currentMatchIndex: number | null
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export function FindReplace({
  open,
  onClose,
  onFind,
  onFindNext,
  onFindPrev,
  onReplace,
  onReplaceAll,
  matchCount,
  currentMatchIndex,
}: FindReplaceProps): React.JSX.Element {
  const [findQuery, setFindQuery] = useState('')
  const [, setReplaceQuery] = useState('')
  const [caseSensitive, setCaseSensitive] = useState(false)
  const [wholeWord, setWholeWord] = useState(false)
  const [useRegex, setUseRegex] = useState(false)

  const findInputRef = useRef<HTMLInputElement>(null)

  // Notify parent when query or options change
  const emitFind = useCallback(
    (query: string, cs: boolean, ww: boolean, rx: boolean) => {
      onFind(query, { caseSensitive: cs, wholeWord: ww, useRegex: rx })
    },
    [onFind],
  )

  const handleFindChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const q = e.target.value
      setFindQuery(q)
      emitFind(q, caseSensitive, wholeWord, useRegex)
    },
    [caseSensitive, emitFind, useRegex, wholeWord],
  )

  const handleReplaceChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    setReplaceQuery(e.target.value)
  }, [])

  const handleCaseSensitiveChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const v = e.target.checked
      setCaseSensitive(v)
      emitFind(findQuery, v, wholeWord, useRegex)
    },
    [emitFind, findQuery, useRegex, wholeWord],
  )

  const handleWholeWordChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const v = e.target.checked
      setWholeWord(v)
      emitFind(findQuery, caseSensitive, v, useRegex)
    },
    [caseSensitive, emitFind, findQuery, useRegex],
  )

  const handleUseRegexChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const v = e.target.checked
      setUseRegex(v)
      emitFind(findQuery, caseSensitive, wholeWord, v)
    },
    [caseSensitive, emitFind, findQuery, wholeWord],
  )

  const handleFindKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLInputElement>) => {
      if (e.key === 'Enter') {
        e.preventDefault()
        onFindNext()
      } else if (e.key === 'Escape') {
        e.preventDefault()
        onClose()
      }
    },
    [onClose, onFindNext],
  )

  const handleClose = useCallback(() => {
    onClose()
  }, [onClose])

  const handleFindNext = useCallback(() => {
    onFindNext()
  }, [onFindNext])

  const handleFindPrev = useCallback(() => {
    onFindPrev()
  }, [onFindPrev])

  const handleReplace = useCallback(() => {
    onReplace()
  }, [onReplace])

  const handleReplaceAll = useCallback(() => {
    onReplaceAll()
  }, [onReplaceAll])

  // Focus find input when dialog opens
  useEffect(() => {
    if (open && findInputRef.current) {
      findInputRef.current.focus()
    }
  }, [open])

  const statusText =
    matchCount === 0
      ? 'No matches'
      : currentMatchIndex !== null
        ? `${currentMatchIndex + 1} of ${matchCount}`
        : `${matchCount} match${matchCount === 1 ? '' : 'es'}`

  return (
    <div className={`docx-find${open ? '' : ' docx-find--closed'}`} role="dialog" aria-label="Find and Replace">
      {/* Row 1: Find input + options */}
      <div className="docx-find__row">
        <Search size={14} aria-hidden="true" />
        <input
          ref={findInputRef}
          className="docx-find__input"
          type="text"
          placeholder="Find…"
          value={findQuery}
          onChange={handleFindChange}
          onKeyDown={handleFindKeyDown}
          aria-label="Find"
        />
        <button
          className="docx-find__button docx-find__button--close"
          type="button"
          onClick={handleClose}
          aria-label="Close"
        >
          <X size={13} />
        </button>
      </div>

      {/* Row 2: Replace input */}
      <div className="docx-find__row">
        <Replace size={14} aria-hidden="true" />
        <input
          className="docx-find__input"
          type="text"
          placeholder="Replace…"
          onChange={handleReplaceChange}
          aria-label="Replace"
        />
      </div>

      {/* Row 3: Checkboxes */}
      <div className="docx-find__row">
        <label className="docx-find__checkbox">
          <input
            type="checkbox"
            checked={caseSensitive}
            onChange={handleCaseSensitiveChange}
            aria-label="Case sensitive"
          />
          Case sensitive
        </label>
        <label className="docx-find__checkbox">
          <input
            type="checkbox"
            checked={wholeWord}
            onChange={handleWholeWordChange}
            aria-label="Whole word"
          />
          Whole word
        </label>
        <label className="docx-find__checkbox">
          <input
            type="checkbox"
            checked={useRegex}
            onChange={handleUseRegexChange}
            aria-label="Regex"
          />
          Regex
        </label>
      </div>

      {/* Row 4: Buttons + status */}
      <div className="docx-find__buttons">
        <button
          className="docx-find__button"
          type="button"
          onClick={handleFindPrev}
          aria-label="Find previous"
        >
          <ChevronUp size={13} />
          Prev
        </button>
        <button
          className="docx-find__button"
          type="button"
          onClick={handleFindNext}
          aria-label="Find next"
        >
          <ChevronDown size={13} />
          Next
        </button>
        <button
          className="docx-find__button"
          type="button"
          onClick={handleReplace}
          aria-label="Replace"
        >
          Replace
        </button>
        <button
          className="docx-find__button"
          type="button"
          onClick={handleReplaceAll}
          aria-label="Replace all"
        >
          Replace All
        </button>
        <span className="docx-find__status" aria-live="polite">
          {statusText}
        </span>
      </div>
    </div>
  )
}
