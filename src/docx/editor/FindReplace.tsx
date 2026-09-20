import { useCallback, useEffect, useRef, useState } from 'react'

import { ChevronDown, ChevronUp, Replace, Search, X } from 'lucide-react'

import type { FindOptions } from './Find'
import { useTranslate } from '../../i18n'
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
  const t = useTranslate()
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
      ? t('docx.findReplace.noMatches')
      : currentMatchIndex !== null
        ? t('search.resultCount', { current: currentMatchIndex + 1, total: matchCount })
        : t('docx.findReplace.matchCount', { count: matchCount })

  return (
    <div className={`docx-find${open ? '' : ' docx-find--closed'}`} role="dialog" aria-label={t('docx.findReplace.dialogAria')}>
      {/* Row 1: Find input + options */}
      <div className="docx-find__row">
        <Search size={14} aria-hidden="true" />
        <input
          ref={findInputRef}
          className="docx-find__input"
          type="text"
          placeholder={t('docx.findReplace.findPlaceholder')}
          value={findQuery}
          onChange={handleFindChange}
          onKeyDown={handleFindKeyDown}
          aria-label={t('docx.findReplace.findAria')}
        />
        <button
          className="docx-find__button docx-find__button--close"
          type="button"
          onClick={handleClose}
          aria-label={t('docx.findReplace.close')}
        >
          <X size={13} />
        </button>
      </div>

      {/* Row 2: Replace input */}
      <div className="docx-find__row">
        <Replace size={14} aria-hidden="true" />
        <input
          // i18n — `docx-find__input--replace` (not the translated
          // `aria-label` below) is what `DocxViewer.tsx`'s `getReplaceValue`
          // queries for, so switching the UI language never breaks that
          // lookup.
          className="docx-find__input docx-find__input--replace"
          type="text"
          placeholder={t('docx.findReplace.replacePlaceholder')}
          onChange={handleReplaceChange}
          aria-label={t('docx.findReplace.replaceFieldAria')}
        />
      </div>

      {/* Row 3: Checkboxes */}
      <div className="docx-find__row">
        <label className="docx-find__checkbox">
          <input
            type="checkbox"
            checked={caseSensitive}
            onChange={handleCaseSensitiveChange}
            aria-label={t('docx.findReplace.caseSensitive')}
          />
          {t('docx.findReplace.caseSensitive')}
        </label>
        <label className="docx-find__checkbox">
          <input
            type="checkbox"
            checked={wholeWord}
            onChange={handleWholeWordChange}
            aria-label={t('docx.findReplace.wholeWord')}
          />
          {t('docx.findReplace.wholeWord')}
        </label>
        <label className="docx-find__checkbox">
          <input
            type="checkbox"
            checked={useRegex}
            onChange={handleUseRegexChange}
            aria-label={t('docx.findReplace.regex')}
          />
          {t('docx.findReplace.regex')}
        </label>
      </div>

      {/* Row 4: Buttons + status */}
      <div className="docx-find__buttons">
        <button
          className="docx-find__button"
          type="button"
          onClick={handleFindPrev}
          aria-label={t('docx.findReplace.findPreviousAria')}
        >
          <ChevronUp size={13} />
          {t('docx.findReplace.prev')}
        </button>
        <button
          className="docx-find__button"
          type="button"
          onClick={handleFindNext}
          aria-label={t('docx.findReplace.findNextAria')}
        >
          <ChevronDown size={13} />
          {t('docx.findReplace.next')}
        </button>
        <button
          className="docx-find__button"
          type="button"
          onClick={handleReplace}
          aria-label={t('docx.findReplace.replace')}
        >
          {t('docx.findReplace.replace')}
        </button>
        <button
          className="docx-find__button"
          type="button"
          onClick={handleReplaceAll}
          aria-label={t('docx.findReplace.replaceAllAria')}
        >
          {t('docx.findReplace.replaceAllLabel')}
        </button>
        <span className="docx-find__status" aria-live="polite">
          {statusText}
        </span>
      </div>
    </div>
  )
}
