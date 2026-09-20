import { useEffect, useRef } from 'react'
import type { KeyboardEvent } from 'react'
import { ChevronDown, ChevronUp, Search, X } from 'lucide-react'
import { useTranslate } from '../../i18n'

export type PdfFindBarProps = {
  readonly isOpen: boolean
  readonly query: string
  readonly matchCount: number
  readonly currentMatchIndex: number
  readonly isIndexing: boolean
  readonly indexedPageCount: number
  readonly totalPageCount: number
  readonly onQueryChange: (query: string) => void
  readonly onNext: () => void
  readonly onPrev: () => void
  readonly onClose: () => void
}

export function PdfFindBar({
  isOpen,
  query,
  matchCount,
  currentMatchIndex,
  isIndexing,
  indexedPageCount,
  totalPageCount,
  onQueryChange,
  onNext,
  onPrev,
  onClose,
}: PdfFindBarProps) {
  const t = useTranslate()
  const inputRef = useRef<HTMLInputElement | null>(null)

  useEffect(() => {
    if (isOpen) {
      const timer = setTimeout(() => inputRef.current?.focus(), 50)
      return () => clearTimeout(timer)
    }
    return undefined
  }, [isOpen])

  // Escape closes the bar wherever the focus is: the input only takes focus
  // 50ms after the bar opens, and a click on a page moves it away again, so
  // keying Escape off the input alone left the bar stuck open (which then
  // swallowed the next shell shortcut).
  useEffect(() => {
    if (!isOpen) return undefined
    const onDocumentKeyDown = (event: globalThis.KeyboardEvent): void => {
      if (event.key !== 'Escape' || event.defaultPrevented) return
      event.preventDefault()
      onClose()
    }
    document.addEventListener('keydown', onDocumentKeyDown)
    return () => document.removeEventListener('keydown', onDocumentKeyDown)
  }, [isOpen, onClose])

  if (!isOpen) return null

  const handleKeyDown = (event: KeyboardEvent<HTMLInputElement>) => {
    if (event.key === 'Enter') {
      event.preventDefault()
      if (event.shiftKey) {
        onPrev()
      } else {
        onNext()
      }
    } else if (event.key === 'Escape') {
      event.preventDefault()
      onClose()
    }
  }

  const statusText = query.length === 0
    ? null
    : matchCount > 0
      ? t('search.resultCount', { current: currentMatchIndex + 1, total: matchCount })
      : isIndexing
        ? t('pdf.findBar.searching')
        : t('search.noResults')

  return (
    <div className="pdf-viewer__find-bar" role="search">
      <Search size={15} className="pdf-viewer__find-icon" aria-hidden="true" />
      <input
        ref={inputRef}
        type="text"
        className="pdf-viewer__find-input"
        placeholder={t('pdf.findBar.placeholder')}
        value={query}
        onChange={(e) => onQueryChange(e.target.value)}
        onKeyDown={handleKeyDown}
        aria-label={t('pdf.findBar.inputAria')}
      />
      {statusText && (
        // A11Y pass 3 — mirrors SearchOverlay's own UX-15 fix: match-count
        // changes ("3 of 12", "No results") reached sighted users only.
        <span className="pdf-viewer__find-count" role="status" aria-live="polite">
          {statusText}
        </span>
      )}
      {isIndexing && (
        <span className="pdf-viewer__find-progress" aria-live="polite">
          {t('pdf.findBar.indexingProgress', { indexed: indexedPageCount, total: totalPageCount })}
        </span>
      )}
      <div className="pdf-viewer__find-nav">
        <button
          className="pdf-viewer__toolbar-button"
          onClick={onPrev}
          disabled={matchCount === 0}
          title={t('pdf.findBar.previousMatchTitle')}
          aria-label={t('search.previousAria')}
        >
          <ChevronUp size={16} />
        </button>
        <button
          className="pdf-viewer__toolbar-button"
          onClick={onNext}
          disabled={matchCount === 0}
          title={t('pdf.findBar.nextMatchTitle')}
          aria-label={t('search.nextAria')}
        >
          <ChevronDown size={16} />
        </button>
      </div>
      <button
        className="pdf-viewer__toolbar-button"
        onClick={onClose}
        title={t('search.closeTitle')}
        aria-label={t('pdf.findBar.closeAria')}
      >
        <X size={16} />
      </button>
    </div>
  )
}
