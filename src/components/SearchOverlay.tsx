import { useRef, useEffect } from 'react';
import { Search, ChevronUp, ChevronDown, X } from 'lucide-react';
import { useFocusTrap } from '../hooks/useFocusTrap';
import { useTranslate } from '../i18n';

interface SearchOverlayProps {
  isOpen: boolean;
  query: string;
  matchCount: number;
  currentMatch: number;
  onQueryChange: (q: string) => void;
  onNext: () => void;
  onPrev: () => void;
  onClose: () => void;
}

export function SearchOverlay({
  isOpen,
  query,
  matchCount,
  currentMatch,
  onQueryChange,
  onNext,
  onPrev,
  onClose,
}: SearchOverlayProps) {
  const t = useTranslate();
  const inputRef = useRef<HTMLInputElement>(null);
  const barRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (isOpen) {
      // Small delay to let the animation start
      const timer = setTimeout(() => inputRef.current?.focus(), 50);
      return () => clearTimeout(timer);
    }
  }, [isOpen]);

  // A11Y-2 — Tab was never contained inside the overlay, so it leaked focus
  // straight into the document behind it. The input above already handles
  // its own (animation-delayed) autofocus, so this only adds Tab
  // containment and restores focus to whatever triggered the search on close.
  useFocusTrap(barRef, isOpen, { focusOnOpen: false });

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      if (e.shiftKey) {
        onPrev();
      } else {
        onNext();
      }
    }
  };

  if (!isOpen) return null;

  return (
    <div className="search-overlay">
      <div className="search-overlay__bar" ref={barRef}>
        <Search size={15} className="search-overlay__icon" />
        <input
          ref={inputRef}
          type="text"
          className="search-overlay__input"
          placeholder={t('search.placeholder')}
          value={query}
          onChange={(e) => onQueryChange(e.target.value)}
          onKeyDown={handleKeyDown}
        />
        {query && (
          // UX-15 — announced to screen readers as the match count changes
          // (was silent: sighted-only feedback).
          <span className="search-overlay__count" role="status" aria-live="polite">
            {matchCount > 0 ? t('search.resultCount', { current: currentMatch + 1, total: matchCount }) : t('search.noResults')}
          </span>
        )}
        <div className="search-overlay__nav">
          <button
            className="search-overlay__btn"
            onClick={onPrev}
            disabled={matchCount === 0}
            title={t('search.previousTitle')}
            aria-label={t('search.previousAria')}
          >
            <ChevronUp size={16} />
          </button>
          <button
            className="search-overlay__btn"
            onClick={onNext}
            disabled={matchCount === 0}
            title={t('search.nextTitle')}
            aria-label={t('search.nextAria')}
          >
            <ChevronDown size={16} />
          </button>
        </div>
        <button className="search-overlay__btn search-overlay__close" onClick={onClose} title={t('search.closeTitle')} aria-label={t('search.closeAria')}>
          <X size={16} />
        </button>
      </div>
    </div>
  );
}
