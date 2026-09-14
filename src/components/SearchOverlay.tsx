import { useRef, useEffect } from 'react';
import { Search, ChevronUp, ChevronDown, X } from 'lucide-react';

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
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (isOpen) {
      // Small delay to let the animation start
      const timer = setTimeout(() => inputRef.current?.focus(), 50);
      return () => clearTimeout(timer);
    }
  }, [isOpen]);

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
      <div className="search-overlay__bar">
        <Search size={15} className="search-overlay__icon" />
        <input
          ref={inputRef}
          type="text"
          className="search-overlay__input"
          placeholder="Search in document..."
          value={query}
          onChange={(e) => onQueryChange(e.target.value)}
          onKeyDown={handleKeyDown}
        />
        {query && (
          // UX-15 — announced to screen readers as the match count changes
          // (was silent: sighted-only feedback).
          <span className="search-overlay__count" role="status" aria-live="polite">
            {matchCount > 0 ? `${currentMatch + 1} of ${matchCount}` : 'No results'}
          </span>
        )}
        <div className="search-overlay__nav">
          <button
            className="search-overlay__btn"
            onClick={onPrev}
            disabled={matchCount === 0}
            title="Previous (Shift+Enter)"
            aria-label="Previous match"
          >
            <ChevronUp size={16} />
          </button>
          <button
            className="search-overlay__btn"
            onClick={onNext}
            disabled={matchCount === 0}
            title="Next (Enter)"
            aria-label="Next match"
          >
            <ChevronDown size={16} />
          </button>
        </div>
        <button className="search-overlay__btn search-overlay__close" onClick={onClose} title="Close (Esc)" aria-label="Close search">
          <X size={16} />
        </button>
      </div>
    </div>
  );
}
