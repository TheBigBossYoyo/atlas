/**
 * Atlas — DOCX SpellCheckMenu (Wave E.3)
 *
 * Floating menu rendered at the click coordinate when Electron's spellcheck
 * detects a misspelled word.  Lists suggestions; clicking one calls
 * `onReplace(word, replacement)` so the host viewer can issue an editor
 * command to swap the misspelled run text in-place.  Also offers
 * "Add to dictionary".
 *
 * Dismissal: Escape, click outside, or window blur.
 */

import { memo, useCallback, useEffect, useRef } from 'react'

import './__styles__/spell-check-menu.css'

export interface SpellCheckMenuProps {
  readonly word: string
  readonly suggestions: ReadonlyArray<string>
  readonly x: number
  readonly y: number
  readonly onReplace: (misspelled: string, replacement: string) => void
  readonly onAddToDictionary: (word: string) => void
  readonly onDismiss: () => void
}

function SpellCheckMenuBase({
  word,
  suggestions,
  x,
  y,
  onReplace,
  onAddToDictionary,
  onDismiss,
}: SpellCheckMenuProps) {
  const menuRef = useRef<HTMLDivElement | null>(null)

  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        onDismiss()
      }
    }

    const handlePointerDown = (event: MouseEvent) => {
      const root = menuRef.current
      if (root !== null && event.target instanceof Node && !root.contains(event.target)) {
        onDismiss()
      }
    }

    const handleBlur = () => onDismiss()

    window.addEventListener('keydown', handleKeyDown)
    window.addEventListener('mousedown', handlePointerDown)
    window.addEventListener('blur', handleBlur)

    return () => {
      window.removeEventListener('keydown', handleKeyDown)
      window.removeEventListener('mousedown', handlePointerDown)
      window.removeEventListener('blur', handleBlur)
    }
  }, [onDismiss])

  const handleSuggestionClick = useCallback(
    (replacement: string) => {
      onReplace(word, replacement)
      onDismiss()
    },
    [onDismiss, onReplace, word],
  )

  const handleAddClick = useCallback(() => {
    onAddToDictionary(word)
    onDismiss()
  }, [onAddToDictionary, onDismiss, word])

  return (
    <div
      ref={menuRef}
      className="spellcheck-menu"
      role="menu"
      aria-label={`Spell check suggestions for "${word}"`}
      style={{ left: x, top: y }}
    >
      <div className="spellcheck-menu__header">
        <span className="spellcheck-menu__item spellcheck-menu__item--word">{word}</span>
      </div>
      {suggestions.length > 0 ? (
        <ul className="spellcheck-menu__list">
          {suggestions.map((suggestion, index) => (
            <li key={`${suggestion}-${index}`} className="spellcheck-menu__entry">
              <button
                type="button"
                className="spellcheck-menu__item"
                role="menuitem"
                onClick={() => handleSuggestionClick(suggestion)}
              >
                {suggestion}
              </button>
            </li>
          ))}
        </ul>
      ) : (
        <div className="spellcheck-menu__item spellcheck-menu__item--empty">No suggestions</div>
      )}
      <div className="spellcheck-menu__footer">
        <button
          type="button"
          className="spellcheck-menu__item"
          role="menuitem"
          onClick={handleAddClick}
        >
          Add to dictionary
        </button>
      </div>
    </div>
  )
}

export const SpellCheckMenu = memo(SpellCheckMenuBase)
