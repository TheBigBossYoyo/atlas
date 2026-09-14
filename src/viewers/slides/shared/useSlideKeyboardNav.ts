/**
 * S15 — keyboard navigation shared by `PptxViewer`/`OdpViewer`: PageUp/PageDown
 * (existing), plus Arrow keys, Space (next/previous), and Home/End (first/last).
 * Ignored while focus is inside an editable control.
 */

import { useEffect } from 'react'

type SetActiveIndex = (updater: (currentIndex: number) => number) => void

function isEditableTarget(target: EventTarget | null): boolean {
  return (
    target instanceof HTMLElement
    && (target.isContentEditable
      || target.tagName === 'INPUT'
      || target.tagName === 'TEXTAREA'
      || target.tagName === 'SELECT')
  )
}

export function useSlideKeyboardNav(slideCount: number, setActiveIndex: SetActiveIndex): void {
  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (slideCount === 0 || isEditableTarget(event.target)) {
        return
      }

      switch (event.key) {
        case 'PageDown':
        case 'ArrowRight':
        case 'ArrowDown':
        case ' ':
          event.preventDefault()
          setActiveIndex(currentIndex => Math.min(currentIndex + 1, slideCount - 1))
          break
        case 'PageUp':
        case 'ArrowLeft':
        case 'ArrowUp':
          event.preventDefault()
          setActiveIndex(currentIndex => Math.max(currentIndex - 1, 0))
          break
        case 'Home':
          event.preventDefault()
          setActiveIndex(() => 0)
          break
        case 'End':
          event.preventDefault()
          setActiveIndex(() => slideCount - 1)
          break
        default:
          break
      }
    }

    window.addEventListener('keydown', handleKeyDown)

    return () => {
      window.removeEventListener('keydown', handleKeyDown)
    }
  }, [setActiveIndex, slideCount])
}
