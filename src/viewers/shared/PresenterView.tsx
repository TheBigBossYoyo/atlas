/**
 * USR-16 — presenter view: the current slide, the next slide, speaker notes,
 * an elapsed timer and the slide counter, full screen. Slide navigation keys
 * keep working (the viewer's keyboard nav drives `activeIndex`); Escape or the
 * exit button leaves.
 */
import { memo, useEffect, useRef, useState } from 'react'
import { ChevronLeft, ChevronRight, X } from 'lucide-react'

import { SlideCanvas } from './SlideCanvas'
import type { SlideData } from './SlideDeck.types'
import { useFocusTrap } from '../../hooks/useFocusTrap'
import { useTranslate } from '../../i18n'

type PresenterViewProps = {
  readonly slides: ReadonlyArray<SlideData>
  readonly activeIndex: number
  readonly onSelect: (index: number) => void
  readonly onExit: () => void
}

function formatElapsed(seconds: number): string {
  const h = Math.floor(seconds / 3600)
  const m = Math.floor((seconds % 3600) / 60)
  const s = seconds % 60
  const pad = (n: number): string => String(n).padStart(2, '0')
  return h > 0 ? `${h}:${pad(m)}:${pad(s)}` : `${pad(m)}:${pad(s)}`
}

function useWindowSize(): { readonly width: number; readonly height: number } {
  const [size, setSize] = useState({ width: window.innerWidth, height: window.innerHeight })
  useEffect(() => {
    const update = (): void => setSize({ width: window.innerWidth, height: window.innerHeight })
    window.addEventListener('resize', update)
    return () => window.removeEventListener('resize', update)
  }, [])
  return size
}

function PresenterViewBase({ slides, activeIndex, onSelect, onExit }: PresenterViewProps) {
  const t = useTranslate()
  const containerRef = useRef<HTMLDivElement | null>(null)
  const [elapsed, setElapsed] = useState(0)
  const { width, height } = useWindowSize()

  // A11Y-2 — Tab containment + focus capture/restore via the shared hook
  // (previously PresenterView had neither: it only captured/restored focus
  // on mount/unmount below, with no Tab handling at all, so Tab walked
  // straight out into the document behind it — worse still, that could
  // happen silently since `requestFullscreen` below can be refused).
  // `focusOnOpen: false` because this view focuses its own container (not a
  // descendant control, see the effect below) so its Escape handler — bound
  // to the container itself — keeps receiving keys regardless of whether
  // fullscreen was granted; the hook's own querySelector-based autofocus
  // would otherwise land on the (initially disabled) Previous-slide button.
  // This component only exists in the tree while presenting (SlideDeck.tsx
  // renders it conditionally), so "open" is simply "mounted" — pass a
  // constant rather than a caller-managed boolean.
  useFocusTrap(containerRef, true, { focusOnOpen: false })

  useEffect(() => {
    const container = containerRef.current
    // Focus the dialog itself (not only on fullscreen success — that request
    // can be refused), or its Escape handler never sees a key.
    container?.focus()
    const handleChange = (): void => {
      if (!document.fullscreenElement) onExit()
    }
    document.addEventListener('fullscreenchange', handleChange)
    void container?.requestFullscreen?.().catch(() => undefined)
    return () => {
      document.removeEventListener('fullscreenchange', handleChange)
      if (document.fullscreenElement) void document.exitFullscreen().catch(() => undefined)
    }
  }, [onExit])

  useEffect(() => {
    const timer = window.setInterval(() => setElapsed((value) => value + 1), 1000)
    return () => window.clearInterval(timer)
  }, [])

  const current = slides[activeIndex]
  const next = slides[activeIndex + 1]
  if (!current) return null

  const currentScale = Math.min((width * 0.62) / current.width, (height - 140) / current.height)
  const nextScale = Math.min((width * 0.3) / current.width, (height * 0.3) / current.height)

  return (
    <div
      ref={containerRef}
      className="presenter-view"
      tabIndex={-1}
      role="dialog"
      aria-label={t('slides.presenter.ariaLabel')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') onExit()
      }}
    >
      <div className="presenter-view__current">
        <SlideCanvas slide={current} scale={currentScale} interactive={false} />
      </div>
      <aside className="presenter-view__side">
        <div className="presenter-view__label">{t('slides.presenter.next')}</div>
        {next ? (
          <SlideCanvas slide={next} scale={nextScale} interactive={false} />
        ) : (
          <div className="presenter-view__end">{t('slides.presenter.endOfPresentation')}</div>
        )}
        <div className="presenter-view__label">{t('slides.presenter.notes')}</div>
        <div className="presenter-view__notes">{current.notes || t('slides.presenter.noNotesForSlide')}</div>
      </aside>
      <footer className="presenter-view__footer">
        <span className="presenter-view__timer" aria-label={t('slides.presenter.elapsedTimeAria')}>{formatElapsed(elapsed)}</span>
        <div className="presenter-view__nav">
          <button type="button" aria-label={t('slides.presenter.previousSlideAria')} disabled={activeIndex === 0} onClick={() => onSelect(activeIndex - 1)}>
            <ChevronLeft size={20} />
          </button>
          <span>
            {t('slides.presenter.slideCounter', { current: activeIndex + 1, total: slides.length })}
          </span>
          <button
            type="button"
            aria-label={t('slides.presenter.nextSlideAria')}
            disabled={activeIndex >= slides.length - 1}
            onClick={() => onSelect(activeIndex + 1)}
          >
            <ChevronRight size={20} />
          </button>
        </div>
        <button type="button" className="presenter-view__exit" aria-label={t('slides.presenter.exitPresenterViewAria')} onClick={onExit}>
          <X size={18} />
        </button>
      </footer>
    </div>
  )
}

export const PresenterView = memo(PresenterViewBase)
