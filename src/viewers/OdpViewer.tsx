import { memo, useCallback, useEffect, useMemo, useState } from 'react'

import type { ViewerProps } from '../formats/types'
import { SlideDeck } from './shared/SlideDeck'
import type { SlideData } from './shared/SlideDeck.types'
import { useSetNavItems, useSetViewerStats } from './shared/useViewerContext'
import { useSlideKeyboardNav } from './slides/shared/useSlideKeyboardNav'
import type { CancelSignal, ZipArchive } from './slides/shared/xmlUtils'
import { parseOdpSlides } from './slides/odp/parser'

function OdpViewerBase({ file }: ViewerProps) {
  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()
  const [slides, setSlides] = useState<ReadonlyArray<SlideData>>([])
  const [activeIndex, setActiveIndex] = useState(0)
  const [error, setError] = useState<string | null>(null)
  const [isLoading, setIsLoading] = useState(false)

  const handleSelectSlide = useCallback((index: number) => {
    setActiveIndex(currentIndex => (currentIndex === index ? currentIndex : index))
  }, [])

  // S14 — hidden slides are excluded from the default view.
  const visibleSlides = useMemo(() => slides.filter(slide => !slide.hidden), [slides])

  // S17 — nav labels prefer each slide's own title placeholder over "Slide N".
  const navItems = useMemo(
    () =>
      visibleSlides.map((slide, index) => ({
        id: slide.id,
        label: slide.title || `Slide ${index + 1}`,
        onSelect: () => {
          handleSelectSlide(index)
        },
      })),
    [handleSelectSlide, visibleSlides],
  )

  useEffect(() => {
    setNavItems(navItems)
  }, [navItems, setNavItems])

  useEffect(() => {
    if (visibleSlides.length === 0) {
      setStats(null)
      return
    }

    setStats({ kind: 'slides', slide: activeIndex + 1, slideCount: visibleSlides.length })
  }, [activeIndex, setStats, visibleSlides.length])

  useEffect(() => {
    if (visibleSlides.length === 0) {
      return
    }

    setActiveIndex(currentIndex => Math.min(currentIndex, visibleSlides.length - 1))
  }, [visibleSlides.length])

  // S15 — PageUp/PageDown plus Arrow/Space (next/previous) and Home/End (first/last).
  useSlideKeyboardNav(visibleSlides.length, setActiveIndex)

  useEffect(() => {
    setNavItems([])
    setStats(null)
    setSlides([])
    setActiveIndex(0)
    setError(null)

    if (file.kind !== 'binary') {
      setIsLoading(false)
      return
    }

    const signal: CancelSignal = { cancelled: false }

    setIsLoading(true)

    void (async () => {
      try {
        const { default: JSZip } = await import('jszip')

        if (signal.cancelled) {
          return
        }

        const zip = await JSZip.loadAsync(file.content)

        if (signal.cancelled) {
          return
        }

        const nextSlides = await parseOdpSlides(zip as ZipArchive, signal)

        if (signal.cancelled) {
          return
        }

        setSlides(nextSlides)
        setActiveIndex(0)
        setError(nextSlides.length > 0 ? null : 'No slides found in ODP.')
        setIsLoading(false)
      } catch (err: unknown) {
        if (signal.cancelled) {
          return
        }

        setError(err instanceof Error ? err.message : String(err))
        setIsLoading(false)
      }
    })()

    return () => {
      signal.cancelled = true
    }
  }, [file, setNavItems, setStats])

  if (file.kind !== 'binary') {
    return (
      <div className="odp-viewer odp-viewer--error">
        OdpViewer received a text file; expected binary.
      </div>
    )
  }

  if (error !== null) {
    return (
      <div className="odp-viewer odp-viewer--error">
        Failed to render ODP: {error}
      </div>
    )
  }

  if (isLoading) {
    return <div className="odp-viewer">Loading ODP slides…</div>
  }

  return (
    <div className="odp-viewer" style={{ width: '100%', height: '100%' }}>
      <SlideDeck
        slides={visibleSlides}
        activeIndex={activeIndex}
        onSelect={handleSelectSlide}
      />
    </div>
  )
}

export const OdpViewer = memo(OdpViewerBase)
