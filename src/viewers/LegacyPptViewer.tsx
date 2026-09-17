import { memo, useCallback, useEffect, useMemo, useState } from 'react'

import type { ViewerProps } from '../formats/types'
import { SlideDeck } from './shared/SlideDeck'
import type { SlideData } from './shared/SlideDeck.types'
import { LegacyFormatBanner } from './shared/LegacyFormatBanner'
import { useSetNavItems, useSetViewerStats } from './shared/useViewerContext'
import { useSlideKeyboardNav } from './slides/shared/useSlideKeyboardNav'
import './__styles__/viewer-legacy-ppt.css'

/**
 * Read-only viewer for legacy PowerPoint 97-2003 (.ppt) presentations
 * (wave-4 legacy-office). Renders the stacked-text `SlideData[]`
 * `legacy/ppt` extracts through the same shared `SlideDeck` PPTX/ODP use —
 * no positions, images, or per-run formatting survive the extraction (see
 * that module's header for exactly what is and isn't covered).
 */
function LegacyPptViewerBase({ file }: ViewerProps) {
  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()
  const [slides, setSlides] = useState<ReadonlyArray<SlideData>>([])
  const [activeIndex, setActiveIndex] = useState(0)
  const [error, setError] = useState<string | null>(null)
  const [isLoading, setIsLoading] = useState(false)

  const handleSelectSlide = useCallback((index: number) => {
    setActiveIndex((currentIndex) => (currentIndex === index ? currentIndex : index))
  }, [])

  const navItems = useMemo(
    () =>
      slides.map((slide, index) => ({
        id: slide.id,
        label: slide.title || `Slide ${index + 1}`,
        onSelect: () => {
          handleSelectSlide(index)
        },
      })),
    [handleSelectSlide, slides],
  )

  useEffect(() => {
    setNavItems(navItems)
  }, [navItems, setNavItems])

  useEffect(() => {
    if (slides.length === 0) {
      setStats(null)
      return
    }
    setStats({ kind: 'slides', slide: activeIndex + 1, slideCount: slides.length })
  }, [activeIndex, setStats, slides.length])

  // Clamped during render (mirrors PptxViewer's own pattern for the same
  // problem) rather than in a `useEffect` — an effect that reads `slides`
  // and calls `setActiveIndex` forms a two-hop render->effect->render
  // "cascading" update the newer react-hooks lint rule flags; adjusting
  // state directly during render is the pattern React itself recommends for
  // exactly this "clamp state to a range derived from other state" case.
  if (slides.length > 0 && activeIndex > slides.length - 1) {
    setActiveIndex(slides.length - 1)
  }

  useSlideKeyboardNav(slides.length, setActiveIndex)

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

    let cancelled = false
    setIsLoading(true)

    // The actual extraction is synchronous (CPU-bound, no I/O) — dynamically
    // importing the parser (rather than a static top-level import) both
    // code-splits it into its own chunk, only ever fetched once a .ppt is
    // actually opened, and gives this effect a genuine async boundary
    // before it touches state, so "Reading…" actually paints first instead
    // of the parse blocking the same tick that requested it.
    void (async () => {
      try {
        const { extractLegacyPptSlides } = await import('../legacy/ppt')
        if (cancelled) return

        const nextSlides = extractLegacyPptSlides(new Uint8Array(file.content))
        if (cancelled) return

        setSlides(nextSlides)
        setActiveIndex(0)
        setError(nextSlides.length > 0 ? null : 'No slides found in this presentation.')
        setIsLoading(false)
      } catch (err) {
        if (cancelled) return
        setError(err instanceof Error ? err.message : String(err))
        setIsLoading(false)
      }
    })()

    return () => {
      cancelled = true
    }
  }, [file, setNavItems, setStats])

  if (file.kind !== 'binary') {
    return (
      <div className="legacy-ppt-viewer legacy-ppt-viewer--error">
        LegacyPptViewer received a text file; expected binary.
      </div>
    )
  }

  if (error !== null) {
    return (
      <div className="legacy-ppt-viewer legacy-ppt-viewer--error">
        Couldn't read this PowerPoint 97-2003 presentation: {error}
      </div>
    )
  }

  if (isLoading) {
    return <div className="legacy-ppt-viewer">Reading legacy PowerPoint presentation…</div>
  }

  return (
    <div className="legacy-ppt-viewer">
      <LegacyFormatBanner formatLabel="PowerPoint 97-2003 presentation (.ppt)" modernExtension=".pptx" />
      <div className="legacy-ppt-viewer__deck">
        <SlideDeck slides={slides} activeIndex={activeIndex} onSelect={handleSelectSlide} />
      </div>
    </div>
  )
}

export const LegacyPptViewer = memo(LegacyPptViewerBase)
