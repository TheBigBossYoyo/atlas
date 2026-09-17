import { memo, useCallback, useEffect, useMemo, useState } from 'react'

import type { ViewerProps } from '../formats/types'
import { SlideDeck } from './shared/SlideDeck'
import type { SlideDeckEditor } from './shared/SlideEditToolbar'
import { useSetNavItems, useSetViewerStats } from './shared/useViewerContext'
import { useSlideKeyboardNav } from './slides/shared/useSlideKeyboardNav'
import { useOdpEditor } from './slides/odp/editing/useOdpEditor'

/** USR-16 — share of the slide a new text box takes, centered. */
const NEW_TEXT_BOX_WIDTH = 0.5
const NEW_TEXT_BOX_HEIGHT_PX = 60

function OdpViewerBase({ file }: ViewerProps) {
  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()
  const [activeIndex, setActiveIndex] = useState(0)
  const editor = useOdpEditor(file.kind === 'binary' ? file.content : null, file.path)
  const { slides } = editor

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

  if (visibleSlides.length > 0 && activeIndex > visibleSlides.length - 1) {
    setActiveIndex(visibleSlides.length - 1)
  }

  // S15 — PageUp/PageDown plus Arrow/Space (next/previous) and Home/End (first/last).
  useSlideKeyboardNav(visibleSlides.length, setActiveIndex)

  const active = visibleSlides[activeIndex]
  // Edits address slides by their position in the whole deck (hidden slides included).
  const deckIndex = active?.index ?? 0
  const visiblePositionOf = useCallback(
    (index: number) => Math.max(0, editor.slides.filter((slide, i) => i <= index && !slide.hidden).length - 1),
    [editor.slides],
  )

  const deckEditor = useMemo<SlideDeckEditor | undefined>(() => {
    if (!active) return undefined
    return {
      canUndo: editor.canUndo,
      canRedo: editor.canRedo,
      onUndo: editor.undo,
      onRedo: editor.redo,
      onAddSlide: () => {
        void editor.addSlide(deckIndex).then(index => setActiveIndex(visiblePositionOf(index)))
      },
      onDuplicateSlide: () => {
        void editor.duplicateSlide(deckIndex).then(() => setActiveIndex(current => current + 1))
      },
      onDeleteSlide: () => {
        void editor.deleteSlide(deckIndex)
      },
      onMoveSlide: (delta) => {
        const target = visibleSlides[activeIndex + delta]
        if (!target) return
        void editor.moveSlide(deckIndex, target.index).then(() => setActiveIndex(current => current + delta))
      },
      onInsertTextBox: () =>
        editor.insertTextBox(deckIndex, {
          x: (active.width * (1 - NEW_TEXT_BOX_WIDTH)) / 2,
          y: (active.height - NEW_TEXT_BOX_HEIGHT_PX) / 2,
          w: active.width * NEW_TEXT_BOX_WIDTH,
          h: NEW_TEXT_BOX_HEIGHT_PX,
        }),
      onSave: () => {
        void editor.save()
      },
      onSaveAs: () => {
        void editor.saveAs()
      },
      canEditNotes: editor.canEditNotes(),
      onNotesChange: (text) => {
        void editor.setNotes(deckIndex, text)
      },
      onShapeText: (sourceId, text) => {
        void editor.setShapeText(deckIndex, sourceId, text)
      },
      onShapeBox: (sourceId, box) => editor.setShapeBox(deckIndex, sourceId, box),
      onDeleteShape: (sourceId) => {
        void editor.deleteShape(deckIndex, sourceId)
      },
      saveError: editor.saveError,
    }
  }, [active, activeIndex, deckIndex, editor, visiblePositionOf, visibleSlides])

  if (file.kind !== 'binary') {
    return (
      <div className="odp-viewer odp-viewer--error">
        OdpViewer received a text file; expected binary.
      </div>
    )
  }

  if (editor.status === 'error') {
    return (
      <div className="odp-viewer odp-viewer--error">
        Failed to render this presentation: {editor.error}
      </div>
    )
  }

  if (editor.status === 'loading') {
    return <div className="odp-viewer">Loading slides…</div>
  }

  return (
    <div className="odp-viewer" style={{ width: '100%', height: '100%' }}>
      <SlideDeck slides={visibleSlides} activeIndex={activeIndex} onSelect={handleSelectSlide} editor={deckEditor} />
    </div>
  )
}

export const OdpViewer = memo(OdpViewerBase)
