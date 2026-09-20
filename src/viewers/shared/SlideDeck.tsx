import {
  memo,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type RefObject,
} from 'react'
import { List, type RowComponentProps } from 'react-window'
import { Maximize2, Minimize2, MonitorPlay, RotateCcw, StickyNote, X, ZoomIn, ZoomOut } from 'lucide-react'

import '../__styles__/viewer-slides.css'
import { PresenterView } from './PresenterView'
import { SlideCanvas } from './SlideCanvas'
import type { SlideData } from './SlideDeck.types'
import { SlideEditCanvas } from './SlideEditCanvas'
import { SlideEditToolbar, type SlideDeckEditor } from './SlideEditToolbar'
import { useTranslate } from '../../i18n'

type SlideDeckProps = {
  readonly slides: ReadonlyArray<SlideData>
  readonly activeIndex: number
  readonly onSelect: (i: number) => void
  /** USR-16 — makes the deck editable (PPTX); omitted for read-only decks. */
  readonly editor?: SlideDeckEditor
}

type ViewportSize = { readonly width: number; readonly height: number }

/** USR-15 — wide enough for slide titles to stay legible in the rail (240px rail minus row chrome). */
const THUMBNAIL_WIDTH = 200
/** Padding + border + row gap around each virtualized thumbnail row. */
const THUMBNAIL_ROW_CHROME = 32
const MIN_ZOOM = 0.25
const MAX_ZOOM = 4
const ZOOM_STEP = 1.25

type ThumbnailRowProps = {
  readonly slides: ReadonlyArray<SlideData>
  readonly activeIndex: number
  readonly onSelect: (i: number) => void
}

function ThumbnailRow({
  ariaAttributes,
  index,
  style,
  slides,
  activeIndex,
  onSelect,
}: RowComponentProps<ThumbnailRowProps>) {
  const t = useTranslate()
  const slide = slides[index]
  if (!slide) {
    return null
  }

  const isActive = index === activeIndex
  const thumbScale = THUMBNAIL_WIDTH / Math.max(slide.width, 1)
  const label = slide.title || t('slides.deck.slideTitle', { n: index + 1 })

  return (
    <div style={style} {...ariaAttributes} className="slide-deck__thumb-row" role="listitem">
      <button
        type="button"
        className={isActive ? 'slide-deck__thumb slide-deck__thumb--active' : 'slide-deck__thumb'}
        aria-current={isActive ? 'true' : undefined}
        // A11Y pass 3 — an explicit aria-label alongside `title` rather than
        // relying on `title` alone as the accessible-name fallback (the
        // button's only other content is `SlideCanvas`'s `<canvas>`, which
        // contributes no accessible text of its own).
        aria-label={label}
        title={label}
        onClick={() => {
          onSelect(index)
        }}
      >
        <SlideCanvas slide={slide} scale={thumbScale} interactive={false} />
      </button>
    </div>
  )
}

function useViewportSize(ref: RefObject<HTMLDivElement | null>): ViewportSize {
  const [size, setSize] = useState<ViewportSize>({ width: 0, height: 0 })

  useEffect(() => {
    const element = ref.current
    if (!element) {
      return
    }

    const update = () => {
      setSize({ width: element.clientWidth, height: element.clientHeight })
    }

    update()

    if (typeof ResizeObserver === 'undefined') {
      window.addEventListener('resize', update)
      return () => {
        window.removeEventListener('resize', update)
      }
    }

    const observer = new ResizeObserver(update)
    observer.observe(element)
    return () => {
      observer.disconnect()
    }
  }, [ref])

  return size
}

function useFullscreen(containerRef: RefObject<HTMLDivElement | null>) {
  const [isFullscreen, setIsFullscreen] = useState(false)

  useEffect(() => {
    const handleChange = () => {
      setIsFullscreen(Boolean(document.fullscreenElement) && document.fullscreenElement === containerRef.current)
    }

    document.addEventListener('fullscreenchange', handleChange)
    return () => {
      document.removeEventListener('fullscreenchange', handleChange)
    }
  }, [containerRef])

  const toggle = useCallback(() => {
    if (document.fullscreenElement) {
      void document.exitFullscreen()
      return
    }

    void containerRef.current?.requestFullscreen()
  }, [containerRef])

  return { isFullscreen, toggle }
}

function SlideDeckBase({ slides, activeIndex, onSelect, editor }: SlideDeckProps) {
  const t = useTranslate()
  const containerRef = useRef<HTMLDivElement | null>(null)
  const mainRef = useRef<HTMLDivElement | null>(null)
  const viewportSize = useViewportSize(mainRef)
  const [zoom, setZoom] = useState(1)
  const [notesOpen, setNotesOpen] = useState(false)
  const [presenterOpen, setPresenterOpen] = useState(false)
  const [selectedShapeId, setSelectedShapeId] = useState<string | null>(null)
  const [editingShapeId, setEditingShapeId] = useState<string | null>(null)
  const exitPresenter = useCallback(() => setPresenterOpen(false), [])
  const { isFullscreen, toggle: toggleFullscreen } = useFullscreen(containerRef)

  const activeSlide = useMemo(() => slides[activeIndex] ?? null, [activeIndex, slides])

  // Selection belongs to one slide: clear it when another slide becomes active.
  const [selectionSlideId, setSelectionSlideId] = useState(activeSlide?.id)
  if (selectionSlideId !== activeSlide?.id) {
    setSelectionSlideId(activeSlide?.id)
    setSelectedShapeId(null)
    setEditingShapeId(null)
  }

  const fitScale = useMemo(() => {
    if (!activeSlide) {
      return 1
    }

    const availableWidth = Math.max(viewportSize.width - 48, 1)
    const availableHeight = Math.max(viewportSize.height - 48, 1)

    return Math.max(
      Math.min(availableWidth / activeSlide.width, availableHeight / activeSlide.height),
      0.05,
    )
  }, [activeSlide, viewportSize.height, viewportSize.width])

  const mainScale = fitScale * zoom

  const zoomIn = useCallback(() => {
    setZoom(current => Math.min(current * ZOOM_STEP, MAX_ZOOM))
  }, [])

  const zoomOut = useCallback(() => {
    setZoom(current => Math.max(current / ZOOM_STEP, MIN_ZOOM))
  }, [])

  const zoomReset = useCallback(() => {
    setZoom(1)
  }, [])

  const rowHeight = useCallback(
    (index: number, rowProps: ThumbnailRowProps) => {
      const slide = rowProps.slides[index]
      if (!slide) {
        return THUMBNAIL_ROW_CHROME
      }

      return (THUMBNAIL_WIDTH / Math.max(slide.width, 1)) * slide.height + THUMBNAIL_ROW_CHROME
    },
    [],
  )

  const railRowProps = useMemo<ThumbnailRowProps>(() => ({ slides, activeIndex, onSelect }), [activeIndex, onSelect, slides])

  return (
    <div ref={containerRef} className={isFullscreen ? 'slide-deck slide-deck--fullscreen' : 'slide-deck'}>
      {!isFullscreen && (
        <div className="slide-deck__rail" role="list" aria-label={t('slides.deck.thumbnailRailAria')}>
          {slides.length > 0 && (
            <List
              rowComponent={ThumbnailRow}
              rowCount={slides.length}
              rowHeight={rowHeight}
              rowProps={railRowProps}
              style={{ height: '100%' }}
              overscanCount={4}
            />
          )}
        </div>
      )}

      <div className="slide-deck__main-column">
        {!isFullscreen && (
          <div className="slide-deck__toolbar">
            <div className="slide-deck__toolbar-group">
              <button type="button" className="slide-deck__toolbar-button" onClick={zoomOut} title={t('slides.deck.zoomOut')}>
                <ZoomOut size={16} />
              </button>
              <span className="slide-deck__zoom-value">{Math.round(zoom * 100)}%</span>
              <button type="button" className="slide-deck__toolbar-button" onClick={zoomIn} title={t('slides.deck.zoomIn')}>
                <ZoomIn size={16} />
              </button>
              <button type="button" className="slide-deck__toolbar-button" onClick={zoomReset} title={t('slides.deck.zoomReset')}>
                <RotateCcw size={16} />
              </button>
            </div>
            {editor && activeSlide && (
              <SlideEditToolbar
                editor={editor}
                slideIndex={activeIndex}
                slideCount={slides.length}
                onTextBoxInserted={(sourceId) => {
                  const shape = activeSlide.shapes.find((candidate) => candidate.sourceId === sourceId)
                  setSelectedShapeId(shape?.id ?? null)
                }}
              />
            )}
            <div className="slide-deck__toolbar-group">
              {(activeSlide?.notes || (editor && editor.canEditNotes)) && (
                <button
                  type="button"
                  className="slide-deck__toolbar-button"
                  aria-pressed={notesOpen}
                  title={t('slides.deck.speakerNotes')}
                  onClick={() => {
                    setNotesOpen(open => !open)
                  }}
                >
                  <StickyNote size={16} />
                </button>
              )}
              <button type="button" className="slide-deck__toolbar-button" title={t('slides.deck.presentFullscreen')} onClick={toggleFullscreen}>
                <Maximize2 size={16} />
              </button>
              <button
                type="button"
                className="slide-deck__toolbar-button"
                title={t('slides.deck.presenterView')}
                aria-label={t('slides.deck.presenterView')}
                onClick={() => setPresenterOpen(true)}
              >
                <MonitorPlay size={16} />
              </button>
            </div>
          </div>
        )}

        <div ref={mainRef} className="slide-deck__main">
          {activeSlide && editor && !isFullscreen ? (
            <SlideEditCanvas
              slide={activeSlide}
              scale={mainScale}
              selectedShapeId={selectedShapeId}
              editingShapeId={editingShapeId}
              onSelectShape={setSelectedShapeId}
              onEditShape={setEditingShapeId}
              onCommitBox={editor.onShapeBox}
              onCommitText={editor.onShapeText}
              onDeleteShape={editor.onDeleteShape}
            />
          ) : activeSlide ? (
            <SlideCanvas slide={activeSlide} scale={mainScale} interactive />
          ) : (
            <div className="slide-deck__empty">{t('slides.deck.noSlidesAvailable')}</div>
          )}

          {isFullscreen && (
            <button
              type="button"
              className="slide-deck__exit-fullscreen"
              title={t('slides.deck.exitPresentation')}
              onClick={toggleFullscreen}
            >
              <Minimize2 size={18} />
            </button>
          )}
        </div>

        {editor?.saveError && <div className="slide-deck__save-error" role="alert">{editor.saveError}</div>}

        {notesOpen && activeSlide && (activeSlide.notes || editor) && (
          <div className="slide-deck__notes">
            <div className="slide-deck__notes-header">
              <span>{t('slides.deck.speakerNotes')}</span>
              <button
                type="button"
                className="slide-deck__toolbar-button"
                title={t('slides.deck.closeNotes')}
                onClick={() => {
                  setNotesOpen(false)
                }}
              >
                <X size={14} />
              </button>
            </div>
            {editor ? (
              <textarea
                key={activeSlide.id}
                className="slide-deck__notes-editor"
                aria-label={t('slides.deck.speakerNotes')}
                defaultValue={activeSlide.notes ?? ''}
                disabled={!editor.canEditNotes}
                placeholder={t('slides.deck.notesPlaceholder')}
                onBlur={(event) => {
                  if (event.currentTarget.value !== (activeSlide.notes ?? '')) editor.onNotesChange(event.currentTarget.value)
                }}
              />
            ) : (
              <div className="slide-deck__notes-body">{activeSlide.notes}</div>
            )}
          </div>
        )}
      </div>

      {presenterOpen && (
        <PresenterView slides={slides} activeIndex={activeIndex} onSelect={onSelect} onExit={exitPresenter} />
      )}
    </div>
  )
}

export const SlideDeck = memo(SlideDeckBase)
