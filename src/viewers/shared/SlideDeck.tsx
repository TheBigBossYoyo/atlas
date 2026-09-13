import { memo, useCallback, useEffect, useMemo, useRef, useState, type CSSProperties } from 'react'

import type { SlideData } from './SlideDeck.types'

type SlideDeckProps = {
  readonly slides: ReadonlyArray<SlideData>
  readonly activeIndex: number
  readonly onSelect: (i: number) => void
}

type ViewportSize = {
  readonly width: number
  readonly height: number
}

const THUMBNAIL_WIDTH = 160

function SlideDeckBase({ slides, activeIndex, onSelect }: SlideDeckProps) {
  const mainRef = useRef<HTMLDivElement | null>(null)
  const [viewportSize, setViewportSize] = useState<ViewportSize>({ width: 0, height: 0 })

  const activeSlide = useMemo(
    () => slides[activeIndex] ?? null,
    [activeIndex, slides],
  )

  useEffect(() => {
    const element = mainRef.current
    if (!element) {
      return
    }

    const updateViewportSize = () => {
      setViewportSize({ width: element.clientWidth, height: element.clientHeight })
    }

    updateViewportSize()

    if (typeof ResizeObserver === 'undefined') {
      window.addEventListener('resize', updateViewportSize)

      return () => {
        window.removeEventListener('resize', updateViewportSize)
      }
    }

    const observer = new ResizeObserver(() => {
      updateViewportSize()
    })

    observer.observe(element)

    return () => {
      observer.disconnect()
    }
  }, [])

  const mainScale = useMemo(() => {
    if (!activeSlide) {
      return 1
    }

    const availableWidth = Math.max(viewportSize.width - 48, 1)
    const availableHeight = Math.max(viewportSize.height - 48, 1)

    return Math.max(
      Math.min(
        availableWidth / activeSlide.width,
        availableHeight / activeSlide.height,
      ),
      0.1,
    )
  }, [activeSlide, viewportSize.height, viewportSize.width])

  const renderSlide = useCallback(
    (slide: SlideData, scale: number, selectable: boolean) => {
      const frameStyle: CSSProperties = {
        width: slide.width * scale,
        height: slide.height * scale,
        position: 'relative',
        overflow: 'hidden',
        borderRadius: 'var(--radius-md)',
        border: '1px solid var(--border-primary)',
        background: 'var(--bg-elevated)',
        boxShadow: 'var(--shadow-sm)',
      }

      const slideStyle: CSSProperties = {
        width: slide.width,
        height: slide.height,
        position: 'relative',
        transform: `scale(${scale})`,
        transformOrigin: 'top left',
        background: 'var(--bg-primary)',
        color: 'var(--text-primary)',
      }

      return (
        <div style={frameStyle}>
          <div className="slide-deck__slide" style={slideStyle}>
            {slide.images.map((image, index) => (
              <img
                key={`${slide.id}-image-${index}`}
                className="slide-deck__image"
                src={image.src}
                alt=""
                draggable={false}
                style={{
                  position: 'absolute',
                  left: image.x,
                  top: image.y,
                  width: image.w,
                  height: image.h,
                  objectFit: 'contain',
                  userSelect: 'none',
                  pointerEvents: selectable ? 'auto' : 'none',
                }}
              />
            ))}
            {slide.texts.map((textItem, index) => {
              const fallbackLeft = 24
              const fallbackTop = 24 + (index * 28)
              const left = textItem.x ?? fallbackLeft
              const top = textItem.y ?? fallbackTop
              const width = textItem.w ?? Math.max(slide.width - left - 24, 120)
              const minHeight = textItem.h ?? 24

              return (
                <div
                  key={`${slide.id}-text-${index}`}
                  className="slide-deck__text"
                  style={{
                    position: 'absolute',
                    left,
                    top,
                    width,
                    minHeight,
                    padding: 2,
                    whiteSpace: 'pre-wrap',
                    overflowWrap: 'anywhere',
                    lineHeight: 1.35,
                    userSelect: selectable ? 'text' : 'none',
                    pointerEvents: selectable ? 'auto' : 'none',
                  }}
                >
                  {textItem.text}
                </div>
              )
            })}
          </div>
        </div>
      )
    },
    [],
  )

  return (
    <div
      className="slide-deck"
      style={{
        display: 'flex',
        width: '100%',
        height: '100%',
        minWidth: 0,
        minHeight: 0,
        background: 'var(--bg-primary)',
        color: 'var(--text-primary)',
        overflow: 'hidden',
      }}
    >
      <div
        className="slide-deck__rail"
        style={{
          width: 192,
          minWidth: 192,
          padding: 12,
          display: 'flex',
          flexDirection: 'column',
          gap: 12,
          overflowY: 'auto',
          borderRight: '1px solid var(--border-primary)',
          background: 'var(--bg-secondary)',
        }}
      >
        {slides.map((slide, index) => {
          const isActive = index === activeIndex
          const thumbScale = THUMBNAIL_WIDTH / Math.max(slide.width, 1)
          const thumbClassName = isActive
            ? 'slide-deck__thumb slide-deck__thumb--active'
            : 'slide-deck__thumb'

          return (
            <button
              key={slide.id}
              type="button"
              className={thumbClassName}
              aria-current={isActive ? 'true' : undefined}
              onClick={() => {
                onSelect(index)
              }}
              style={{
                width: '100%',
                padding: 8,
                display: 'flex',
                justifyContent: 'center',
                borderRadius: 'var(--radius-md)',
                background: isActive ? 'var(--accent-subtle)' : 'var(--bg-elevated)',
                border: `1px solid ${isActive ? 'var(--accent)' : 'var(--border-primary)'}`,
                boxShadow: 'var(--shadow-sm)',
              }}
            >
              {renderSlide(slide, thumbScale, false)}
            </button>
          )
        })}
      </div>
      <div
        ref={mainRef}
        className="slide-deck__main"
        style={{
          flex: 1,
          minWidth: 0,
          minHeight: 0,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          padding: 24,
          overflow: 'auto',
          background: 'var(--bg-primary)',
        }}
      >
        {activeSlide
          ? renderSlide(activeSlide, mainScale, true)
          : (
              <div style={{ color: 'var(--text-secondary)' }}>No slides available.</div>
            )}
      </div>
    </div>
  )
}

export const SlideDeck = memo(SlideDeckBase)
