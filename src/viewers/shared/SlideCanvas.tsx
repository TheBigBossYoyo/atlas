/**
 * Renders one `SlideData`'s shapes at a given CSS scale. Shared by the main
 * viewport and the thumbnail rail so both stay pixel-identical to the parsed
 * model — the whole point of Wave 3-S's format-agnostic `SlideShape[]`.
 */

import { memo, useEffect, useState, type CSSProperties } from 'react'
import type { SlideData, SlideImage, SlideShape } from './SlideDeck.types'
import { SlideTableGrid, TextParagraphs } from './SlideShapeContent'
import { fillToCss, geometryToCss, transformToCss } from './slideStyleHelpers'
import { getDownscaledImage } from './downscaleImage'

/** Device pixel ratio at module load — thumbnails stay reasonably crisp on
 * HiDPI screens without needing to react to a live DPR change (a small
 * fixed size, unlike PDF's page canvases, so this isn't worth re-deriving
 * per render). */
const THUMBNAIL_DPR = typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1

function boxStyle(shape: SlideShape): CSSProperties {
  const { transform } = shape
  return {
    position: 'absolute',
    left: transform.x,
    top: transform.y,
    width: transform.w,
    height: transform.h,
    transform: transformToCss(transform),
  }
}

/**
 * S10 — reproduces `a:srcRect` cropping without knowing the image's natural
 * pixel size.
 *
 * S18/SLD-19 — `downscaleToPx` (non-null only in the thumbnail rail; see
 * `ShapeRenderer` below) caps how many pixels of the source image this ever
 * decodes. Renders nothing (an empty placeholder box, not the full-res
 * image) while the downscale is pending, rather than briefly painting the
 * full-resolution `src` and then swapping it out — that would defeat the
 * whole point by still forcing the expensive decode.
 */
function CroppedImage({
  image,
  downscaleToPx,
}: {
  readonly image: SlideImage
  readonly downscaleToPx: number | null
}) {
  // Only the thumbnail path (downscaleToPx !== null) needs state at all —
  // the interactive main view renders `image.src` directly below with no
  // asynchrony involved.
  const [downscaledSrc, setDownscaledSrc] = useState<string | null>(null)
  const [trackedKey, setTrackedKey] = useState<string | null>(null)
  const key = downscaleToPx === null ? null : `${downscaleToPx}:${image.src}`

  // Render-time state adjustment (React's documented "you might not need an
  // effect" pattern, already used elsewhere in this codebase — e.g.
  // SpreadsheetViewer's active-sheet reset) instead of resetting inside the
  // effect below: clears the previous (now stale) downscaled bitmap in the
  // SAME render the target (src, size) pair changes, rather than briefly
  // re-committing with last render's image before the effect below fires.
  if (key !== trackedKey) {
    setTrackedKey(key)
    setDownscaledSrc(null)
  }

  useEffect(() => {
    if (downscaleToPx === null) {
      return undefined
    }

    let cancelled = false
    void getDownscaledImage(image.src, downscaleToPx).then((src) => {
      if (!cancelled) setDownscaledSrc(src)
    })
    return () => {
      cancelled = true
    }
  }, [image.src, downscaleToPx])

  const resolvedSrc = downscaleToPx === null ? image.src : downscaledSrc

  if (resolvedSrc === null) {
    return <div className="slide-shape__image slide-shape__image--pending" aria-hidden="true" />
  }

  if (!image.crop) {
    return (
      <img
        className="slide-shape__image slide-shape__image--contain"
        src={resolvedSrc}
        alt={image.alt}
        draggable={false}
      />
    )
  }

  const { top, right, bottom, left } = image.crop
  const scaleX = 1 / Math.max(1 - left - right, 0.01)
  const scaleY = 1 / Math.max(1 - top - bottom, 0.01)

  return (
    <div className="slide-shape__image-crop">
      <img
        className="slide-shape__image"
        src={resolvedSrc}
        alt={image.alt}
        draggable={false}
        style={{
          position: 'absolute',
          width: `${scaleX * 100}%`,
          height: `${scaleY * 100}%`,
          left: `${-left * scaleX * 100}%`,
          top: `${-top * scaleY * 100}%`,
        }}
      />
    </div>
  )
}

function ShapeRenderer({
  shape,
  interactive,
  scale,
}: {
  readonly shape: SlideShape
  readonly interactive: boolean
  readonly scale: number
}) {
  const style = boxStyle(shape)
  const pointerEvents: CSSProperties['pointerEvents'] = interactive ? 'auto' : 'none'

  switch (shape.kind) {
    case 'text': {
      const fill = fillToCss(shape.fill)
      return (
        <div
          className="slide-shape slide-shape--text"
          style={{
            ...style,
            ...geometryToCss(shape.geometry),
            background: fill,
            border: shape.border ? `${shape.border.widthPx}px solid ${shape.border.color}` : undefined,
            userSelect: interactive ? 'text' : 'none',
            pointerEvents,
          }}
        >
          <TextParagraphs paragraphs={shape.paragraphs} fontScale={shape.fontScale} />
        </div>
      )
    }

    case 'shape':
      return (
        <div
          className="slide-shape slide-shape--plain"
          style={{
            ...style,
            ...geometryToCss(shape.geometry),
            background: fillToCss(shape.fill),
            border: shape.border ? `${shape.border.widthPx}px solid ${shape.border.color}` : undefined,
            pointerEvents,
          }}
        />
      )

    case 'image': {
      // S18/SLD-19 — the interactive main view needs full fidelity
      // (downscaleToPx: null); the (non-interactive) thumbnail rail never
      // displays more than roughly its own on-screen size, so cap the
      // decoded bitmap to that instead of the source image's full
      // resolution. `shape.transform` is in the slide's native coordinate
      // space, and `scale` is the same CSS scale `SlideCanvasBase` applies
      // to the whole slide, so `transform.{w,h} * scale` is this shape's
      // actual on-screen size.
      const downscaleToPx = interactive
        ? null
        : Math.max(
            16,
            Math.ceil(Math.max(shape.transform.w, shape.transform.h) * scale * THUMBNAIL_DPR),
          )
      return (
        <div className="slide-shape slide-shape--image" style={{ ...style, pointerEvents, overflow: 'hidden' }}>
          <CroppedImage image={shape} downscaleToPx={downscaleToPx} />
        </div>
      )
    }

    case 'table':
      return (
        <div className="slide-shape slide-shape--table" style={{ ...style, pointerEvents }}>
          <SlideTableGrid rows={shape.rows} />
        </div>
      )

    case 'unsupported':
      return (
        <div className="slide-shape slide-shape--unsupported" style={{ ...style, pointerEvents }}>
          {shape.label}
        </div>
      )

    default:
      return null
  }
}

type SlideCanvasProps = {
  readonly slide: SlideData
  readonly scale: number
  readonly interactive: boolean
}

function SlideCanvasBase({ slide, scale, interactive }: SlideCanvasProps) {
  const frameStyle: CSSProperties = {
    width: slide.width * scale,
    height: slide.height * scale,
  }

  const slideStyle: CSSProperties = {
    width: slide.width,
    height: slide.height,
    transform: `scale(${scale})`,
    background: fillToCss(slide.background) ?? 'var(--bg-primary)',
  }

  return (
    <div className="slide-deck__frame" style={frameStyle}>
      <div className="slide-deck__slide" style={slideStyle}>
        {slide.error
          ? <div className="slide-deck__slide-error">{slide.error}</div>
          : slide.shapes.map(shape => (
              <ShapeRenderer key={shape.id} shape={shape} interactive={interactive} scale={scale} />
            ))}
      </div>
    </div>
  )
}

export const SlideCanvas = memo(SlideCanvasBase)
