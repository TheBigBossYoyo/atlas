/**
 * Renders one `SlideData`'s shapes at a given CSS scale. Shared by the main
 * viewport and the thumbnail rail so both stay pixel-identical to the parsed
 * model — the whole point of Wave 3-S's format-agnostic `SlideShape[]`.
 */

import { memo, type CSSProperties } from 'react'
import type { SlideData, SlideImage, SlideShape } from './SlideDeck.types'
import { SlideTableGrid, TextParagraphs } from './SlideShapeContent'
import { fillToCss, geometryToCss, transformToCss } from './slideStyleHelpers'

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

/** S10 — reproduces `a:srcRect` cropping without knowing the image's natural pixel size. */
function CroppedImage({ image }: { readonly image: SlideImage }) {
  if (!image.crop) {
    return (
      <img
        className="slide-shape__image slide-shape__image--contain"
        src={image.src}
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
        src={image.src}
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

function ShapeRenderer({ shape, interactive }: { readonly shape: SlideShape; readonly interactive: boolean }) {
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

    case 'image':
      return (
        <div className="slide-shape slide-shape--image" style={{ ...style, pointerEvents, overflow: 'hidden' }}>
          <CroppedImage image={shape} />
        </div>
      )

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
          : slide.shapes.map(shape => <ShapeRenderer key={shape.id} shape={shape} interactive={interactive} />)}
      </div>
    </div>
  )
}

export const SlideCanvas = memo(SlideCanvasBase)
