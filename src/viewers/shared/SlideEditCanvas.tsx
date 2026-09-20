/**
 * USR-16 — the editable main slide: click to select a shape, drag to move,
 * drag a handle to resize, double-click (or type Enter) to edit text in
 * place, Delete to remove. All geometry is in slide pixels; the parent
 * `SlideCanvas` applies the zoom scale to this whole layer.
 */
import {
  memo,
  useEffect,
  useRef,
  useState,
  type CSSProperties,
  type KeyboardEvent,
  type MouseEvent,
  type PointerEvent,
} from 'react'

import { SlideCanvas } from './SlideCanvas'
import type { SlideData, SlideShape, SlideTextBox, SlideTransform } from './SlideDeck.types'
import { isPointInShape, resizeBox, type DragMode, type Handle } from './slideGeometry'
import { textBodyToCss } from './slideStyleHelpers'
import { useTranslate } from '../../i18n'

export type SlideShapeBox = { readonly x: number; readonly y: number; readonly w: number; readonly h: number }

const HANDLES: ReadonlyArray<Handle> = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']
const HANDLE_SCREEN_PX = 9
const DRAG_THRESHOLD_PX = 2

type Drag = {
  readonly shapeId: string
  readonly sourceId: string
  readonly handle: DragMode
  readonly startX: number
  readonly startY: number
  readonly origin: SlideTransform
  readonly current: SlideTransform
}

export type SlideEditCanvasProps = {
  readonly slide: SlideData
  readonly scale: number
  readonly selectedShapeId: string | null
  readonly editingShapeId: string | null
  readonly onSelectShape: (shapeId: string | null) => void
  readonly onEditShape: (shapeId: string | null) => void
  readonly onCommitBox: (sourceId: string, box: SlideShapeBox) => Promise<void>
  readonly onCommitText: (sourceId: string, text: string) => void
  readonly onDeleteShape: (sourceId: string) => void
}

/** Handle position INSIDE the selection layer, which is already placed and rotated over the shape. */
function handleStyle(handle: Handle, box: SlideTransform, size: number): CSSProperties {
  const left = handle.includes('w') ? 0 : handle.includes('e') ? box.w : box.w / 2
  const top = handle.startsWith('n') ? 0 : handle.startsWith('s') ? box.h : box.h / 2
  return { left: left - size / 2, top: top - size / 2, width: size, height: size }
}

/** Topmost editable shape under a slide-space point (rotation included). */
function hitTest(shapes: ReadonlyArray<SlideShape>, x: number, y: number): SlideShape | null {
  for (let i = shapes.length - 1; i >= 0; i--) {
    const shape = shapes[i]
    if (shape.sourceId && isPointInShape({ x, y }, shape.transform)) return shape
  }
  return null
}

/** Plain-text editor laid over a text shape, styled like its first run so typing looks like the slide. */
function TextShapeEditor({
  shape,
  onDone,
}: {
  readonly shape: SlideTextBox
  readonly onDone: (text: string | null) => void
}) {
  const t = useTranslate()
  const ref = useRef<HTMLDivElement | null>(null)
  const doneRef = useRef(false)
  const firstParagraph = shape.paragraphs[0]
  const firstRun = shape.paragraphs.flatMap((p) => p.runs).find((run) => run.text !== '\n')
  const fontScale = shape.fontScale ?? 1

  useEffect(() => {
    const element = ref.current
    if (!element) return
    element.innerText = shape.text
    element.focus()
    const selection = window.getSelection()
    selection?.selectAllChildren(element)
    if (shape.text !== '') selection?.collapseToEnd()
    // Mount-only: the editor owns its DOM text while open.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const finish = (commit: boolean): void => {
    if (doneRef.current) return
    doneRef.current = true
    const text = (ref.current?.innerText ?? '').replace(/\n$/, '')
    onDone(commit && text !== shape.text ? text : null)
  }

  const { transform } = shape
  const style: CSSProperties = {
    position: 'absolute',
    left: transform.x,
    top: transform.y,
    width: transform.w,
    minHeight: transform.h,
    ...textBodyToCss(shape.body),
    fontSize: (firstRun?.fontSizePx ?? 24) * fontScale,
    fontFamily: firstRun?.fontFamily,
    fontWeight: firstRun?.bold ? 700 : undefined,
    fontStyle: firstRun?.italic ? 'italic' : undefined,
    color: firstRun?.color,
    textAlign: firstParagraph?.align,
  }

  return (
    <div
      ref={ref}
      className="slide-edit__text-editor"
      style={style}
      contentEditable="plaintext-only"
      role="textbox"
      aria-multiline="true"
      aria-label={t('slides.editCanvas.editTextAria')}
      suppressContentEditableWarning
      onBlur={() => finish(true)}
      onPointerDown={(event) => event.stopPropagation()}
      onKeyDown={(event) => {
        event.stopPropagation()
        if (event.key === 'Escape') {
          event.preventDefault()
          // PowerPoint's behavior: Escape leaves text editing and KEEPS what
          // was typed (Ctrl+Z undoes it), rather than discarding it silently.
          finish(true)
        }
      }}
    />
  )
}

function SlideEditCanvasBase({
  slide,
  scale,
  selectedShapeId,
  editingShapeId,
  onSelectShape,
  onEditShape,
  onCommitBox,
  onCommitText,
  onDeleteShape,
}: SlideEditCanvasProps) {
  const t = useTranslate()
  const layerRef = useRef<HTMLDivElement | null>(null)
  const [drag, setDrag] = useState<Drag | null>(null)
  const [pendingBox, setPendingBox] = useState<{ readonly shapeId: string; readonly transform: SlideTransform } | null>(null)

  const selected = slide.shapes.find((shape) => shape.id === selectedShapeId) ?? null
  const editing = slide.shapes.find((shape) => shape.id === editingShapeId && shape.kind === 'text') as
    | SlideTextBox
    | undefined

  const toSlidePoint = (event: PointerEvent): { x: number; y: number } => {
    const rect = layerRef.current!.getBoundingClientRect()
    return { x: (event.clientX - rect.left) / scale, y: (event.clientY - rect.top) / scale }
  }

  const startDrag = (event: PointerEvent<HTMLDivElement>, shape: SlideShape, handle: DragMode): void => {
    if (!shape.sourceId || !shape.movable) return
    const point = toSlidePoint(event)
    layerRef.current?.setPointerCapture(event.pointerId)
    setDrag({
      shapeId: shape.id,
      sourceId: shape.sourceId,
      handle,
      startX: point.x,
      startY: point.y,
      origin: shape.transform,
      current: shape.transform,
    })
  }

  const handlePointerDown = (event: PointerEvent<HTMLDivElement>): void => {
    if (event.button !== 0) return
    layerRef.current?.focus()
    const point = toSlidePoint(event)
    const hit = hitTest(slide.shapes, point.x, point.y)
    onSelectShape(hit?.id ?? null)
    if (hit) startDrag(event, hit, 'move')
  }

  const handlePointerMove = (event: PointerEvent<HTMLDivElement>): void => {
    if (!drag) return
    const point = toSlidePoint(event)
    const delta = { x: point.x - drag.startX, y: point.y - drag.startY }
    setDrag({ ...drag, current: resizeBox(drag.origin, drag.handle, delta) })
  }

  const handlePointerUp = (event: PointerEvent<HTMLDivElement>): void => {
    if (!drag) return
    layerRef.current?.releasePointerCapture(event.pointerId)
    const { current, origin } = drag
    setDrag(null)
    const moved =
      Math.abs(current.x - origin.x) + Math.abs(current.y - origin.y) + Math.abs(current.w - origin.w) + Math.abs(current.h - origin.h) >
      DRAG_THRESHOLD_PX / scale
    if (!moved) return
    // The preview stays until the edit (and its re-parse) has landed.
    setPendingBox({ shapeId: drag.shapeId, transform: current })
    void onCommitBox(drag.sourceId, { x: current.x, y: current.y, w: current.w, h: current.h }).finally(() => setPendingBox(null))
  }

  const handleDoubleClick = (event: MouseEvent<HTMLDivElement>): void => {
    const rect = layerRef.current!.getBoundingClientRect()
    const hit = hitTest(slide.shapes, (event.clientX - rect.left) / scale, (event.clientY - rect.top) / scale)
    if (hit?.kind === 'text') onEditShape(hit.id)
  }

  const handleKeyDown = (event: KeyboardEvent<HTMLDivElement>): void => {
    if (!selected?.sourceId) return
    if (event.key === 'Delete' || event.key === 'Backspace') {
      event.preventDefault()
      event.stopPropagation()
      onSelectShape(null)
      onDeleteShape(selected.sourceId)
    } else if (event.key === 'Enter' && selected.kind === 'text') {
      event.preventDefault()
      onEditShape(selected.id)
    } else if (event.key === 'Escape') {
      onSelectShape(null)
    }
  }

  const override = drag ? { shapeId: drag.shapeId, transform: drag.current } : (pendingBox ?? undefined)
  const selectionBox = selected ? (override?.shapeId === selected.id ? override.transform : selected.transform) : null
  const handleSize = HANDLE_SCREEN_PX / scale
  const lineWidth = 1.5 / scale

  return (
    <SlideCanvas
      slide={slide}
      scale={scale}
      interactive
      hiddenShapeId={editing?.id}
      transformOverride={override}
      showPlaceholderPrompts
    >
      <div
        ref={layerRef}
        className="slide-edit__layer"
        tabIndex={0}
        aria-label={t('slides.editCanvas.surfaceAria')}
        onPointerDown={handlePointerDown}
        onPointerMove={handlePointerMove}
        onPointerUp={handlePointerUp}
        onDoubleClick={handleDoubleClick}
        onKeyDown={handleKeyDown}
      />
      {selectionBox && !editing && (
        <div
          className="slide-edit__selection-layer"
          style={{
            left: selectionBox.x,
            top: selectionBox.y,
            width: selectionBox.w,
            height: selectionBox.h,
            transform: selectionBox.rotationDeg ? `rotate(${selectionBox.rotationDeg}deg)` : undefined,
          }}
        >
          <div className="slide-edit__selection" style={{ outlineWidth: lineWidth }} />
          {selected?.movable &&
            HANDLES.map((handle) => (
              <div
                key={handle}
                className={`slide-edit__handle slide-edit__handle--${handle}`}
                data-handle={handle}
                style={{ ...handleStyle(handle, selectionBox, handleSize), borderWidth: lineWidth }}
                onPointerDown={(event) => {
                  event.stopPropagation()
                  if (selected) startDrag(event, selected, handle)
                }}
              />
            ))}
        </div>
      )}
      {editing?.sourceId && (
        <TextShapeEditor
          key={editing.id}
          shape={editing}
          onDone={(text) => {
            onEditShape(null)
            if (text !== null && editing.sourceId) onCommitText(editing.sourceId, text)
            layerRef.current?.focus()
          }}
        />
      )}
    </SlideCanvas>
  )
}

export const SlideEditCanvas = memo(SlideEditCanvasBase)
