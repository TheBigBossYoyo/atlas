/**
 * One page's editing surface (PDF-EDIT-1/2): draws the page's pending edits
 * over the pdf.js render and turns pointer input into new ones.
 *
 * Annotations are drawn in PDF user space — an SVG group carrying the page
 * viewport's own transform — so a draft sits exactly where it will be written
 * at any zoom or rotation, and the strokes, text sizes and boxes on screen are
 * the same numbers the writer uses (`textLayout.ts`).
 *
 * Highlights live in their own layer under the text layer, blended with
 * `multiply` like the `/Multiply` appearance the writer gives them, so the
 * paper turns yellow while the ink stays dark — and the text above stays
 * selectable.
 *
 * Edits of the page's own text are NOT drawn here once the live preview has
 * caught up: the page itself is then rendered from the edited file (see
 * `useTextPreview.ts`). Only while the preview is being rebuilt does this
 * layer cover the old paragraph and sketch the new text, for a fraction of a
 * second.
 */

import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent, PointerEvent as ReactPointerEvent } from 'react'

import { useTranslate } from '../../../i18n'
import type { PdfPageOverlayInfo } from '../PdfPage'
import type { PageRotation } from '../types'
import { normalizeRotation } from '../rotation'
import { translateAnnotation } from './editState'
import { TEXT_BOX_PADDING, boundingRect, cssFontFamily, layoutTextBox, rotateVector, textBoxCorners } from './textLayout'
import { sampleRunColors } from './textRuns'
import type { TextParagraph } from './content/paragraphs'
import { editParagraph, newTextEdit, spansText, spansUnchanged } from './content/textEdits'
import { ParagraphEditor, type ParagraphEditorValue } from './ParagraphEditor'
import { cssFill } from './paragraphEditorDom'
import type { PdfAnnotationDraft, PdfPoint, PdfRect, RgbColor, TextBoxAnnotation, TextEdit } from './types'
import type { PdfEditor } from './usePdfEditor'

const HELVETICA = { family: 'helvetica', bold: false, italic: false } as const
/** Ink points closer together than this (points) add nothing visible and bloat the file. */
const INK_MIN_STEP = 0.75

let nextDraftId = 0
function newDraftId(prefix: string): string {
  nextDraftId += 1
  return `atlas-${prefix}-${Date.now().toString(36)}-${nextDraftId}`
}

function cssColor(color: RgbColor): string {
  return `rgb(${Math.round(color.r * 255)}, ${Math.round(color.g * 255)}, ${Math.round(color.b * 255)})`
}

type ViewportLike = PdfPageOverlayInfo['viewport']

function toPdfPoint(viewport: ViewportLike, layer: HTMLElement, clientX: number, clientY: number): PdfPoint {
  const box = layer.getBoundingClientRect()
  const [x, y] = viewport.convertToPdfPoint(clientX - box.left, clientY - box.top)
  return { x, y }
}

function toScreenPoint(viewport: ViewportLike, point: PdfPoint): { left: number; top: number } {
  const [left, top] = viewport.convertToViewportPoint(point.x, point.y)
  return { left, top }
}

function screenBox(viewport: ViewportLike, points: ReadonlyArray<PdfPoint>) {
  const screen = points.map((point) => toScreenPoint(viewport, point))
  const left = Math.min(...screen.map((p) => p.left))
  const top = Math.min(...screen.map((p) => p.top))
  return {
    left,
    top,
    width: Math.max(...screen.map((p) => p.left)) - left,
    height: Math.max(...screen.map((p) => p.top)) - top,
  }
}

/** Whether `point` is inside a convex polygon (a paragraph's four corners). */
function insidePolygon(point: PdfPoint, corners: ReadonlyArray<PdfPoint>): boolean {
  let sign = 0
  for (let i = 0; i < corners.length; i += 1) {
    const a = corners[i]
    const b = corners[(i + 1) % corners.length]
    const cross = (b.x - a.x) * (point.y - a.y) - (b.y - a.y) * (point.x - a.x)
    if (Math.abs(cross) < 1e-9) continue
    const current = Math.sign(cross)
    if (sign === 0) sign = current
    else if (current !== sign) return false
  }
  return true
}

/** The paper colour behind an area, read off the rendered canvas (white when it can't be). */
function sampleBackground(info: PdfPageOverlayInfo, corners: ReadonlyArray<PdfPoint>): RgbColor {
  const canvas = info.canvas
  const context = canvas?.getContext('2d')
  if (!canvas || !context) return { r: 1, g: 1, b: 1 }
  const box = screenBox(info.viewport, corners)
  const ratio = canvas.width / Math.max(1, info.viewport.width)
  const x = Math.max(0, Math.floor(box.left * ratio))
  const y = Math.max(0, Math.floor(box.top * ratio))
  const width = Math.min(canvas.width - x, Math.ceil(box.width * ratio))
  const height = Math.min(canvas.height - y, Math.ceil(box.height * ratio))
  if (width <= 0 || height <= 0) return { r: 1, g: 1, b: 1 }
  try {
    return sampleRunColors(context.getImageData(x, y, width, height)).background
  } catch {
    return { r: 1, g: 1, b: 1 }
  }
}

function boundsOf(annotation: PdfAnnotationDraft, measure: (line: string, size: number) => number): PdfRect {
  switch (annotation.kind) {
    case 'highlight':
      return boundingRect(annotation.rects.flatMap((r) => [{ x: r.x, y: r.y }, { x: r.x + r.width, y: r.y + r.height }]))
    case 'ink': {
      const rect = boundingRect(annotation.points)
      const pad = annotation.width / 2 + 1
      return { x: rect.x - pad, y: rect.y - pad, width: rect.width + pad * 2, height: rect.height + pad * 2 }
    }
    case 'text': {
      const layout = layoutTextBox(annotation.text, annotation.fontSize, (line) => measure(line, annotation.fontSize))
      return boundingRect(textBoxCorners(annotation.anchor, annotation.rotation, layout.width, layout.height))
    }
  }
}

/** Degrees the text runs at on screen, clockwise from horizontal. */
function screenAngle(viewport: ViewportLike, origin: PdfPoint, direction: PdfPoint): number {
  const a = toScreenPoint(viewport, origin)
  const b = toScreenPoint(viewport, { x: origin.x + direction.x, y: origin.y + direction.y })
  return (Math.atan2(b.top - a.top, b.left - a.left) * 180) / Math.PI
}

// ---- Editors ------------------------------------------------------------------------

type TextBoxDraftEditor = {
  readonly kind: 'text-box'
  /** The annotation being edited, or null for a new box. */
  readonly existing: TextBoxAnnotation | null
  readonly anchor: PdfPoint
  readonly rotation: PageRotation
  readonly fontSize: number
  readonly color: RgbColor
  readonly text: string
}

type ParagraphDraftEditor = {
  readonly kind: 'paragraph'
  /** The paragraph being edited, or null for new text. */
  readonly paragraph: TextParagraph | null
  readonly existing: TextEdit | null
  readonly initial: ParagraphEditorValue
  /** Where the editor sits: the text's top-left corner and its baseline direction. */
  readonly corner: PdfPoint
  readonly direction: PdfPoint
  /** PDF points: the width to wrap at (null: no wrapping) and the area's height. */
  readonly width: number | null
  readonly height: number
  /** For new text: where its first baseline starts. */
  readonly newOrigin: PdfPoint | null
}

type OpenEditor = TextBoxDraftEditor | ParagraphDraftEditor

export type PdfEditLayerProps = {
  readonly info: PdfPageOverlayInfo
  /** The page's saved rotation on top of its own `/Rotate` — how other readers will show it. */
  readonly documentRotation: PageRotation
  readonly annotations: ReadonlyArray<PdfAnnotationDraft>
  readonly textEdits: ReadonlyArray<TextEdit>
  /** Whether the page is being drawn from a preview that already contains `textEdits`. */
  readonly previewCurrent: boolean
  readonly editor: PdfEditor
}

export function PdfEditLayer({ info, documentRotation, annotations, textEdits, previewCurrent, editor }: PdfEditLayerProps) {
  const t = useTranslate()
  const { viewport, sourcePageNumber } = info
  const layerRef = useRef<HTMLDivElement | null>(null)
  const [openEditor, setOpenEditorState] = useState<OpenEditor | null>(null)
  // Mirrored in a ref and cleared synchronously on commit: a click outside an
  // open editor commits it from `pointerdown`, and the `blur` that follows
  // would otherwise commit the same draft a second time from its stale closure.
  const openEditorRef = useRef<OpenEditor | null>(null)
  const setOpenEditor = useCallback((next: OpenEditor | null) => {
    openEditorRef.current = next
    setOpenEditorState(next)
  }, [])
  useEffect(() => {
    openEditorRef.current = openEditor
  }, [openEditor])
  /** Reads the paragraph editor's current content; registered by ParagraphEditor while it is open. */
  const readParagraphRef = useRef<(() => ParagraphEditorValue) | null>(null)
  const registerRead = useCallback((read: (() => ParagraphEditorValue) | null) => {
    readParagraphRef.current = read
  }, [])

  const [inkPoints, setInkPoints] = useState<ReadonlyArray<PdfPoint> | null>(null)
  const [hoverParagraph, setHoverParagraph] = useState<TextParagraph | null>(null)
  const [paragraphs, setParagraphs] = useState<ReadonlyArray<TextParagraph> | null>(null)
  const [drag, setDrag] = useState<{ readonly id: string; readonly start: PdfPoint; readonly delta: PdfPoint } | null>(
    null,
  )

  const { tool, fontTools, selectedId, setSelectedId, toolColors, fontSize, inkWidth, registerFlush, textAnalysis } = editor
  const isEditing = editor.isEditing && !editor.isLocked
  const measure = useCallback(
    (line: string, size: number) => fontTools?.widthOf(HELVETICA, line, size) ?? line.length * size * 0.5,
    [fontTools],
  )

  // The page's paragraphs, found once the edit-text tool is in use.
  const wantsParagraphs = isEditing && (tool === 'edit-text' || tool === 'select') && textAnalysis !== null
  useEffect(() => {
    if (!wantsParagraphs || !textAnalysis) return undefined
    let cancelled = false
    textAnalysis
      .paragraphs(sourcePageNumber)
      .then((found) => {
        if (!cancelled) setParagraphs(found)
      })
      .catch((error: unknown) => {
        console.error('Could not read the page text for editing', error)
        if (!cancelled) setParagraphs([])
      })
    return () => {
      cancelled = true
    }
  }, [sourcePageNumber, textAnalysis, wantsParagraphs])

  // A tool switch or leaving edit mode closes whatever was open on this page.
  const [editorTool, setEditorTool] = useState(tool)
  if (editorTool !== tool) {
    setEditorTool(tool)
    setOpenEditorState(null)
    setHoverParagraph(null)
  }

  const pointFromEvent = useCallback(
    (event: { clientX: number; clientY: number }): PdfPoint | null => {
      const layer = layerRef.current
      return layer ? toPdfPoint(viewport, layer, event.clientX, event.clientY) : null
    },
    [viewport],
  )

  // ---- Committing editors -------------------------------------------------------

  const commitEditor = useCallback(async (): Promise<boolean> => {
    const current = openEditorRef.current
    if (!current) return true
    if (current.kind === 'text-box') {
      const text = current.text.replace(/\s+$/, '')
      if (fontTools && !fontTools.canEncode(HELVETICA, text)) {
        editor.setMessage(t('pdf.edit.unsupportedCharacters'))
        return false
      }
      if (current.existing) {
        if (text.length === 0) editor.removeById(current.existing.id)
        else if (text !== current.existing.text) editor.updateDraft({ ...current.existing, text })
      } else if (text.length > 0) {
        editor.addDraft({
          kind: 'text',
          id: newDraftId('text'),
          page: sourcePageNumber,
          anchor: current.anchor,
          rotation: current.rotation,
          text,
          fontSize: current.fontSize,
          color: current.color,
        })
      }
      editor.setMessage(null)
      setOpenEditor(null)
      return true
    }

    // A paragraph: read what was typed, then check every character can be drawn.
    const value = readParagraphRef.current?.() ?? current.initial
    setOpenEditor(null)
    const id = current.existing?.id ?? newDraftId('text-edit')
    let edit: TextEdit
    let unchanged: boolean
    if (current.paragraph) {
      edit = editParagraph(id, sourcePageNumber, current.paragraph, value.spans, value.styles)
      unchanged = value.styles.length === current.paragraph.styles.length && spansUnchanged(current.paragraph, value.spans)
    } else {
      const text = spansText(value.spans).replace(/\s+$/, '')
      const style = value.styles[0]
      edit = {
        ...newTextEdit(id, sourcePageNumber, current.newOrigin ?? current.corner, current.direction, text, style.font, style.size, { r: 0, g: 0, b: 0 }),
        styles: value.styles,
        spans: value.spans,
      }
      // New text left empty is no text at all.
      unchanged = text.length === 0
    }
    const missing = unchanged || !textAnalysis ? null : await textAnalysis.undrawable(edit)
    if (missing !== null) {
      editor.setMessage(t('pdf.edit.unsupportedCharacter', { character: missing }))
      setOpenEditor({ ...current, initial: value })
      return false
    }
    editor.setMessage(null)
    editor.editText(edit, unchanged)
    return true
  }, [editor, fontTools, setOpenEditor, sourcePageNumber, t, textAnalysis])

  // Commit an open editor when a save asks for it, and when this page leaves
  // the render window — unmounting fires no blur, so the draft would just go.
  const commitEditorRef = useRef(commitEditor)
  useEffect(() => {
    commitEditorRef.current = commitEditor
  }, [commitEditor])
  useEffect(() => {
    const unregister = registerFlush?.(() => commitEditorRef.current().then(() => undefined))
    return () => {
      unregister?.()
      void commitEditorRef.current()
    }
  }, [registerFlush])

  /** Commits on losing focus — but not when the whole window lost it (another app came forward). */
  const commitOnBlur = useCallback(() => {
    if (document.hasFocus()) void commitEditor()
  }, [commitEditor])

  const cancelEditor = useCallback(() => {
    editor.setMessage(null)
    setOpenEditor(null)
  }, [editor, setOpenEditor])

  const openTextBoxEditor = useCallback(
    (existing: TextBoxAnnotation | null, anchor: PdfPoint) => {
      setOpenEditor({
        kind: 'text-box',
        existing,
        anchor: existing?.anchor ?? anchor,
        rotation: existing?.rotation ?? documentRotation,
        fontSize: existing?.fontSize ?? fontSize,
        color: existing?.color ?? toolColors.text,
        text: existing?.text ?? '',
      })
    },
    [documentRotation, fontSize, setOpenEditor, toolColors.text],
  )

  const openParagraphEditor = useCallback(
    (paragraph: TextParagraph) => {
      if (paragraph.blocked) {
        const reasons = {
          invisible: 'pdf.edit.scannedText',
          font: 'pdf.edit.uneditableFont',
          structure: 'pdf.edit.uneditableStructure',
          shared: 'pdf.edit.sharedText',
        } as const
        editor.setMessage(t(reasons[paragraph.blocked]))
        return
      }
      editor.setMessage(null)
      const existing = textEdits.find((edit) => edit.paragraphId === paragraph.id) ?? null
      const top = paragraph.bounds[0]
      const height = Math.hypot(paragraph.bounds[3].x - top.x, paragraph.bounds[3].y - top.y)
      setOpenEditor({
        kind: 'paragraph',
        paragraph,
        existing,
        initial: { spans: existing?.spans ?? paragraph.spans, styles: existing?.styles ?? paragraph.styles },
        corner: top,
        direction: paragraph.direction,
        width: paragraph.lineCount > 1 ? paragraph.width : null,
        height: Math.max(height, paragraph.styles[0].size * 1.2),
        newOrigin: null,
      })
    },
    [editor, setOpenEditor, t, textEdits],
  )

  const openNewTextEditor = useCallback(
    (point: PdfPoint, existing: TextEdit | null) => {
      const direction = existing?.direction ?? rotateVector({ x: 1, y: 0 }, documentRotation)
      const size = existing?.styles[0].size ?? fontSize
      const normal = { x: -direction.y, y: direction.x }
      // The click marks the top-left of the text; the baseline sits one ascent below it.
      const origin = existing?.origin ?? { x: point.x - normal.x * size * 0.8, y: point.y - normal.y * size * 0.8 }
      const color = toolColors.text
      const style = newTextEdit('probe', sourcePageNumber, origin, direction, '', { family: editor.textFamily, bold: false, italic: false }, size, color).styles[0]
      setOpenEditor({
        kind: 'paragraph',
        paragraph: null,
        existing,
        initial: existing ? { spans: existing.spans, styles: existing.styles } : { spans: [], styles: [style] },
        corner: existing ? existing.bounds[0] : point,
        direction,
        width: null,
        height: size * 1.2,
        newOrigin: origin,
      })
    },
    [documentRotation, editor.textFamily, fontSize, setOpenEditor, sourcePageNumber, toolColors.text],
  )

  const findParagraph = useCallback(
    (point: PdfPoint): TextParagraph | null => paragraphs?.find((paragraph) => insidePolygon(point, paragraph.bounds)) ?? null,
    [paragraphs],
  )

  // ---- Pointer input on the layer itself (text, ink and edit-text tools) ---------------

  const handleLayerPointerDown = useCallback(
    (event: ReactPointerEvent<HTMLDivElement>) => {
      if (event.button !== 0 || event.target !== event.currentTarget) return
      const point = pointFromEvent(event)
      if (!point) return
      if (openEditorRef.current) {
        // A click outside an open editor commits it, and goes no further.
        event.preventDefault()
        void commitEditor()
        return
      }
      setSelectedId(null)
      if (tool === 'text') {
        event.preventDefault()
        openTextBoxEditor(null, point)
      } else if (tool === 'ink') {
        event.preventDefault()
        event.currentTarget.setPointerCapture?.(event.pointerId)
        setInkPoints([point])
      } else if (tool === 'edit-text') {
        event.preventDefault()
        const added = textEdits.find((edit) => edit.paragraphId === null && insidePolygon(point, edit.bounds))
        if (added) {
          openNewTextEditor(point, added)
          return
        }
        const paragraph = findParagraph(point)
        if (paragraph) openParagraphEditor(paragraph)
        else if (paragraphs !== null) openNewTextEditor(point, null)
      }
    },
    [commitEditor, findParagraph, openNewTextEditor, openParagraphEditor, openTextBoxEditor, paragraphs, pointFromEvent, setSelectedId, textEdits, tool],
  )

  const handleLayerPointerMove = useCallback(
    (event: ReactPointerEvent<HTMLDivElement>) => {
      const point = pointFromEvent(event)
      if (!point) return
      if (inkPoints) {
        const last = inkPoints[inkPoints.length - 1]
        if (Math.hypot(point.x - last.x, point.y - last.y) >= INK_MIN_STEP) setInkPoints([...inkPoints, point])
      } else if (tool === 'edit-text' && !openEditor) {
        const paragraph = findParagraph(point)
        setHoverParagraph((current) => (current?.id === paragraph?.id ? current : paragraph))
      }
    },
    [findParagraph, inkPoints, openEditor, pointFromEvent, tool],
  )

  const handleLayerPointerUp = useCallback(() => {
    if (!inkPoints) return
    editor.addDraft({
      kind: 'ink',
      id: newDraftId('ink'),
      page: sourcePageNumber,
      points: inkPoints,
      color: toolColors.ink,
      width: inkWidth,
    })
    setInkPoints(null)
  }, [editor, inkPoints, inkWidth, sourcePageNumber, toolColors.ink])

  // ---- Pointer input on shapes (select tool: select, drag, double-click to edit) -------

  const handleShapePointerDown = useCallback(
    (event: ReactPointerEvent<SVGElement>, id: string, draggable: boolean) => {
      if (event.button !== 0 || !isEditing || (tool !== 'select' && tool !== 'text')) return
      event.stopPropagation()
      event.preventDefault()
      setSelectedId(id)
      const point = pointFromEvent(event)
      if (draggable && point) {
        event.currentTarget.setPointerCapture?.(event.pointerId)
        setDrag({ id, start: point, delta: { x: 0, y: 0 } })
      }
    },
    [isEditing, pointFromEvent, setSelectedId, tool],
  )

  const handleShapePointerMove = useCallback(
    (event: ReactPointerEvent<SVGElement>) => {
      if (!drag) return
      const point = pointFromEvent(event)
      if (point) setDrag({ ...drag, delta: { x: point.x - drag.start.x, y: point.y - drag.start.y } })
    },
    [drag, pointFromEvent],
  )

  const handleShapePointerUp = useCallback(() => {
    if (!drag) return
    const moved = Math.hypot(drag.delta.x, drag.delta.y) > 0.5
    const annotation = annotations.find((candidate) => candidate.id === drag.id)
    if (moved && annotation) editor.updateDraft(translateAnnotation(annotation, drag.delta.x, drag.delta.y))
    setDrag(null)
  }, [annotations, drag, editor])

  const handleShapeDoubleClick = useCallback(
    (target: PdfAnnotationDraft | TextEdit) => {
      if (!isEditing) return
      if (target.kind === 'text') {
        openTextBoxEditor(target, target.anchor)
      } else if (target.kind === 'text-edit') {
        if (target.paragraphId === null) {
          openNewTextEditor(target.origin, target)
          return
        }
        const paragraph = paragraphs?.find((candidate) => candidate.id === target.paragraphId)
        if (paragraph) openParagraphEditor(paragraph)
      }
    },
    [isEditing, openNewTextEditor, openParagraphEditor, openTextBoxEditor, paragraphs],
  )

  const handleEditorKeyDown = useCallback(
    (event: ReactKeyboardEvent<HTMLTextAreaElement>) => {
      if (event.key === 'Escape') {
        event.preventDefault()
        event.stopPropagation()
        cancelEditor()
      } else if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
        event.preventDefault()
        void commitEditor()
      }
    },
    [cancelEditor, commitEditor],
  )

  /**
   * The paper colour behind each paragraph that needs a cover right now (a
   * sketched edit, or the paragraph open in the editor), read off the canvas
   * only for those: reading pixels back is slow.
   */
  const coverColors = useMemo(() => {
    const colors = new Map<string, string>()
    const openParagraph = openEditor?.kind === 'paragraph' ? openEditor.paragraph : null
    const editingId = openEditor?.existing?.id
    for (const edit of textEdits) {
      if (edit.paragraphId === null) continue
      const needed = !previewCurrent || edit.id === editingId || edit.paragraphId === openParagraph?.id
      if (needed) colors.set(edit.paragraphId, cssColor(sampleBackground(info, edit.bounds)))
    }
    if (openParagraph && !colors.has(openParagraph.id)) {
      colors.set(openParagraph.id, cssColor(sampleBackground(info, openParagraph.bounds)))
    }
    return colors
  }, [info, openEditor, previewCurrent, textEdits])
  const coverColor = (key: string): string => coverColors.get(key) ?? 'rgb(255, 255, 255)'

  const transform = (viewport as { transform?: ReadonlyArray<number> }).transform
  if (!transform || transform.length !== 6) return null
  const matrix = `matrix(${transform.join(' ')})`
  const interactive = isEditing && (tool === 'text' || tool === 'ink' || tool === 'edit-text')
  const shapesInteractive = isEditing && (tool === 'select' || tool === 'text')
  const shapeProps = (id: string, draggable: boolean, target: PdfAnnotationDraft | TextEdit) => ({
    'data-edit-id': id,
    className: `pdf-edit-layer__shape${selectedId === id ? ' pdf-edit-layer__shape--selected' : ''}`,
    style: { pointerEvents: shapesInteractive ? ('all' as const) : ('none' as const) },
    onPointerDown: (event: ReactPointerEvent<SVGElement>) => handleShapePointerDown(event, id, draggable),
    onPointerMove: handleShapePointerMove,
    onPointerUp: handleShapePointerUp,
    onDoubleClick: () => handleShapeDoubleClick(target),
  })

  const withDrag = (annotation: PdfAnnotationDraft): PdfAnnotationDraft =>
    drag && drag.id === annotation.id ? translateAnnotation(annotation, drag.delta.x, drag.delta.y) : annotation

  const highlights = annotations.filter((annotation) => annotation.kind === 'highlight').map(withDrag)
  const others = annotations.filter((annotation) => annotation.kind !== 'highlight').map(withDrag)
  const editingId = openEditor?.existing?.id
  const editingParagraphId = openEditor?.kind === 'paragraph' ? openEditor.paragraph?.id : undefined
  const selected = annotations.find((annotation) => annotation.id === selectedId)
  const selectedTextEdit = textEdits.find((edit) => edit.id === selectedId)
  const polygon = (corners: ReadonlyArray<PdfPoint>) => corners.map((p) => `${p.x},${p.y}`).join(' ')

  return (
    <>
      <svg
        className="pdf-edit-layer__highlights"
        width={viewport.width}
        height={viewport.height}
        aria-hidden="true"
      >
        <g transform={matrix}>
          {highlights.map((annotation) =>
            annotation.kind === 'highlight' ? (
              <g key={annotation.id} fill={cssColor(annotation.color)}>
                {annotation.rects.map((rect, index) => (
                  <rect key={index} x={rect.x} y={rect.y} width={rect.width} height={rect.height} />
                ))}
              </g>
            ) : null,
          )}
        </g>
      </svg>

      <div
        ref={layerRef}
        className={`pdf-edit-layer pdf-edit-layer--${isEditing ? tool : 'view'}`}
        style={{ pointerEvents: interactive ? 'auto' : 'none' }}
        onPointerDown={handleLayerPointerDown}
        onPointerMove={handleLayerPointerMove}
        onPointerUp={handleLayerPointerUp}
        onPointerLeave={() => setHoverParagraph(null)}
      >
        <svg width={viewport.width} height={viewport.height} className="pdf-edit-layer__svg">
          <g transform={matrix}>
            {/* Text edits: a quick stand-in until the preview has them, then just a hit area. */}
            {textEdits.map((edit) => {
              const hidden = edit.id === editingId || (edit.paragraphId !== null && edit.paragraphId === editingParagraphId)
              const sketch = !previewCurrent && !hidden
              const style = edit.styles[edit.spans[0]?.style ?? 0] ?? edit.styles[0]
              const normal = { x: -edit.direction.y, y: edit.direction.x }
              const angle = (Math.atan2(edit.direction.y, edit.direction.x) * 180) / Math.PI
              const lines = spansText(edit.spans).split('\n')
              return (
                <g key={edit.id}>
                  {(sketch || hidden) && edit.paragraphId !== null && (
                    <polygon points={polygon(edit.bounds)} fill={coverColor(edit.paragraphId)} />
                  )}
                  {sketch &&
                    style &&
                    lines.map((line, index) => (
                      <text
                        key={index}
                        transform={`translate(${edit.origin.x - normal.x * edit.leading * index} ${edit.origin.y - normal.y * edit.leading * index}) rotate(${angle}) scale(1 -1)`}
                        fontSize={style.size}
                        fill={cssFill(style)}
                        fontFamily={cssFontFamily(style.font.family)}
                        fontWeight={style.font.bold ? 700 : 400}
                        fontStyle={style.font.italic ? 'italic' : 'normal'}
                        style={{ whiteSpace: 'pre', pointerEvents: 'none' }}
                      >
                        {line.replace(/\t/g, '    ')}
                      </text>
                    ))}
                  <polygon {...shapeProps(edit.id, false, edit)} points={polygon(edit.bounds)} fill="transparent" />
                </g>
              )
            })}

            {/* While a paragraph is open, cover the original under the editor. */}
            {openEditor?.kind === 'paragraph' && openEditor.paragraph && !openEditor.existing && (
              <polygon points={polygon(openEditor.paragraph.bounds)} fill={coverColor(openEditor.paragraph.id)} />
            )}

            {others.map((annotation) => {
              if (annotation.id === editingId) return null
              if (annotation.kind === 'ink') {
                const points = annotation.points.map((p) => `${p.x},${p.y}`).join(' ')
                const hitProps = shapeProps(annotation.id, true, annotation)
                return (
                  <g key={annotation.id}>
                    <polyline
                      points={points}
                      fill="none"
                      stroke={cssColor(annotation.color)}
                      strokeWidth={annotation.width}
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      style={{ pointerEvents: 'none' }}
                    />
                    {/* A wider, invisible twin takes the clicks: a 1pt stroke is too thin to hit. */}
                    <polyline
                      {...hitProps}
                      style={{ pointerEvents: shapesInteractive ? 'stroke' : 'none' }}
                      points={points}
                      fill="none"
                      stroke="transparent"
                      strokeWidth={Math.max(annotation.width, 8 / viewport.scale)}
                      strokeLinecap="round"
                      strokeLinejoin="round"
                    />
                  </g>
                )
              }
              if (annotation.kind === 'text') {
                const layout = layoutTextBox(annotation.text, annotation.fontSize, (line) => measure(line, annotation.fontSize))
                return (
                  <g
                    key={annotation.id}
                    transform={`translate(${annotation.anchor.x} ${annotation.anchor.y}) rotate(${annotation.rotation}) scale(1 -1)`}
                  >
                    <rect
                      {...shapeProps(annotation.id, true, annotation)}
                      x={0}
                      y={0}
                      width={layout.width}
                      height={layout.height}
                      fill="transparent"
                    />
                    {layout.lines.map((line, index) => (
                      <text
                        key={index}
                        x={TEXT_BOX_PADDING}
                        y={layout.baselines[index]}
                        fontSize={annotation.fontSize}
                        fill={cssColor(annotation.color)}
                        fontFamily={cssFontFamily('helvetica')}
                        style={{ whiteSpace: 'pre', pointerEvents: 'none' }}
                        {...(fontTools && line.length > 0
                          ? { textLength: measure(line, annotation.fontSize), lengthAdjust: 'spacingAndGlyphs' }
                          : {})}
                      >
                        {line}
                      </text>
                    ))}
                  </g>
                )
              }
              return null
            })}

            {/* Highlights are drawn in the multiply layer; here they only take pointer input. */}
            {highlights.map((annotation) =>
              annotation.kind === 'highlight' ? (
                <g key={annotation.id} {...shapeProps(annotation.id, true, annotation)} fill="transparent">
                  {annotation.rects.map((rect, index) => (
                    <rect key={index} x={rect.x} y={rect.y} width={rect.width} height={rect.height} />
                  ))}
                </g>
              ) : null,
            )}

            {inkPoints && (
              <polyline
                points={inkPoints.map((p) => `${p.x},${p.y}`).join(' ')}
                fill="none"
                stroke={cssColor(toolColors.ink)}
                strokeWidth={inkWidth}
                strokeLinecap="round"
                strokeLinejoin="round"
              />
            )}

            {tool === 'edit-text' && isEditing && hoverParagraph && !openEditor && (
              <polygon className="pdf-edit-layer__hover" points={polygon(hoverParagraph.bounds)} vectorEffect="non-scaling-stroke" />
            )}

            {isEditing && (selected || selectedTextEdit) && (() => {
              const rect = selected ? boundsOf(withDrag(selected), measure) : boundingRect((selectedTextEdit as TextEdit).bounds)
              return (
                <rect
                  className="pdf-edit-layer__selection"
                  x={rect.x - 1}
                  y={rect.y - 1}
                  width={rect.width + 2}
                  height={rect.height + 2}
                  vectorEffect="non-scaling-stroke"
                />
              )
            })()}
          </g>
        </svg>

        {openEditor?.kind === 'text-box' && (() => {
          const position = toScreenPoint(viewport, openEditor.anchor)
          const screenTurn = normalizeRotation(viewport.rotation - openEditor.rotation)
          const lines = openEditor.text.split('\n')
          const widest = Math.max(...lines.map((line) => measure(line, openEditor.fontSize)), openEditor.fontSize * 4)
          return (
            <textarea
              autoFocus
              className="pdf-edit-layer__text-editor"
              aria-label={t('pdf.edit.textBoxAria')}
              value={openEditor.text}
              rows={Math.max(1, lines.length)}
              spellCheck
              style={{
                left: position.left,
                top: position.top,
                width: (widest + TEXT_BOX_PADDING * 2 + openEditor.fontSize) * viewport.scale,
                fontSize: openEditor.fontSize * viewport.scale,
                lineHeight: 1.2,
                padding: TEXT_BOX_PADDING * viewport.scale,
                color: cssColor(openEditor.color),
                transform: screenTurn === 0 ? undefined : `rotate(${screenTurn}deg)`,
              }}
              onChange={(event) => setOpenEditor({ ...openEditor, text: event.target.value })}
              onKeyDown={handleEditorKeyDown}
              onBlur={commitOnBlur}
            />
          )
        })()}

        {openEditor?.kind === 'paragraph' && (() => {
          const corner = toScreenPoint(viewport, openEditor.corner)
          const scale = viewport.scale
          return (
            <ParagraphEditor
              key={`${openEditor.paragraph?.id ?? 'new'}:${openEditor.existing?.id ?? ''}`}
              initial={openEditor.initial}
              box={{
                left: corner.left,
                top: corner.top,
                width: openEditor.width === null ? 0 : openEditor.width * scale,
                height: openEditor.height * scale,
              }}
              wrap={openEditor.width !== null}
              scale={scale}
              screenTurn={screenAngle(viewport, openEditor.corner, openEditor.direction)}
              label={t('pdf.edit.paragraphAria')}
              onCommit={() => void commitEditor()}
              onCancel={cancelEditor}
              onBlur={commitOnBlur}
              registerRead={registerRead}
            />
          )
        })()}
      </div>
    </>
  )
}
