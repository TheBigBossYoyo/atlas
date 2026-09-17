/**
 * USR-16 — geometry for the slide editor's selection, hit testing and
 * resizing, including rotated shapes.
 *
 * OOXML rotates a shape around its box CENTER (`a:xfrm@rot`), and CSS does
 * the same for `transform: rotate(...)` with the default origin, so the
 * rendered shape and these helpers agree by construction. Everything here is
 * in slide pixels; the editor divides pointer deltas by the zoom scale before
 * calling in.
 */
import type { SlideTransform } from './SlideDeck.types'

export type Point = { readonly x: number; readonly y: number }
export type Handle = 'n' | 's' | 'e' | 'w' | 'ne' | 'nw' | 'se' | 'sw'
export type DragMode = Handle | 'move'

/** Smallest box the editor will resize a shape down to. */
export const MIN_SHAPE_PX = 8

function center(box: SlideTransform): Point {
  return { x: box.x + box.w / 2, y: box.y + box.h / 2 }
}

/** Rotates `point` by `degrees` around `origin` (positive = clockwise, like OOXML and CSS). */
export function rotateAround(point: Point, origin: Point, degrees: number): Point {
  if (!degrees) return point
  const radians = (degrees * Math.PI) / 180
  const cos = Math.cos(radians)
  const sin = Math.sin(radians)
  const dx = point.x - origin.x
  const dy = point.y - origin.y
  return { x: origin.x + dx * cos - dy * sin, y: origin.y + dx * sin + dy * cos }
}

/** True when a slide-space point is inside the shape, taking its rotation into account. */
export function isPointInShape(point: Point, box: SlideTransform): boolean {
  const local = box.rotationDeg ? rotateAround(point, center(box), -box.rotationDeg) : point
  return local.x >= box.x && local.x <= box.x + box.w && local.y >= box.y && local.y <= box.y + box.h
}

/** Resizes the box in its OWN (unrotated) space, without moving anything. */
function resizeLocal(origin: SlideTransform, handle: DragMode, dx: number, dy: number): SlideTransform {
  if (handle === 'move') return { ...origin, x: origin.x + dx, y: origin.y + dy }
  let { x, y, w, h } = origin
  if (handle.includes('e')) w = Math.max(MIN_SHAPE_PX, origin.w + dx)
  if (handle.includes('s')) h = Math.max(MIN_SHAPE_PX, origin.h + dy)
  if (handle.includes('w')) {
    w = Math.max(MIN_SHAPE_PX, origin.w - dx)
    x = origin.x + origin.w - w
  }
  if (handle.startsWith('n')) {
    h = Math.max(MIN_SHAPE_PX, origin.h - dy)
    y = origin.y + origin.h - h
  }
  return { ...origin, x, y, w, h }
}

/**
 * The box after dragging `handle` by a slide-space delta.
 *
 * For a rotated shape the pointer delta is first rotated into the shape's own
 * space (dragging its "east" handle grows it along ITS east, not the slide's),
 * and the result is then shifted so the edge/corner OPPOSITE the dragged one
 * stays visually put — otherwise resizing a rotated shape makes it drift,
 * because the box grows from a top-left origin while it is drawn around its
 * rotated center. With no rotation both steps are identities.
 */
export function resizeBox(origin: SlideTransform, handle: DragMode, delta: Point): SlideTransform {
  const rotation = origin.rotationDeg ?? 0
  const local = rotation && handle !== 'move' ? rotateAround(delta, { x: 0, y: 0 }, -rotation) : delta
  const resized = resizeLocal(origin, handle, local.x, local.y)
  if (!rotation || handle === 'move') return resized

  const before = center(origin)
  const after = center(resized)
  const drift = { x: before.x - after.x, y: before.y - after.y }
  const rotatedDrift = rotateAround(drift, { x: 0, y: 0 }, rotation)
  return { ...resized, x: resized.x + drift.x - rotatedDrift.x, y: resized.y + drift.y - rotatedDrift.y }
}
