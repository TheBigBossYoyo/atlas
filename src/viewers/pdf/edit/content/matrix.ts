/**
 * PDF transformation matrices: `[a b c d e f]`, applied to row vectors
 * (`[x y 1] × M`), which is the convention every PDF operator uses — so
 * `multiply(A, B)` means "A, then B".
 */

export type Matrix = readonly [number, number, number, number, number, number]

export const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0]

export function multiply(m1: Matrix, m2: Matrix): Matrix {
  return [
    m1[0] * m2[0] + m1[1] * m2[2],
    m1[0] * m2[1] + m1[1] * m2[3],
    m1[2] * m2[0] + m1[3] * m2[2],
    m1[2] * m2[1] + m1[3] * m2[3],
    m1[4] * m2[0] + m1[5] * m2[2] + m2[4],
    m1[4] * m2[1] + m1[5] * m2[3] + m2[5],
  ]
}

export function translate(tx: number, ty: number): Matrix {
  return [1, 0, 0, 1, tx, ty]
}

export function apply(m: Matrix, x: number, y: number): { x: number; y: number } {
  return { x: x * m[0] + y * m[2] + m[4], y: x * m[1] + y * m[3] + m[5] }
}

/** The inverse, or null for a degenerate (zero-area) matrix. */
export function invert(m: Matrix): Matrix | null {
  const det = m[0] * m[3] - m[1] * m[2]
  if (Math.abs(det) < 1e-12) return null
  return [
    m[3] / det,
    -m[1] / det,
    -m[2] / det,
    m[0] / det,
    (m[2] * m[5] - m[3] * m[4]) / det,
    (m[1] * m[4] - m[0] * m[5]) / det,
  ]
}

/** Length of the image of a unit x vector (ignoring translation). */
export function xScale(m: Matrix): number {
  return Math.hypot(m[0], m[1])
}

/** Length of the image of a unit y vector (ignoring translation). */
export function yScale(m: Matrix): number {
  return Math.hypot(m[2], m[3])
}
