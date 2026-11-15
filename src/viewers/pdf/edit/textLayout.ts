/**
 * Geometry shared by the on-screen preview of a pending text edit and the
 * writer that turns it into PDF objects (PDF-EDIT-1). Both sides must agree to
 * the point, or what the user placed is not what gets saved — so the numbers
 * live here, once, and neither side re-derives them.
 *
 * All of it is plain math on PDF user-space coordinates (y up); nothing here
 * knows about pdf.js, pdf-lib or the DOM.
 */

import type { FontStyle, PdfPoint, PdfRect, StandardFontFamily } from './types'

/** Space between a text box's border and its text, in points. */
export const TEXT_BOX_PADDING = 2
/** Baseline-to-baseline distance as a multiple of the font size. */
export const LINE_HEIGHT = 1.2
/** Helvetica's ascender and descender (AFM units / 1000). Text boxes are always Helvetica. */
export const HELVETICA_ASCENT = 0.718
export const HELVETICA_DESCENT = 0.207

export type TextMeasure = (line: string) => number

export type TextBoxLayout = {
  readonly lines: ReadonlyArray<string>
  readonly width: number
  readonly height: number
  /** Distance from the box's top edge down to each line's baseline. */
  readonly baselines: ReadonlyArray<number>
}

/**
 * Room past the last glyph's advance, as a multiple of the font size. Text
 * boxes use Helvetica without embedding it, and readers draw it with whatever
 * substitute they have (Arial, Liberation Sans, …), whose glyphs — an italic
 * or a final "f" especially — can run past Helvetica's advance width. The
 * appearance stream clips at its box, so without this the last letter loses
 * its right edge.
 */
export const TEXT_BOX_OVERHANG = 0.3

/** Lays out a text box's lines; `measure` returns a line's advance width in points at `fontSize`. */
export function layoutTextBox(text: string, fontSize: number, measure: TextMeasure): TextBoxLayout {
  const lines = text.split(/\r\n|\r|\n/)
  const widest = lines.reduce((max, line) => Math.max(max, measure(line)), 0)
  const width = Math.max(widest, fontSize * 0.5) + TEXT_BOX_PADDING * 2 + fontSize * TEXT_BOX_OVERHANG
  const height =
    TEXT_BOX_PADDING * 2 + fontSize * (LINE_HEIGHT * (lines.length - 1) + HELVETICA_ASCENT + HELVETICA_DESCENT)
  const baselines = lines.map((_, index) => TEXT_BOX_PADDING + fontSize * (HELVETICA_ASCENT + LINE_HEIGHT * index))
  return { lines, width, height, baselines }
}

/** Rotates `vector` counter-clockwise by `degrees` (y-up space). */
export function rotateVector(vector: PdfPoint, degrees: number): PdfPoint {
  const radians = (degrees * Math.PI) / 180
  const cos = Math.cos(radians)
  const sin = Math.sin(radians)
  return { x: vector.x * cos - vector.y * sin, y: vector.x * sin + vector.y * cos }
}

/** Maps a point given in a frame at `origin`, turned `degrees` counter-clockwise, to user space. */
export function frameToUserSpace(origin: PdfPoint, degrees: number, local: PdfPoint): PdfPoint {
  const turned = rotateVector(local, degrees)
  return { x: origin.x + turned.x, y: origin.y + turned.y }
}

/** A loop rather than `Math.min(...xs)`: a long freehand stroke can exceed the engine's argument limit. */
export function boundingRect(points: ReadonlyArray<PdfPoint>): PdfRect {
  let minX = Infinity
  let minY = Infinity
  let maxX = -Infinity
  let maxY = -Infinity
  for (const { x, y } of points) {
    if (x < minX) minX = x
    if (y < minY) minY = y
    if (x > maxX) maxX = x
    if (y > maxY) maxY = y
  }
  return { x: minX, y: minY, width: maxX - minX, height: maxY - minY }
}

/**
 * The four user-space corners of a text box whose top-left corner is `anchor`
 * and which is turned `degrees` counter-clockwise: top-left, top-right,
 * bottom-right, bottom-left.
 */
export function textBoxCorners(anchor: PdfPoint, degrees: number, width: number, height: number): PdfPoint[] {
  return [
    { x: 0, y: 0 },
    { x: width, y: 0 },
    { x: width, y: -height },
    { x: 0, y: -height },
  ].map((local) => frameToUserSpace(anchor, degrees, local))
}

/** The 12 standard PDF font names (always available to every reader, never embedded). */
export function standardFontName(style: FontStyle): string {
  const { family, bold, italic } = style
  const table: Record<StandardFontFamily, readonly [string, string, string, string]> = {
    // regular, bold, italic, bold italic
    helvetica: ['Helvetica', 'Helvetica-Bold', 'Helvetica-Oblique', 'Helvetica-BoldOblique'],
    times: ['Times-Roman', 'Times-Bold', 'Times-Italic', 'Times-BoldItalic'],
    courier: ['Courier', 'Courier-Bold', 'Courier-Oblique', 'Courier-BoldOblique'],
  }
  const variants = table[family]
  return variants[(bold ? 1 : 0) + (italic ? 2 : 0)]
}

/** CSS font stack that looks like each standard family on screen. */
export function cssFontFamily(family: StandardFontFamily): string {
  switch (family) {
    case 'times':
      return '"Times New Roman", Times, serif'
    case 'courier':
      return '"Courier New", Courier, monospace'
    default:
      return 'Helvetica, Arial, sans-serif'
  }
}

/**
 * Guesses the closest standard font for a run from what pdf.js knows about it:
 * the generic family it assigns (`textContent.styles[...].fontFamily`) and,
 * when the font has been loaded for rendering, its PostScript name.
 */
export function guessFontStyle(genericFamily: string | undefined, fontName: string | undefined): FontStyle {
  const name = (fontName ?? '').toLowerCase()
  const generic = (genericFamily ?? '').toLowerCase()
  let family: StandardFontFamily = 'helvetica'
  if (generic.includes('monospace') || /courier|mono|consol/.test(name)) {
    family = 'courier'
  } else if (!name.includes('sans')) {
    const genericSerif = generic.includes('serif') && !generic.includes('sans')
    if (genericSerif || /times|serif|georgia|garamond|cambria|roman|minion/.test(name)) family = 'times'
  }
  return {
    family,
    bold: /bold|black|heavy|semibold|demi/.test(name),
    italic: /italic|oblique/.test(name),
  }
}
