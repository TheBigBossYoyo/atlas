/**
 * Turns a page's text pieces into editable paragraphs (PDF-EDIT-2).
 *
 * A PDF has no paragraphs — only glyph runs placed at coordinates — so they
 * are rebuilt from geometry, the way an editor like Acrobat's has to:
 *
 *   1. pieces on the same baseline, close enough to read as one line, form a
 *      line (a big gap is a column or table boundary, so it splits);
 *   2. lines that stack at a steady leading, share a left edge and a size,
 *      form a paragraph — ended early by a short line (the last line of a
 *      paragraph rarely reaches the margin) or a line starting with a bullet.
 *
 * Each paragraph carries its text as styled spans (a bold word stays bold),
 * its alignment, and the exact operators to remove when it is replaced.
 */

import { invert, multiply, xScale, type Matrix } from './matrix'
import type { FillColor, PageText, TextPiece } from './interpreter'
import type { FontStyle } from '../types'

export type Point = { readonly x: number; readonly y: number }

export type TextAlign = 'left' | 'center' | 'right' | 'justify'

/** How one stretch of text looks — enough to draw new text the same way. */
export type SpanStyle = {
  /** Identity of the original font, or null for a style that never existed in the file. */
  readonly fontKey: string | null
  /** The font's resource name in the paragraph's stream, when it came from the file. */
  readonly fontResource: string | null
  /** Typeface family/weight/slant — what a fallback font must match. */
  readonly font: FontStyle
  /** Font size in user space. */
  readonly size: number
  readonly fill: FillColor
  /** Character and word spacing in user space units. */
  readonly charSpacing: number
  readonly wordSpacing: number
  readonly horizontalScale: number
  readonly renderMode: number
}

export type StyledSpan = { readonly text: string; readonly style: number }

export type PieceRef = { readonly opIndex: number; readonly element: number }

export type TextParagraph = {
  /** Stable for a given file: stream plus the first piece's operator. */
  readonly id: string
  readonly streamKey: string
  readonly pieces: ReadonlyArray<PieceRef>
  /** Where new content for this paragraph is inserted: after this `ET`. */
  readonly insertAfterOp: number
  /** The paragraph's left edge on the first baseline, and the baseline direction (unit vector), in user space. */
  readonly origin: Point
  readonly direction: Point
  /** Distance between baselines, in user space. */
  readonly leading: number
  /** Widest line, in user space — the width new text wraps to. */
  readonly width: number
  readonly lineCount: number
  readonly align: TextAlign
  /** How far the first line starts right of the others (negative for a hanging first line). */
  readonly firstLineIndent: number
  /**
   * Where text resumed after a wide positioned gap (a bullet's hanging indent,
   * the next table cell, a form label's value), as distances along the
   * baseline from `origin`. The gap reads as a tab in the text, and editing
   * puts the text after it back at the same stop instead of collapsing it to
   * a space.
   */
  readonly tabStops: ReadonlyArray<number>
  readonly styles: ReadonlyArray<SpanStyle>
  readonly spans: ReadonlyArray<StyledSpan>
  /** The paragraph's area: four user-space corners (top-left, top-right, bottom-right, bottom-left). */
  readonly bounds: ReadonlyArray<Point>
  /** Why it can't be edited, or null. */
  readonly blocked: 'invisible' | 'font' | 'structure' | 'shared' | null
}

type Line = {
  readonly pieces: TextPiece[]
  readonly streamKey: string
  readonly angle: number
  /** Position across the baseline (frame y) and the extent along it (frame x). */
  readonly across: number
  readonly left: number
  readonly right: number
  readonly size: number
  readonly text: string
}

const BULLET = /^\s*([•◦▪▫‣⁃●○■□–—\-*]|\(?\d{1,3}[.)]|\(?[a-zA-Z][.)])\s/

function along(p: Point, angle: number): number {
  return p.x * Math.cos(angle) + p.y * Math.sin(angle)
}

function across(p: Point, angle: number): number {
  return -p.x * Math.sin(angle) + p.y * Math.cos(angle)
}

function pieceAngle(piece: TextPiece): number {
  const dx = piece.end.x - piece.start.x
  const dy = piece.end.y - piece.start.y
  if (Math.hypot(dx, dy) > 1e-6) return Math.atan2(dy, dx)
  // A zero-width piece (a lone space at size 0): fall back to the text matrix direction.
  const m = multiply(piece.textMatrix, piece.ctm)
  return Math.atan2(m[1], m[0])
}

function sameAngle(a: number, b: number): boolean {
  let d = Math.abs(a - b) % (2 * Math.PI)
  if (d > Math.PI) d = 2 * Math.PI - d
  return d < 0.02
}

/** Joins two pieces' text, adding the space the layout implied when the PDF drew it as a gap. */
function joinText(previous: string, next: string, gap: number, size: number): string {
  if (previous.length === 0) return next
  const needsSpace = gap > size * 0.15 && !/\s$/.test(previous) && !/^\s/.test(next)
  return needsSpace ? `${previous} ${next}` : `${previous}${next}`
}

function buildLines(pieces: ReadonlyArray<TextPiece>): Line[] {
  const lines: Line[] = []
  const bins = new Map<string, TextPiece[]>()
  for (const piece of pieces) {
    if (piece.text.length === 0 || piece.sizeUser <= 0) continue
    const angle = pieceAngle(piece)
    const binKey = `${piece.streamKey}|${Math.round(angle * 50)}`
    const bin = bins.get(binKey) ?? []
    bin.push(piece)
    bins.set(binKey, bin)
  }
  bins.forEach((bin) => {
    const angle = pieceAngle(bin[0])
    // Cluster by baseline, then sort each baseline left to right.
    const sorted = [...bin].sort((a, b) => across(b.start, angle) - across(a.start, angle))
    const baselines: TextPiece[][] = []
    for (const piece of sorted) {
      const y = across(piece.start, angle)
      const last = baselines[baselines.length - 1]
      if (last && Math.abs(across(last[0].start, angle) - y) <= Math.max(0.5, piece.sizeUser * 0.3)) last.push(piece)
      else baselines.push([piece])
    }
    for (const baseline of baselines) {
      baseline.sort((a, b) => along(a.start, angle) - along(b.start, angle))
      let current: TextPiece[] = []
      let right = -Infinity
      const flush = () => {
        if (current.length === 0) return
        let text = ''
        let previousEnd = -Infinity
        for (const piece of current) {
          text = joinText(text, piece.text, along(piece.start, angle) - previousEnd, piece.sizeUser)
          previousEnd = along(piece.end, angle)
        }
        const size = Math.max(...current.map((piece) => piece.sizeUser))
        lines.push({
          pieces: current,
          streamKey: current[0].streamKey,
          angle,
          across: across(current[0].start, angle),
          left: along(current[0].start, angle),
          right: Math.max(...current.map((piece) => along(piece.end, angle))),
          size,
          text,
        })
        current = []
        right = -Infinity
      }
      for (const piece of baseline) {
        const start = along(piece.start, angle)
        // A gap wider than ~1.5em is a column or table cell boundary.
        if (current.length > 0 && start - right > piece.sizeUser * 1.5) flush()
        current.push(piece)
        right = Math.max(right, along(piece.end, angle))
      }
      flush()
    }
  })
  return lines
}

function styleOf(piece: TextPiece): SpanStyle {
  const state = piece.state
  const rendering = multiply(piece.textMatrix, piece.ctm)
  const userPerTextUnit = xScale(rendering)
  return {
    fontKey: state.font?.key ?? null,
    fontResource: state.fontResource,
    font: state.font?.style ?? { family: 'helvetica', bold: false, italic: false },
    size: piece.sizeUser,
    fill: piece.fill,
    charSpacing: state.charSpacing * userPerTextUnit,
    wordSpacing: state.wordSpacing * userPerTextUnit,
    horizontalScale: state.horizontalScale,
    renderMode: state.renderMode,
  }
}

function styleKey(style: SpanStyle): string {
  return [
    style.fontKey,
    style.size.toFixed(2),
    style.fill.space,
    style.fill.color,
    style.charSpacing.toFixed(3),
    style.wordSpacing.toFixed(3),
    style.horizontalScale,
    style.renderMode,
  ].join('|')
}

function detectAlign(lines: ReadonlyArray<Line>): TextAlign {
  if (lines.length < 2) return 'left'
  const tolerance = lines[0].size * 0.5
  const lefts = lines.map((line) => line.left)
  const rights = lines.map((line) => line.right)
  const centers = lines.map((line) => (line.left + line.right) / 2)
  const spread = (values: number[]) => Math.max(...values) - Math.min(...values)
  const leftAligned = spread(lefts.slice(1)) <= tolerance
  const bodyRights = rights.slice(0, -1)
  if (leftAligned && bodyRights.length >= 1 && spread(bodyRights) <= tolerance && lines.length >= 3) return 'justify'
  if (!leftAligned && spread(centers) <= tolerance) return 'center'
  if (!leftAligned && spread(rights) <= tolerance) return 'right'
  return 'left'
}

function finishParagraph(lines: ReadonlyArray<Line>, sharedStreams: ReadonlySet<string>): TextParagraph {
  const first = lines[0]
  const angle = first.angle
  const direction = { x: Math.cos(angle), y: Math.sin(angle) }
  const normal = { x: -direction.y, y: direction.x }
  const pieces = lines.flatMap((line) => line.pieces)
  const firstPiece = first.pieces[0]
  const size = first.size
  const gaps = lines.slice(1).map((line, index) => lines[index].across - line.across)
  const leading = gaps.length > 0 ? gaps.reduce((a, b) => a + b, 0) / gaps.length : size * 1.2

  const top = first.across + size * 0.95
  const bottom = lines[lines.length - 1].across - lines[lines.length - 1].size * 0.25
  const corner = (a: number, c: number): Point => ({
    x: a * direction.x + c * normal.x,
    y: a * direction.y + c * normal.y,
  })

  const left = Math.min(...lines.map((line) => line.left))
  const right = Math.max(...lines.map((line) => line.right))
  const align = detectAlign(lines)
  // Justified text stretches its word gaps on purpose; only elsewhere is a
  // wide gap a tab.
  const allowTabs = align !== 'justify'
  const tabStops = new Set<number>()

  const styles: SpanStyle[] = []
  const styleIndex = new Map<string, number>()
  const spans: StyledSpan[] = []
  const push = (text: string, style: SpanStyle) => {
    if (text.length === 0) return
    const key = styleKey(style)
    let index = styleIndex.get(key)
    if (index === undefined) {
      index = styles.length
      styles.push(style)
      styleIndex.set(key, index)
    }
    const last = spans[spans.length - 1]
    if (last && last.style === index) spans[spans.length - 1] = { text: last.text + text, style: index }
    else spans.push({ text, style: index })
  }
  lines.forEach((line, lineIndex) => {
    let previousEnd = -Infinity
    line.pieces.forEach((piece) => {
      const start = along(piece.start, angle)
      const style = styleOf(piece)
      const atLineStart = previousEnd === -Infinity
      if (atLineStart && lineIndex > 0) {
        const previousText = spans.map((span) => span.text).join('')
        if (!/[\s-]$/.test(previousText)) push(' ', style)
      } else if (!atLineStart && start - previousEnd > piece.sizeUser * 0.15 && !/^\s/.test(piece.text)) {
        const previousText = spans.map((span) => span.text).join('')
        if (allowTabs && start - previousEnd > piece.sizeUser * 0.5) {
          push('\t', style)
          tabStops.add(Math.round((start - left) * 100) / 100)
        } else if (!/\s$/.test(previousText)) {
          push(' ', style)
        }
      }
      push(piece.text, style)
      previousEnd = along(piece.end, angle)
    })
  })

  let blocked: TextParagraph['blocked'] = null
  if (pieces.some((piece) => piece.invisible)) blocked = 'invisible'
  else if (pieces.some((piece) => !piece.state.font?.editable || piece.state.fontSize * piece.state.horizontalScale === 0)) blocked = 'font'
  else if (pieces.some((piece) => piece.blockEnd < 0) || new Set(pieces.map((piece) => piece.streamKey)).size > 1) blocked = 'structure'
  else if (sharedStreams.has(firstPiece.streamKey)) blocked = 'shared'

  return {
    id: `${firstPiece.streamKey}:${firstPiece.opIndex}:${firstPiece.element}`,
    streamKey: firstPiece.streamKey,
    pieces: pieces.map((piece) => ({ opIndex: piece.opIndex, element: piece.element })),
    insertAfterOp: firstPiece.blockEnd,
    origin: corner(left, first.across),
    firstLineIndent: first.left - left,
    direction,
    leading,
    width: right - left,
    lineCount: lines.length,
    align,
    tabStops: [...tabStops].sort((a, b) => a - b),
    styles,
    spans,
    bounds: [corner(left, top), corner(right, top), corner(right, bottom), corner(left, bottom)],
    blocked,
  }
}

/**
 * Groups a page's text into paragraphs, top to bottom within each text
 * direction. Text in a form XObject drawn more than once on this page, or
 * named in `sharedForms` (drawn on other pages too), comes back blocked as
 * 'shared': rewriting it would change every copy.
 */
export function buildParagraphs(pageText: PageText, sharedForms: ReadonlySet<string> = new Set()): TextParagraph[] {
  const sharedStreams = new Set(sharedForms)
  pageText.formUses.forEach((count, key) => {
    if (count > 1) sharedStreams.add(key)
  })
  const lines = buildLines(pageText.pieces)
  const groups = new Map<string, Line[]>()
  for (const line of lines) {
    const key = `${line.streamKey}|${Math.round(line.angle * 50)}`
    const group = groups.get(key) ?? []
    group.push(line)
    groups.set(key, group)
  }

  const paragraphs: TextParagraph[] = []
  groups.forEach((group) => {
    const pending = [...group].sort((a, b) => b.across - a.across || a.left - b.left)
    const used = new Set<Line>()
    for (const start of pending) {
      if (used.has(start)) continue
      used.add(start)
      const members: Line[] = [start]
      for (;;) {
        const last = members[members.length - 1]
        const width = Math.max(...members.map((line) => line.right)) - Math.min(...members.map((line) => line.left))
        const candidate = pending.find((line) => {
          if (used.has(line) || !sameAngle(line.angle, last.angle)) return false
          const gap = last.across - line.across
          if (gap < last.size * 0.9 || gap > last.size * 2.2) return false
          if (Math.abs(line.size - start.size) > start.size * 0.2) return false
          if (members.length >= 2) {
            const leading = members[members.length - 2].across - last.across
            if (Math.abs(gap - leading) > start.size * 0.25) return false
          }
          // Shares the paragraph's left edge (a first-line indent is allowed).
          const bodyLeft = members.length >= 2 ? members[1].left : last.left
          const leftOk = members.length >= 2
            ? Math.abs(line.left - bodyLeft) <= line.size
            : line.left >= last.left - line.size * 4 && line.left <= last.left + line.size
          if (!leftOk) return false
          // The line above must run to the paragraph's right edge, or it ended the paragraph.
          const paragraphRight = Math.max(last.right, line.right, ...members.map((m) => m.right))
          if (last.right < paragraphRight - Math.max(start.size * 2, width * 0.15)) return false
          if (BULLET.test(line.text)) return false
          return true
        })
        if (!candidate) break
        used.add(candidate)
        members.push(candidate)
      }
      paragraphs.push(finishParagraph(members, sharedStreams))
    }
  })
  return paragraphs
}

/** Text-space-to-user-space matrix helpers the writer shares with the layout. */
export function rotationMatrix(direction: Point, origin: Point): Matrix {
  return [direction.x, direction.y, -direction.y, direction.x, origin.x, origin.y]
}

/** The text matrix that puts text at `frame` in user space, given the CTM in force. */
export function textMatrixFor(frame: Matrix, ctm: Matrix): Matrix | null {
  const inverse = invert(ctm)
  return inverse ? multiply(frame, inverse) : null
}
