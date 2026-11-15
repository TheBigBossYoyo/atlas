/**
 * Building `TextEdit`s from what the editor knows (PDF-EDIT-2), and the
 * bundled fallback fonts' file names.
 */

import type { FontStyle, PdfPoint, RgbColor, TextEdit } from '../types'
import type { SpanStyle, StyledSpan, TextParagraph } from './paragraphs'

/** The edit that turns `paragraph` into `spans`. */
export function editParagraph(
  id: string,
  page: number,
  paragraph: TextParagraph,
  spans: ReadonlyArray<StyledSpan>,
  styles: ReadonlyArray<SpanStyle> = paragraph.styles,
): TextEdit {
  return {
    kind: 'text-edit',
    id,
    page,
    paragraphId: paragraph.id,
    streamKey: paragraph.streamKey,
    pieces: paragraph.pieces,
    insertAfterOp: paragraph.insertAfterOp,
    origin: paragraph.origin,
    direction: paragraph.direction,
    leading: paragraph.leading,
    // A one-line paragraph (a heading, a label) keeps its line as typed; a
    // real paragraph re-wraps to its own width.
    width: paragraph.lineCount > 1 ? paragraph.width + 0.5 : null,
    align: paragraph.lineCount > 1 ? paragraph.align : 'left',
    firstLineIndent: paragraph.firstLineIndent,
    tabStops: paragraph.tabStops,
    styles,
    spans,
    bounds: paragraph.bounds,
  }
}

/** Text added where there was none: one style, written at `origin` in a bundled font. */
export function newTextEdit(
  id: string,
  page: number,
  origin: PdfPoint,
  direction: PdfPoint,
  text: string,
  font: FontStyle,
  size: number,
  color: RgbColor,
): TextEdit {
  const style: SpanStyle = {
    fontKey: null,
    fontResource: null,
    font,
    size,
    fill: { space: null, color: `${round(color.r)} ${round(color.g)} ${round(color.b)} rg` },
    charSpacing: 0,
    wordSpacing: 0,
    horizontalScale: 1,
    renderMode: 0,
  }
  const normal = { x: -direction.y, y: direction.x }
  const lines = text.split('\n').length
  const corner = (along: number, across: number) => ({
    x: origin.x + direction.x * along + normal.x * across,
    y: origin.y + direction.y * along + normal.y * across,
  })
  const width = size * 0.55 * Math.max(...text.split('\n').map((line) => line.length), 1)
  return {
    kind: 'text-edit',
    id,
    page,
    paragraphId: null,
    streamKey: 'page',
    pieces: [],
    insertAfterOp: -1,
    origin,
    direction,
    leading: size * 1.2,
    width: null,
    align: 'left',
    firstLineIndent: 0,
    tabStops: [],
    styles: [style],
    spans: [{ text, style: 0 }],
    bounds: [corner(0, size * 0.95), corner(width, size * 0.95), corner(width, -size * (0.25 + 1.2 * (lines - 1))), corner(0, -size * (0.25 + 1.2 * (lines - 1)))],
  }
}

function round(value: number): string {
  return String(Math.round(value * 1000) / 1000)
}

export function spansText(spans: ReadonlyArray<StyledSpan>): string {
  return spans.map((span) => span.text).join('')
}

/** Whether `spans` say exactly what the paragraph already says, style for style. */
export function spansUnchanged(paragraph: TextParagraph, spans: ReadonlyArray<StyledSpan>): boolean {
  const normalize = (list: ReadonlyArray<StyledSpan>) => {
    const merged: StyledSpan[] = []
    for (const span of list) {
      if (span.text.length === 0) continue
      const last = merged[merged.length - 1]
      if (last && last.style === span.style) merged[merged.length - 1] = { text: last.text + span.text, style: span.style }
      else merged.push(span)
    }
    return merged
  }
  const a = normalize(paragraph.spans)
  const b = normalize(spans)
  return a.length === b.length && a.every((span, index) => span.text === b[index].text && span.style === b[index].style)
}

/**
 * The bundled font standing in for a style the document's own font can't
 * draw, each metric-compatible with a common family so replaced text keeps
 * about the original's width: Arimo for Arial/Helvetica, Tinos for Times New
 * Roman, Courier Prime for Courier.
 *
 * Arimo is bundled only as a variable font, and embedding drops its variation
 * tables (see `trueTypeSubset.ts`), leaving its default instance: the regular
 * weight. Bold sans-serif text therefore uses Carlito's bold instead.
 */
export function fallbackFontFile(style: FontStyle): string {
  const { family, bold, italic } = style
  const variant = bold ? (italic ? 'BoldItalic' : 'Bold') : italic ? 'Italic' : 'Regular'
  if (family === 'times') return `Tinos-${variant}.ttf`
  if (family === 'courier') return `CourierPrime-${variant}.ttf`
  if (bold) return `Carlito-${variant}.ttf`
  return italic ? 'Arimo-Italic-wght.ttf' : 'Arimo-wght.ttf'
}
