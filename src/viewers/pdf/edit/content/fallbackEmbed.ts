/**
 * A bundled fallback font embedded into a PDF (PDF-EDIT-2), written by hand
 * as a Type 0 font over a CIDFontType2 with an Identity CID-to-GID map, so a
 * CID in the content stream IS the glyph ID in the embedded font program.
 *
 * Text is encoded as it is written, recording each glyph used; `finalize`
 * then embeds a subset holding exactly those glyphs (see `trueTypeSubset.ts`
 * for why pdf-lib's own subsetting isn't used) plus their widths and a
 * ToUnicode map, so the new text can be selected, searched and copied.
 */

import type { PDFContext, PDFRef } from 'pdf-lib'

import type { PdfLib } from '../writePdf'
import type { FontStyle } from '../types'
import { subsetTrueType } from './trueTypeSubset'

/** The fontkit font API this needs. */
export type FontkitFont = {
  readonly unitsPerEm: number
  readonly postscriptName?: string | null
  readonly ascent: number
  readonly descent: number
  readonly capHeight?: number
  readonly italicAngle?: number
  readonly bbox: { readonly minX: number; readonly minY: number; readonly maxX: number; readonly maxY: number }
  glyphForCodePoint(codePoint: number): { readonly id: number; readonly advanceWidth: number }
  hasGlyphForCodePoint(codePoint: number): boolean
}

export class EmbeddedFallbackFont {
  readonly ref: PDFRef
  private readonly lib: PdfLib
  private readonly context: PDFContext
  private readonly bytes: Uint8Array
  private readonly font: FontkitFont
  private readonly style: FontStyle
  private readonly used = new Map<number, string>()
  private readonly widths = new Map<number, number>()
  private finalized = false

  constructor(lib: PdfLib, context: PDFContext, bytes: Uint8Array, font: FontkitFont, style: FontStyle) {
    this.lib = lib
    this.context = context
    this.bytes = bytes
    this.font = font
    this.style = style
    this.ref = context.nextRef()
  }

  covers(codePoint: number): boolean {
    return this.font.hasGlyphForCodePoint(codePoint)
  }

  /** Advance of `text` at `size`, in the same units as `size`. */
  width(text: string, size: number): number {
    let total = 0
    for (const ch of text) total += this.font.glyphForCodePoint(ch.codePointAt(0) ?? 0).advanceWidth
    return (total * size) / this.font.unitsPerEm
  }

  /** A hex string of glyph IDs for `text`, recording each glyph for the subset. */
  encode(text: string): string {
    if (this.finalized) throw new Error('This font has already been written into the file.')
    let hex = ''
    for (const ch of text) {
      const glyph = this.font.glyphForCodePoint(ch.codePointAt(0) ?? 0)
      if (!this.used.has(glyph.id)) {
        this.used.set(glyph.id, ch)
        this.widths.set(glyph.id, Math.round((glyph.advanceWidth * 1000) / this.font.unitsPerEm))
      }
      hex += glyph.id.toString(16).padStart(4, '0')
    }
    return `<${hex}>`
  }

  /** Writes the font into the file. Call once, after every piece of text has been encoded. */
  finalize(): void {
    if (this.finalized) return
    this.finalized = true
    // Tried for coverage and not chosen: nothing refers to it, so write nothing.
    if (this.used.size === 0) return
    const { lib, context, font, style } = this
    const ids = [...this.used.keys()].sort((a, b) => a - b)
    const subset = subsetTrueType(this.bytes, new Set(ids))
    const scale = 1000 / font.unitsPerEm
    const name = `${subsetTag(ids)}+${(font.postscriptName ?? 'AtlasFallback').replace(/[^A-Za-z0-9-]/g, '')}`

    const fontFile = context.register(context.flateStream(subset, { Length1: subset.length }))
    const flags = (style.family === 'courier' ? 1 : 0) | (style.family === 'times' ? 2 : 0) | 4 | (style.italic ? 64 : 0)
    const descriptor = context.register(
      context.obj({
        Type: 'FontDescriptor',
        FontName: name,
        Flags: flags,
        FontBBox: [font.bbox.minX, font.bbox.minY, font.bbox.maxX, font.bbox.maxY].map((v) => Math.round(v * scale)),
        ItalicAngle: font.italicAngle ?? 0,
        Ascent: Math.round(font.ascent * scale),
        Descent: Math.round(font.descent * scale),
        CapHeight: Math.round((font.capHeight ?? font.ascent) * scale),
        StemV: style.bold ? 120 : 80,
        FontFile2: fontFile,
      }),
    )
    const w: Array<number | number[]> = []
    for (const id of ids) w.push(id, [this.widths.get(id) ?? 0])
    const cidFont = context.register(
      context.obj({
        Type: 'Font',
        Subtype: 'CIDFontType2',
        BaseFont: name,
        CIDSystemInfo: { Registry: lib.PDFString.of('Adobe'), Ordering: lib.PDFString.of('Identity'), Supplement: 0 },
        FontDescriptor: descriptor,
        W: w,
        CIDToGIDMap: 'Identity',
      }),
    )
    const toUnicode = context.register(
      context.flateStream(new TextEncoder().encode(toUnicodeCMap(this.used)), {}),
    )
    context.assign(
      this.ref,
      context.obj({
        Type: 'Font',
        Subtype: 'Type0',
        BaseFont: name,
        Encoding: 'Identity-H',
        DescendantFonts: [cidFont],
        ToUnicode: toUnicode,
      }),
    )
  }
}

/** Six capital letters, as subset font names require, derived from the glyph set. */
function subsetTag(ids: ReadonlyArray<number>): string {
  let hash = 2166136261
  for (const id of ids) hash = Math.imul(hash ^ id, 16777619) >>> 0
  let tag = ''
  for (let i = 0; i < 6; i += 1) {
    tag += String.fromCharCode(65 + (hash % 26))
    hash = Math.floor(hash / 26) + i * 7919
  }
  return tag
}

function utf16Hex(text: string): string {
  let out = ''
  for (let i = 0; i < text.length; i += 1) out += text.charCodeAt(i).toString(16).padStart(4, '0')
  return out
}

function toUnicodeCMap(used: ReadonlyMap<number, string>): string {
  const entries = [...used].sort((a, b) => a[0] - b[0])
  const chunks: string[] = []
  for (let i = 0; i < entries.length; i += 100) {
    const chunk = entries.slice(i, i + 100)
    chunks.push(
      `${chunk.length} beginbfchar\n${chunk.map(([id, text]) => `<${id.toString(16).padStart(4, '0')}> <${utf16Hex(text)}>`).join('\n')}\nendbfchar`,
    )
  }
  return [
    '/CIDInit /ProcSet findresource begin',
    '12 dict begin',
    'begincmap',
    '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def',
    '/CMapName /Adobe-Identity-UCS def',
    '/CMapType 2 def',
    '1 begincodespacerange',
    '<0000> <ffff>',
    'endcodespacerange',
    ...chunks,
    'endcmap',
    'CMapName currentdict /CMap defineresource pop',
    'end',
    'end',
  ].join('\n')
}
