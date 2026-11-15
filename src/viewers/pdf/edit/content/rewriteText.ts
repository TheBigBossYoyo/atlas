/**
 * Applies text edits to a page's drawing instructions (PDF-EDIT-2).
 *
 * For each edited paragraph:
 *
 *   - every operator piece that showed its old text is rewritten into a pure
 *     positioning step of exactly the same width — `(Hello) Tj` becomes
 *     `[-2780] TJ` — so the old glyphs are gone from the file while every
 *     other glyph on the line stays exactly where it was;
 *   - the new text is typeset (wrapped to the paragraph's width, aligned and
 *     justified like the original) and written in a text block of its own
 *     straight after the original block's `ET`, so it shares that block's
 *     clipping and drawing order;
 *   - the text state and fill colour in force at that point are restored
 *     after it, so nothing later on the page draws differently.
 *
 * Fonts: a style keeps the document's own font when that font can encode
 * every character of the new text; otherwise it switches to a bundled font
 * of the same family (see `fallbackFontFile`), embedded as a subset.
 */

import type { PDFContext, PDFDict, PDFDocument, PDFObject, PDFPage, PDFRef } from 'pdf-lib'

import type { PdfLib } from '../writePdf'
import { StaleTextEditError, UnsupportedCharacterError } from '../errors'
import type { FontStyle, TextEdit } from '../types'
import { readFontInfo, type FontInfo } from './fontInfo'
import { interpretPage, type FillColor, type PageText, type StreamInfo, type TextPiece, type TextState } from './interpreter'
import { binaryStringToBytes, formatHexString, formatName, formatNumber, formatToken } from './lexer'
import type { Matrix } from './matrix'
import { rotationMatrix, textMatrixFor, type SpanStyle } from './paragraphs'
import { asDict, resolve } from './pdfObjects'
import { EmbeddedFallbackFont, type FontkitFont } from './fallbackEmbed'

/** Supplies bundled fonts for text the document's own fonts cannot draw. */
export type FallbackFontSource = {
  /** fontkit, for reading the font files (glyph lookup, metrics). */
  readonly fontkit: unknown
  readonly load: (style: FontStyle) => Promise<Uint8Array>
}

export type FallbackFont = EmbeddedFallbackFont

/** Fallback fonts embedded into each document, by style — one per style per file. */
const docFonts = new WeakMap<PDFDocument, Map<string, Promise<EmbeddedFallbackFont>>>()

type FontkitLike = { create(bytes: Uint8Array): FontkitFont }

export async function fallbackFont(
  lib: PdfLib,
  doc: PDFDocument,
  source: FallbackFontSource,
  style: FontStyle,
): Promise<EmbeddedFallbackFont> {
  let fonts = docFonts.get(doc)
  if (!fonts) {
    fonts = new Map()
    docFonts.set(doc, fonts)
  }
  const key = `${style.family}|${style.bold}|${style.italic}`
  let font = fonts.get(key)
  if (!font) {
    font = source.load(style).then(
      (bytes) => new EmbeddedFallbackFont(lib, doc.context, bytes, (source.fontkit as FontkitLike).create(bytes), style),
    )
    fonts.set(key, font)
  }
  return font
}

/**
 * The fallback font for `style` that can draw all of `text`: the style's own
 * family first, then the others (Courier Prime has no Cyrillic, for one).
 * Throws `UnsupportedCharacterError` when none can.
 */
export async function coveringFallback(
  lib: PdfLib,
  doc: PDFDocument,
  source: FallbackFontSource,
  style: FontStyle,
  text: string,
): Promise<EmbeddedFallbackFont> {
  const families: Array<FontStyle['family']> = [style.family, 'helvetica', 'times', 'courier']
  for (const family of families.filter((value, index) => families.indexOf(value) === index)) {
    const font = await fallbackFont(lib, doc, source, { ...style, family })
    const covers = Array.from(text).every((ch) => {
      const codePoint = ch.codePointAt(0) ?? 0
      return codePoint <= 0x20 || font.covers(codePoint)
    })
    if (covers) return font
  }
  throw new UnsupportedCharacterError(text)
}

/** Writes every fallback font used in `doc` into it; call after all text is encoded, before saving. */
export async function finalizeFallbackFonts(doc: PDFDocument): Promise<void> {
  const fonts = docFonts.get(doc)
  if (!fonts) return
  for (const font of await Promise.all(fonts.values())) font.finalize()
}

// ---- Fonts for a style -------------------------------------------------------------

/** How one style's text is written: through the document's own font, or a fallback. */
type StyleWriter = {
  readonly resource: string
  readonly style: SpanStyle
  /** Advance of `text` in user space, spacing included. */
  width(text: string): number
  /** A TJ item for `text`: a string, or for a space the font has no glyph for, an equal-width gap. */
  encode(text: string): string
  /** Word spacing to set: only single-byte fonts honour Tw. */
  readonly wordSpacing: number
}

/** Word-space width, as a fraction of the size, for a font that has no space glyph. */
function syntheticSpace(family: FontStyle['family']): number {
  return family === 'courier' ? 0.6 : family === 'times' ? 0.25 : 0.278
}

function originalWriter(font: FontInfo, resource: string, style: SpanStyle): StyleWriter {
  const wordSpacing = font.composite ? 0 : style.wordSpacing
  // Many subset fonts (TeX output especially) carry no space glyph: their
  // producers placed words with positioning gaps. Spaces are written the same way.
  const hasSpace = font.encode(' ') !== null
  const gapFor = (text: string) => text.length * syntheticSpace(font.style.family) * style.size
  return {
    resource,
    style,
    wordSpacing,
    width(text) {
      if (/^ +$/.test(text) && !hasSpace) return gapFor(text) * style.horizontalScale
      const codes = font.encode(text) ?? []
      let total = 0
      for (const code of codes) {
        const isSpace = code.length === 1 && code.code === 32
        total += (font.width(code.code) / 1000) * style.size + style.charSpacing + (isSpace ? wordSpacing : 0)
      }
      return total * style.horizontalScale
    },
    encode(text) {
      if (/^ +$/.test(text) && !hasSpace) return formatNumber((-gapFor(text) * 1000) / style.size)
      const codes = font.encode(text) ?? []
      return formatHexString(codes.flatMap((code) => font.bytesOf(code)))
    },
  }
}

function fallbackWriter(font: FallbackFont, resource: string, style: SpanStyle): StyleWriter {
  return {
    resource,
    style,
    wordSpacing: 0,
    width(text) {
      const glyphs = Array.from(text).length
      return (font.width(text, style.size) + style.charSpacing * glyphs) * style.horizontalScale
    },
    encode(text) {
      return font.encode(text)
    },
  }
}

// ---- Typesetting ---------------------------------------------------------------------

type Token = {
  readonly text: string
  readonly writer: StyleWriter
  width: number
  readonly space: boolean
  readonly tab: boolean
}
type TypesetLine = { readonly tokens: Token[]; readonly width: number; readonly spaces: number; readonly lastOfParagraph: boolean }

function tokenize(edit: TextEdit, writers: ReadonlyArray<StyleWriter>): Array<Token | 'break'> {
  const tokens: Array<Token | 'break'> = []
  for (const span of edit.spans) {
    const writer = writers[span.style]
    for (const part of span.text.split(/(\r\n|\r|\n| +|\t)/)) {
      if (part.length === 0) continue
      if (/^(\r\n|\r|\n)$/.test(part)) {
        tokens.push('break')
        continue
      }
      if (part === '\t') {
        // Its width depends on where the line has got to; set while typesetting.
        tokens.push({ text: part, writer, width: 0, space: false, tab: true })
        continue
      }
      const space = /^ +$/.test(part)
      tokens.push({ text: part, writer, width: writer.width(part), space, tab: false })
    }
  }
  return tokens
}

function typeset(edit: TextEdit, writers: ReadonlyArray<StyleWriter>): TypesetLine[] {
  const lines: TypesetLine[] = []
  let current: Token[] = []
  /** Where the line being filled starts, from the paragraph's left edge. */
  const indentOf = (lineIndex: number) => (lineIndex === 0 ? edit.firstLineIndent : 0)
  const finish = (lastOfParagraph: boolean) => {
    while (current.length > 0 && current[current.length - 1].space) current.pop()
    while (current.length > 0 && current[0].space && lines.length > 0 && !lines[lines.length - 1].lastOfParagraph) current.shift()
    const width = current.reduce((sum, token) => sum + token.width, 0)
    const spaces = current.filter((token) => token.space).length
    lines.push({ tokens: current, width, spaces, lastOfParagraph })
    current = []
  }
  const limit = edit.width
  for (const token of tokenize(edit, writers)) {
    if (token === 'break') {
      finish(true)
      continue
    }
    const width = current.reduce((sum, t) => sum + t.width, 0)
    const lineStart = indentOf(lines.length)
    if (token.tab) {
      // Advance to the next stop past the text so far, or by half an em.
      const stop = edit.tabStops.find((candidate) => candidate > lineStart + width + 0.01)
      token.width = stop !== undefined ? stop - lineStart - width : token.writer.style.size * 0.5
      current.push(token)
      continue
    }
    // A word that doesn't fit starts a new line — unless it's alone on its line.
    if (limit !== null && !token.space && lineStart + width + token.width > limit + 0.01 && current.some((t) => !t.space)) {
      finish(false)
    }
    if (token.space && current.length === 0 && lines.length > 0 && !lines[lines.length - 1].lastOfParagraph) continue
    current.push(token)
  }
  finish(true)
  return lines
}

// ---- Emitting -------------------------------------------------------------------------

function fillOps(fill: FillColor): string {
  return [fill.space, fill.color].filter((part): part is string => part !== null).join(' ') || '0 g'
}

function matrixOps(m: Matrix): string {
  return m.map(formatNumber).join(' ')
}

/** The content of a new text block for `edit`, assuming the CTM `ctm`. */
function emitEdit(edit: TextEdit, writers: ReadonlyArray<StyleWriter>, ctm: Matrix): string {
  const lines = typeset(edit, writers)
  const out: string[] = ['BT', '0 Ts']
  const normal = { x: -edit.direction.y, y: edit.direction.x }
  let active: { writer: StyleWriter | null; fill: string | null } = { writer: null, fill: null }

  lines.forEach((line, index) => {
    if (line.tokens.length === 0) return
    const indent = index === 0 ? edit.firstLineIndent : 0
    let offset = indent
    let extraPerSpace = 0
    const room = edit.width === null ? 0 : edit.width - indent - line.width
    if (edit.align === 'right') offset += room
    else if (edit.align === 'center') offset += room / 2
    else if (edit.align === 'justify' && !line.lastOfParagraph && line.spaces > 0 && room > 0) extraPerSpace = room / line.spaces
    const origin = {
      x: edit.origin.x + edit.direction.x * offset - normal.x * edit.leading * index,
      y: edit.origin.y + edit.direction.y * offset - normal.y * edit.leading * index,
    }
    const tm = textMatrixFor(rotationMatrix(edit.direction, origin), ctm)
    if (!tm) return
    out.push(`${matrixOps(tm)} Tm`)

    let items: string[] = []
    const flushItems = () => {
      if (items.length > 0) out.push(`[${items.join(' ')}] TJ`)
      items = []
    }
    for (const token of line.tokens) {
      const { writer } = token
      const fill = fillOps(writer.style.fill)
      if (active.writer !== writer) {
        flushItems()
        if (active.writer?.resource !== writer.resource || active.writer.style.size !== writer.style.size) {
          out.push(`${formatName(writer.resource)} ${formatNumber(writer.style.size)} Tf`)
        }
        out.push(
          `${formatNumber(writer.style.charSpacing)} Tc ${formatNumber(writer.wordSpacing)} Tw ${formatNumber(writer.style.horizontalScale * 100)} Tz ${writer.style.renderMode} Tr`,
        )
        active = { writer, fill: active.fill }
      }
      if (active.fill !== fill) {
        flushItems()
        out.push(fill)
        active = { ...active, fill }
      }
      if (token.tab) {
        items.push(formatNumber((-token.width * 1000) / (writer.style.size * writer.style.horizontalScale)))
        continue
      }
      items.push(writer.encode(token.text))
      if (token.space && extraPerSpace > 0) {
        items.push(formatNumber((-extraPerSpace * 1000) / (writer.style.size * writer.style.horizontalScale)))
      }
    }
    flushItems()
  })
  out.push('ET')
  return out.join('\n')
}

/** Puts the text state and fill colour back to what the page had, after an inserted block. */
function restoreOps(state: { readonly text: TextState; readonly fill: FillColor }): string {
  const t = state.text
  const ops: string[] = []
  if (t.fontResource) ops.push(`${formatName(t.fontResource)} ${formatNumber(t.fontSize)} Tf`)
  ops.push(
    `${formatNumber(t.charSpacing)} Tc ${formatNumber(t.wordSpacing)} Tw ${formatNumber(t.horizontalScale * 100)} Tz ${formatNumber(t.leading)} TL ${formatNumber(t.rise)} Ts ${t.renderMode} Tr`,
  )
  ops.push(fillOps(state.fill))
  return ops.join('\n')
}

/** A text-showing operator with some of its strings turned into equal-width positioning. */
function rewriteOp(stream: StreamInfo, opIndex: number, removed: ReadonlyMap<number, TextPiece>): string {
  const op = stream.ops[opIndex]
  const toNumber = (piece: TextPiece) => (-piece.advance * 1000) / (piece.state.fontSize * piece.state.horizontalScale)
  switch (op.operator) {
    case 'Tj':
    case "'":
    case '"': {
      const piece = removed.get(0)
      if (!piece) return stream.src.slice(op.start, op.end)
      const prefix = op.operator === "'" ? 'T* ' : op.operator === '"' ? `${formatToken(op.operands[0])} Tw ${formatToken(op.operands[1])} Tc T* ` : ''
      return `${prefix}[${formatNumber(toNumber(piece))}] TJ`
    }
    case 'TJ': {
      const array = op.operands[0]
      if (array?.type !== 'array') return stream.src.slice(op.start, op.end)
      const items: Array<string | number> = []
      array.items.forEach((item, element) => {
        const piece = removed.get(element)
        const value = piece ? toNumber(piece) : item.type === 'number' ? item.value : null
        if (value !== null) {
          // Merge neighbouring adjustments into one number.
          const last = items[items.length - 1]
          if (typeof last === 'number') items[items.length - 1] = last + value
          else items.push(value)
        } else {
          items.push(formatToken(item))
        }
      })
      return `[${items.map((item) => (typeof item === 'number' ? formatNumber(item) : item)).join(' ')}] TJ`
    }
    default:
      return stream.src.slice(op.start, op.end)
  }
}

// ---- Resources ------------------------------------------------------------------------

let resourceCounter = 0

/**
 * Adds `font` to a resources dictionary under a fresh name, once per
 * dictionary. Keyed by the dictionary too: a page and a form XObject on it
 * have different resources, and a name registered in one means nothing in
 * the other.
 */
function fontResourceName(
  lib: PdfLib,
  context: PDFContext,
  resources: PDFDict,
  font: { readonly ref: PDFRef },
  names: WeakMap<PDFDict, Map<string, string>>,
): string {
  let forDict = names.get(resources)
  if (!forDict) {
    forDict = new Map()
    names.set(resources, forDict)
  }
  const existing = forDict.get(font.ref.toString())
  if (existing) return existing
  let fonts = asDict(lib, context, resources.get(lib.PDFName.of('Font')))
  if (!fonts) {
    fonts = context.obj({})
    resources.set(lib.PDFName.of('Font'), fonts)
  }
  let name: string
  do {
    resourceCounter += 1
    name = `AtlasF${resourceCounter}`
  } while (fonts.get(lib.PDFName.of(name)) !== undefined)
  fonts.set(lib.PDFName.of(name), font.ref)
  forDict.set(font.ref.toString(), name)
  return name
}

function fontInfoFor(lib: PdfLib, context: PDFContext, resources: PDFDict | undefined, name: string | null): FontInfo | null {
  if (!name || !resources) return null
  const fonts = asDict(lib, context, resources.get(lib.PDFName.of('Font')))
  const raw = fonts?.get(lib.PDFName.of(name))
  const dict = asDict(lib, context, raw)
  return dict ? readFontInfo(lib, context, dict, raw instanceof lib.PDFRef ? raw.toString() : name) : null
}

async function writersFor(
  lib: PdfLib,
  doc: PDFDocument,
  edit: TextEdit,
  resources: PDFDict,
  fallback: FallbackFontSource,
  resourceNames: WeakMap<PDFDict, Map<string, string>>,
): Promise<StyleWriter[]> {
  const context = doc.context
  return Promise.all(
    edit.styles.map(async (style, index) => {
      const text = edit.spans.filter((span) => span.style === index).map((span) => span.text.replace(/[\r\n]/g, '')).join('')
      const own = fontInfoFor(lib, context, resources, style.fontResource)
      if (own && own.editable && own.key === style.fontKey && own.encode(text.replace(/\s/g, '')) !== null) {
        return originalWriter(own, style.fontResource as string, style)
      }
      const font = await coveringFallback(lib, doc, fallback, style.font, text)
      return fallbackWriter(font, fontResourceName(lib, context, resources, font, resourceNames), style)
    }),
  )
}

// ---- Streams -----------------------------------------------------------------------------

function pageResources(lib: PdfLib, context: PDFContext, page: PDFPage): PDFDict {
  const inherited = resolve(lib, context, page.node.getInheritableAttribute(lib.PDFName.of('Resources')))
  if (inherited instanceof lib.PDFDict) {
    if (page.node.get(lib.PDFName.of('Resources')) === undefined) page.node.set(lib.PDFName.of('Resources'), inherited)
    return inherited
  }
  const created = context.obj({})
  page.node.set(lib.PDFName.of('Resources'), created)
  return created
}

/** Replaces a form XObject's content, keeping its dictionary (bounding box, resources, matrix…). */
function replaceFormStream(lib: PdfLib, context: PDFContext, key: string, src: string): void {
  const ref = Array.from(context.enumerateIndirectObjects()).find(([candidate]) => candidate.toString() === key)?.[0]
  const existing = ref ? context.lookup(ref) : undefined
  if (!ref || !(existing instanceof lib.PDFStream)) throw new StaleTextEditError()
  const entries: Record<string, PDFObject> = {}
  for (const [name, value] of existing.dict.entries()) {
    const text = name.decodeText()
    if (text !== 'Length' && text !== 'Filter' && text !== 'DecodeParms') entries[text] = value
  }
  context.assign(ref as PDFRef, context.flateStream(binaryStringToBytes(src), entries))
}

function rewriteStream(
  stream: StreamInfo,
  pieces: ReadonlyArray<TextPiece>,
  edits: ReadonlyArray<{ readonly edit: TextEdit; readonly block: string }>,
): string {
  const removedByOp = new Map<number, Map<number, TextPiece>>()
  const pieceByKey = new Map(pieces.map((piece) => [`${piece.opIndex}:${piece.element}`, piece]))
  const insertAfter = new Map<number, string[]>()
  for (const { edit, block } of edits) {
    for (const ref of edit.pieces) {
      const piece = pieceByKey.get(`${ref.opIndex}:${ref.element}`)
      if (!piece) throw new StaleTextEditError()
      const removed = removedByOp.get(ref.opIndex) ?? new Map<number, TextPiece>()
      removed.set(ref.element, piece)
      removedByOp.set(ref.opIndex, removed)
    }
    const blocks = insertAfter.get(edit.insertAfterOp) ?? []
    blocks.push(block)
    insertAfter.set(edit.insertAfterOp, blocks)
  }

  let out = ''
  let cursor = 0
  stream.ops.forEach((op, index) => {
    const removed = removedByOp.get(index)
    const inserted = insertAfter.get(index)
    if (!removed && !inserted) return
    out += stream.src.slice(cursor, op.start)
    out += removed ? rewriteOp(stream, index, removed) : stream.src.slice(op.start, op.end)
    cursor = op.end
    if (inserted) {
      const state = stream.stateAtBlockEnd.get(index)
      if (!state) throw new StaleTextEditError()
      out += `\n${inserted.join('\n')}\n${restoreOps(state)}\n`
    }
  })
  return out + stream.src.slice(cursor)
}

/**
 * Applies every text edit for one page. Edits for a stream that changed
 * since they were made (the file isn't the one they were made on) are
 * refused with `StaleTextEditError` rather than guessed at.
 */
export async function applyTextEditsToPage(
  lib: PdfLib,
  doc: PDFDocument,
  page: PDFPage,
  edits: ReadonlyArray<TextEdit>,
  fallback: FallbackFontSource,
): Promise<void> {
  if (edits.length === 0) return
  const context = doc.context
  const resources = pageResources(lib, context, page)
  const pageText: PageText | null = interpretPage(lib, context, page.node, resources)
  if (!pageText) throw new StaleTextEditError()
  const resourceNames = new WeakMap<PDFDict, Map<string, string>>()

  const byStream = new Map<string, Array<{ edit: TextEdit; block: string }>>()
  const appended: string[] = []
  for (const edit of edits) {
    const stream = pageText.streams.get(edit.streamKey)
    if (!stream) throw new StaleTextEditError()
    const streamResources = stream.resources ?? resources
    const writers = await writersFor(lib, doc, edit, streamResources, fallback, resourceNames)
    if (edit.insertAfterOp < 0) {
      appended.push(emitEdit(edit, writers, [1, 0, 0, 1, 0, 0]))
      continue
    }
    const state = stream.stateAtBlockEnd.get(edit.insertAfterOp)
    if (!state) throw new StaleTextEditError()
    const list = byStream.get(edit.streamKey) ?? []
    list.push({ edit, block: emitEdit(edit, writers, state.ctm) })
    byStream.set(edit.streamKey, list)
  }

  let pageSrc = pageText.streams.get('page')?.src ?? ''
  for (const [key, list] of byStream) {
    const stream = pageText.streams.get(key) as StreamInfo
    const src = rewriteStream(stream, pageText.pieces.filter((piece) => piece.streamKey === key), list)
    if (key === 'page') pageSrc = src
    else replaceFormStream(lib, context, key, src)
  }

  // New text assumes the default coordinate system: close our own `q` and any
  // the page itself left open, so the page's last `cm` can't move it.
  const openStates = 1 + Math.max(0, pageText.streams.get('page')?.endDepth ?? 0)
  const finalSrc = appended.length > 0 ? `q\n${pageSrc}\n${'Q\n'.repeat(openStates)}${appended.join('\n')}\n` : pageSrc
  if (byStream.has('page') || appended.length > 0) {
    const ref = context.register(context.flateStream(binaryStringToBytes(finalSrc)))
    page.node.set(lib.PDFName.of('Contents'), ref)
  }
}
