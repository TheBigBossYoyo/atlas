/**
 * Walks a page's drawing instructions and records every piece of text it
 * shows (PDF-EDIT-2): which operator showed it (so it can be rewritten), what
 * it says, where it sits on the page, and the full text and colour state it
 * was drawn with (so replacement text can match it).
 *
 * This tracks exactly the state PDF text depends on — the CTM through `q`/`Q`/
 * `cm`, the text matrices, font, size, character/word spacing, horizontal
 * scaling, rise, render mode and fill colour — and follows `Do` into form
 * XObjects, since headers, footers and whole pages are often drawn from one.
 */

import type { PDFContext, PDFDict, PDFObject } from 'pdf-lib'

import type { PdfLib } from '../writePdf'
import { readFontInfo, type CharCode, type FontInfo } from './fontInfo'
import { parseContentStream, type ContentOp, type ContentToken } from './lexer'
import { apply, IDENTITY, multiply, translate, yScale, type Matrix } from './matrix'
import { asDict, dictGet, nameOf, numbersOf, resolve, streamText } from './pdfObjects'

export type TextState = {
  readonly font: FontInfo | null
  /** The font's name in the stream's resources (`/F1`). */
  readonly fontResource: string | null
  readonly fontSize: number
  readonly charSpacing: number
  readonly wordSpacing: number
  /** `Tz` / 100. */
  readonly horizontalScale: number
  readonly leading: number
  readonly rise: number
  readonly renderMode: number
}

/** The fill colour as the operators that set it, copied verbatim (any colour space works). */
export type FillColor = {
  /** A `cs` operator, or null when the colour came from `g`/`rg`/`k` (or the default). */
  readonly space: string | null
  /** The colour operator, or null for the initial black. */
  readonly color: string | null
}

type GraphicsState = {
  readonly ctm: Matrix
  readonly text: TextState
  readonly fill: FillColor
}

export type TextPiece = {
  /** Which stream: `page` for the page's own content, else the form XObject's reference. */
  readonly streamKey: string
  readonly opIndex: number
  /** Index of the string in a `TJ` array; 0 for `Tj`, `'` and `"`. */
  readonly element: number
  readonly codes: ReadonlyArray<CharCode>
  readonly text: string
  readonly state: TextState
  readonly fill: FillColor
  readonly ctm: Matrix
  /** Text matrix before the first glyph. */
  readonly textMatrix: Matrix
  /** Baseline start and end in user space (rise included). */
  readonly start: { readonly x: number; readonly y: number }
  readonly end: { readonly x: number; readonly y: number }
  /** Font size in user space: what a reader sees, whatever mix of Tf, Tm and cm produced it. */
  readonly sizeUser: number
  /** Total horizontal displacement in text space, which a removal must keep. */
  readonly advance: number
  /** Index of the `BT` and `ET` operators around this piece. */
  readonly blockStart: number
  blockEnd: number
  readonly invisible: boolean
}

export type StreamInfo = {
  readonly key: string
  readonly src: string
  readonly ops: ReadonlyArray<ContentOp>
  /** Resources the stream's names resolve in. */
  readonly resources: PDFDict | undefined
  /** Graphics state at each `ET` operator, by op index — what an insertion after it must restore. */
  readonly stateAtBlockEnd: ReadonlyMap<number, { readonly text: TextState; readonly fill: FillColor; readonly ctm: Matrix }>
  /** The CTM in force when this stream starts (a form's Matrix × its caller's CTM). */
  readonly baseCtm: Matrix
  /** `q`s left open at the end of the stream (some producers don't balance them). */
  readonly endDepth: number
}

export type PageText = {
  readonly pieces: ReadonlyArray<TextPiece>
  readonly streams: ReadonlyMap<string, StreamInfo>
  /** How many times each form XObject was drawn on this page. */
  readonly formUses: ReadonlyMap<string, number>
}

const INITIAL_TEXT: TextState = {
  font: null,
  fontResource: null,
  fontSize: 0,
  charSpacing: 0,
  wordSpacing: 0,
  horizontalScale: 1,
  leading: 0,
  rise: 0,
  renderMode: 0,
}

const MAX_FORM_DEPTH = 8

function num(token: ContentToken | undefined): number {
  return token?.type === 'number' ? token.value : 0
}

type Context = {
  readonly lib: PdfLib
  readonly context: PDFContext
  readonly pieces: TextPiece[]
  readonly streams: Map<string, StreamInfo>
  readonly visitingForms: Set<string>
  readonly formUses: Map<string, number>
}

function fontFor(ctx: Context, resources: PDFDict | undefined, name: string): FontInfo | null {
  const { lib, context } = ctx
  const fonts = asDict(lib, context, resources?.get(lib.PDFName.of('Font')))
  const raw = fonts?.get(lib.PDFName.of(name))
  const dict = asDict(lib, context, raw)
  if (!dict) return null
  return readFontInfo(lib, context, dict, raw instanceof lib.PDFRef ? raw.toString() : `${name}@inline`)
}

function interpretStream(
  ctx: Context,
  key: string,
  src: string,
  resources: PDFDict | undefined,
  initial: GraphicsState,
  depth: number,
): void {
  const { lib, context } = ctx
  const ops = parseContentStream(src)
  const stateAtBlockEnd = new Map<number, { text: TextState; fill: FillColor; ctm: Matrix }>()
  const info: StreamInfo = { key, src, ops, resources, stateAtBlockEnd, baseCtm: initial.ctm, endDepth: 0 }
  ctx.streams.set(key, info)

  const stack: GraphicsState[] = []
  let gs: GraphicsState = initial
  let tm: Matrix = IDENTITY
  let tlm: Matrix = IDENTITY
  let blockStart = -1
  let blockPieces: TextPiece[] = []

  const setText = (patch: Partial<TextState>) => {
    gs = { ...gs, text: { ...gs.text, ...patch } }
  }
  const nextLine = () => {
    tlm = multiply(translate(0, -gs.text.leading), tlm)
    tm = tlm
  }

  const show = (opIndex: number, element: number, bytes: string) => {
    const { text: ts } = gs
    const font = ts.font
    const codes = font ? font.split(bytes) : Array.from(bytes, (ch) => ({ code: ch.charCodeAt(0), length: 1 }))
    const startTm = tm
    let advance = 0
    let text = ''
    for (const code of codes) {
      const w0 = font ? font.width(code.code) / 1000 : 0.5
      const isSpace = code.length === 1 && code.code === 32
      const tx = (w0 * ts.fontSize + ts.charSpacing + (isSpace ? ts.wordSpacing : 0)) * ts.horizontalScale
      advance += tx
      text += font ? font.unicode(code.code) : ''
    }
    tm = multiply(translate(advance, 0), tm)
    const startRendering = multiply(startTm, gs.ctm)
    const endRendering = multiply(tm, gs.ctm)
    const piece: TextPiece = {
      streamKey: key,
      opIndex,
      element,
      codes,
      text,
      state: ts,
      fill: gs.fill,
      ctm: gs.ctm,
      textMatrix: startTm,
      start: apply(startRendering, 0, ts.rise),
      end: apply(endRendering, 0, ts.rise),
      sizeUser: Math.abs(ts.fontSize) * yScale(startRendering),
      advance,
      blockStart,
      blockEnd: -1,
      invisible: ts.renderMode === 3 || ts.renderMode === 7,
    }
    ctx.pieces.push(piece)
    blockPieces.push(piece)
  }

  ops.forEach((op, index) => {
    const o = op.operands
    switch (op.operator) {
      case 'q':
        stack.push(gs)
        break
      case 'Q':
        gs = stack.pop() ?? gs
        break
      case 'cm':
        gs = { ...gs, ctm: multiply([num(o[0]), num(o[1]), num(o[2]), num(o[3]), num(o[4]), num(o[5])], gs.ctm) }
        break
      case 'g':
      case 'rg':
      case 'k':
        gs = { ...gs, fill: { space: null, color: src.slice(op.start, op.end) } }
        break
      case 'cs':
        gs = { ...gs, fill: { space: src.slice(op.start, op.end), color: null } }
        break
      case 'sc':
      case 'scn':
        gs = { ...gs, fill: { ...gs.fill, color: src.slice(op.start, op.end) } }
        break
      case 'BT':
        tm = IDENTITY
        tlm = IDENTITY
        blockStart = index
        blockPieces = []
        break
      case 'ET':
        for (const piece of blockPieces) piece.blockEnd = index
        stateAtBlockEnd.set(index, { text: gs.text, fill: gs.fill, ctm: gs.ctm })
        blockPieces = []
        break
      case 'Tf': {
        const name = o[0]?.type === 'name' ? o[0].value : null
        setText({ font: name ? fontFor(ctx, resources, name) : null, fontResource: name, fontSize: num(o[1]) })
        break
      }
      case 'Tc':
        setText({ charSpacing: num(o[0]) })
        break
      case 'Tw':
        setText({ wordSpacing: num(o[0]) })
        break
      case 'Tz':
        setText({ horizontalScale: num(o[0]) / 100 })
        break
      case 'TL':
        setText({ leading: num(o[0]) })
        break
      case 'Ts':
        setText({ rise: num(o[0]) })
        break
      case 'Tr':
        setText({ renderMode: num(o[0]) })
        break
      case 'Td':
        tlm = multiply(translate(num(o[0]), num(o[1])), tlm)
        tm = tlm
        break
      case 'TD':
        setText({ leading: -num(o[1]) })
        tlm = multiply(translate(num(o[0]), num(o[1])), tlm)
        tm = tlm
        break
      case 'Tm':
        tlm = [num(o[0]), num(o[1]), num(o[2]), num(o[3]), num(o[4]), num(o[5])]
        tm = tlm
        break
      case 'T*':
        nextLine()
        break
      case 'Tj':
        if (o[0]?.type === 'string') show(index, 0, o[0].bytes)
        break
      case "'":
        nextLine()
        if (o[0]?.type === 'string') show(index, 0, o[0].bytes)
        break
      case '"':
        setText({ wordSpacing: num(o[0]), charSpacing: num(o[1]) })
        nextLine()
        if (o[2]?.type === 'string') show(index, 0, o[2].bytes)
        break
      case 'TJ':
        if (o[0]?.type === 'array') {
          o[0].items.forEach((item, element) => {
            if (item.type === 'string') show(index, element, item.bytes)
            else if (item.type === 'number') {
              tm = multiply(translate((-item.value / 1000) * gs.text.fontSize * gs.text.horizontalScale, 0), tm)
            }
          })
        }
        break
      case 'Do': {
        if (depth >= MAX_FORM_DEPTH || o[0]?.type !== 'name') break
        const xobjects = asDict(lib, context, resources?.get(lib.PDFName.of('XObject')))
        const raw = xobjects?.get(lib.PDFName.of(o[0].value))
        const stream = resolve(lib, context, raw)
        if (!(stream instanceof lib.PDFStream) || !(raw instanceof lib.PDFRef)) break
        if (nameOf(lib, dictGet(lib, context, stream.dict, 'Subtype')) !== 'Form') break
        const formKey = raw.toString()
        ctx.formUses.set(formKey, (ctx.formUses.get(formKey) ?? 0) + 1)
        if (ctx.visitingForms.has(formKey) || ctx.streams.has(formKey)) break
        const formSrc = streamText(lib, stream)
        if (formSrc === null) break
        const matrix = numbersOf(lib, context, stream.dict.get(lib.PDFName.of('Matrix')))
        const formMatrix: Matrix = matrix && matrix.length === 6 ? (matrix as unknown as Matrix) : IDENTITY
        const formResources = asDict(lib, context, stream.dict.get(lib.PDFName.of('Resources'))) ?? resources
        ctx.visitingForms.add(formKey)
        interpretStream(ctx, formKey, formSrc, formResources, { ...gs, ctm: multiply(formMatrix, gs.ctm) }, depth + 1)
        ctx.visitingForms.delete(formKey)
        break
      }
      default:
        break
    }
  })
  ctx.streams.set(key, { ...info, endDepth: stack.length })
}

/** Concatenates a page's content streams the way a reader does (a newline between parts). */
export function pageContentSource(lib: PdfLib, context: PDFContext, page: PDFDict): string | null {
  const contents = resolve(lib, context, page.get(lib.PDFName.of('Contents')))
  const parts: PDFObject[] = contents instanceof lib.PDFArray ? contents.asArray() : contents ? [contents] : []
  const texts: string[] = []
  for (const part of parts) {
    const text = streamText(lib, resolve(lib, context, part))
    if (text === null) return null
    texts.push(text)
  }
  return texts.join('\n')
}

/** Every piece of text on a page, with the streams it came from. Null when the content can't be decoded. */
export function interpretPage(lib: PdfLib, context: PDFContext, page: PDFDict, resources: PDFDict | undefined): PageText | null {
  const src = pageContentSource(lib, context, page)
  if (src === null) return null
  const ctx: Context = { lib, context, pieces: [], streams: new Map(), visitingForms: new Set(), formUses: new Map() }
  interpretStream(ctx, 'page', src, resources, { ctm: IDENTITY, text: INITIAL_TEXT, fill: { space: null, color: null } }, 0)
  return { pieces: ctx.pieces, streams: ctx.streams, formUses: ctx.formUses }
}

/**
 * Form XObjects reachable from more than one page's resources (headers,
 * footers, stamped templates). Editing text inside one would change it on
 * every page that draws it, so such text is not offered for editing.
 */
export function formsSharedAcrossPages(
  lib: PdfLib,
  context: PDFContext,
  /** Each page's effective resources (inherited ones included). */
  pageResources: ReadonlyArray<PDFDict | undefined>,
): Set<string> {
  const pagesUsing = new Map<string, number>()
  pageResources.forEach((resources) => {
    const seen = new Set<string>()
    const visit = (resources: PDFDict | undefined, depth: number) => {
      if (!resources || depth > MAX_FORM_DEPTH) return
      const xobjects = asDict(lib, context, resources.get(lib.PDFName.of('XObject')))
      if (!xobjects) return
      for (const [, value] of xobjects.entries()) {
        if (!(value instanceof lib.PDFRef)) continue
        const key = value.toString()
        if (seen.has(key)) continue
        const stream = context.lookup(value)
        if (!(stream instanceof lib.PDFStream) || nameOf(lib, dictGet(lib, context, stream.dict, 'Subtype')) !== 'Form') continue
        seen.add(key)
        visit(asDict(lib, context, stream.dict.get(lib.PDFName.of('Resources'))), depth + 1)
      }
    }
    visit(resources, 0)
    seen.forEach((key) => pagesUsing.set(key, (pagesUsing.get(key) ?? 0) + 1))
  })
  return new Set([...pagesUsing].filter(([, count]) => count > 1).map(([key]) => key))
}
