/**
 * The editor's read-only view of the original file's text (PDF-EDIT-2): its
 * paragraphs per page, found once and cached, and whether new text for a
 * paragraph can be drawn at all.
 *
 * Always built from the bytes the document was opened with — the same bytes
 * every save starts from — so a paragraph found here is exactly what the
 * writer will find again.
 */

import type { PDFDocument } from 'pdf-lib'

import type { FallbackFontSource } from './rewriteText'
import type { PdfLib } from '../writePdf'
import type { FontStyle, TextEdit } from '../types'
import type { FontInfo } from './fontInfo'
import { formsSharedAcrossPages, interpretPage, type PageText } from './interpreter'
import { buildParagraphs, type TextParagraph } from './paragraphs'

type FontkitLike = { create(bytes: Uint8Array): { hasGlyphForCodePoint(codePoint: number): boolean } }

export type TextAnalysis = {
  paragraphs(sourcePage: number): Promise<ReadonlyArray<TextParagraph>>
  /** The first character of `edit` no font can draw, or null when all can be. */
  undrawable(edit: TextEdit): Promise<string | null>
}

export function createTextAnalysis(
  lib: PdfLib,
  original: Uint8Array,
  loadFallback: () => Promise<FallbackFontSource>,
): TextAnalysis {
  let docPromise: Promise<PDFDocument> | null = null
  let sharedForms: Set<string> | null = null
  const pages = new Map<number, Promise<{ text: PageText | null; paragraphs: TextParagraph[] }>>()
  const coverage = new Map<string, Promise<(codePoint: number) => boolean>>()

  const doc = () => {
    docPromise ??= lib.PDFDocument.load(original, { updateMetadata: false })
    return docPromise
  }

  const page = (sourcePage: number) => {
    let pending = pages.get(sourcePage)
    if (!pending) {
      pending = doc().then((loaded) => {
        sharedForms ??= formsSharedAcrossPages(lib, loaded.context, loaded.getPages().map((p) => p.node.Resources()))
        const pdfPage = loaded.getPage(sourcePage - 1)
        const text = interpretPage(lib, loaded.context, pdfPage.node, pdfPage.node.Resources())
        return { text, paragraphs: text ? buildParagraphs(text, sharedForms) : [] }
      })
      pages.set(sourcePage, pending)
    }
    return pending
  }

  const coverageOf = (style: FontStyle) => {
    const key = `${style.family}|${style.bold}|${style.italic}`
    let pending = coverage.get(key)
    if (!pending) {
      pending = loadFallback().then(async (source) => {
        const font = (source.fontkit as FontkitLike).create(await source.load(style))
        return (codePoint: number) => font.hasGlyphForCodePoint(codePoint)
      })
      coverage.set(key, pending)
    }
    return pending
  }

  /** Whether some bundled family (the writer tries them all — see `coveringFallback`) can draw a character. */
  const fallbackCovers = async (edit: TextEdit, styleIndex: number) => {
    const style = edit.styles[styleIndex].font
    const families: Array<FontStyle['family']> = ['helvetica', 'times', 'courier']
    const checks = await Promise.all(families.map((family) => coverageOf({ ...style, family })))
    return (codePoint: number) => checks.some((covers) => covers(codePoint))
  }

  return {
    paragraphs: async (sourcePage) => (await page(sourcePage)).paragraphs,
    async undrawable(edit) {
      const { text } = await page(edit.page)
      const fonts = new Map<string, FontInfo>()
      for (const piece of text?.pieces ?? []) {
        if (piece.state.font) fonts.set(piece.state.font.key, piece.state.font)
      }
      for (const span of edit.spans) {
        const style = edit.styles[span.style]
        const own = style.fontKey ? fonts.get(style.fontKey) : undefined
        const covers = await fallbackCovers(edit, span.style)
        for (const ch of span.text) {
          if (/\s/.test(ch)) continue
          if (own?.encode(ch)) continue
          if (!covers(ch.codePointAt(0) ?? 0)) return ch
        }
      }
      return null
    },
  }
}
