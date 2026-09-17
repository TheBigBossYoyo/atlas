/**
 * D23 — per-paragraph line-box cache.
 *
 * Every edit re-paginates the whole document, and line breaking (itemizing
 * runs, measuring glyph widths, building line boxes) is where nearly all of
 * that time goes: on a 35-page document a single keystroke cost seconds.
 *
 * The document model is immutable, so an edit rebuilds only the paragraph it
 * touched — every other `Paragraph` object keeps its identity. That makes
 * object identity a sound cache key: if the same paragraph object is laid out
 * again at the same width, with the same styles, fonts and theme, its lines
 * cannot have changed.
 *
 * What is deliberately NOT cached (the caller decides, see
 * `paragraphLineCacheKey`): paragraphs whose layout depends on state built up
 * from OTHER paragraphs — list markers (the counter advances as earlier items
 * are added or removed) and footnote/endnote references (their marks are
 * numbered across the document). Those are re-laid-out every time.
 *
 * The cache is keyed weakly on the paragraph, so entries disappear with the
 * document itself; the small per-paragraph map bounds how many widths (zoom
 * levels, column layouts) are remembered.
 */
import type { LineBox } from './types'
import type { Paragraph } from '../model/document'

/** Distinct widths/style-generations remembered per paragraph (zoom in and out, then back). */
const MAX_VARIANTS_PER_PARAGRAPH = 4

const cache = new WeakMap<Paragraph, Map<string, ReadonlyArray<LineBox>>>()

const identities = new WeakMap<object, number>()
let nextIdentity = 1

/** A stable id for an object reference, so unrelated objects never share a cache key. */
function identityOf(value: object | undefined | null): string {
  if (value === null || value === undefined) return '-'
  const existing = identities.get(value)
  if (existing !== undefined) return String(existing)
  identities.set(value, nextIdentity)
  return String(nextIdentity++)
}

/**
 * Builds the cache key for a paragraph's lines, or `null` when this paragraph
 * must not be cached (its layout depends on the document around it).
 */
export function paragraphLineCacheKey(parts: {
  readonly widthPt: number
  readonly styles: object | undefined
  readonly numbering: object | undefined
  readonly fontResolver: object
  readonly theme: object | undefined
  /** True when a list marker or a note mark is prepended to this paragraph. */
  readonly dependsOnDocumentState: boolean
}): string | null {
  if (parts.dependsOnDocumentState) return null
  return [
    parts.widthPt.toFixed(2),
    identityOf(parts.styles),
    identityOf(parts.numbering),
    identityOf(parts.fontResolver),
    identityOf(parts.theme),
  ].join('|')
}

/** Returns the cached lines for `paragraph`/`key`, or runs `compute` and remembers the result. */
export async function cachedParagraphLines(
  paragraph: Paragraph,
  key: string | null,
  compute: () => Promise<ReadonlyArray<LineBox>>,
): Promise<ReadonlyArray<LineBox>> {
  if (key === null) return compute()

  const variants = cache.get(paragraph)
  const hit = variants?.get(key)
  if (hit !== undefined) return hit

  const lines = await compute()
  const target = variants ?? new Map<string, ReadonlyArray<LineBox>>()
  if (variants === undefined) cache.set(paragraph, target)
  if (target.size >= MAX_VARIANTS_PER_PARAGRAPH) {
    const oldest = target.keys().next()
    if (!oldest.done) target.delete(oldest.value)
  }
  target.set(key, lines)
  return lines
}
