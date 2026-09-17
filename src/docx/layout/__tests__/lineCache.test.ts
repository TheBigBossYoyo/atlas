/** D23 — the per-paragraph line cache keys on paragraph identity and layout inputs. */
import { describe, expect, it, vi } from 'vitest'

import { cachedParagraphLines, paragraphLineCacheKey } from '../lineCache'
import type { Paragraph } from '../../model'
import type { LineBox } from '../types'

const lines = (label: string): ReadonlyArray<LineBox> => [{ label } as unknown as LineBox]

function makeParagraph(): Paragraph {
  return { kind: 'paragraph', props: {}, children: [] } as unknown as Paragraph
}

const styles = {}
const numbering = {}
const fontResolver = {}
const theme = {}

function key(widthPt: number, dependsOnDocumentState = false): string | null {
  return paragraphLineCacheKey({ widthPt, styles, numbering, fontResolver, theme, dependsOnDocumentState })
}

describe('paragraphLineCacheKey', () => {
  it('is stable for the same inputs and differs per width', () => {
    expect(key(468)).toBe(key(468))
    expect(key(468)).not.toBe(key(300))
  })

  it('refuses to cache a paragraph whose layout depends on the rest of the document', () => {
    expect(key(468, true)).toBeNull()
  })

  it('separates documents that merely look alike', () => {
    const other = paragraphLineCacheKey({
      widthPt: 468,
      styles: {},
      numbering,
      fontResolver,
      theme,
      dependsOnDocumentState: false,
    })
    expect(other).not.toBe(key(468))
  })
})

describe('cachedParagraphLines', () => {
  it('computes once per paragraph and key, then reuses', async () => {
    const paragraph = makeParagraph()
    const compute = vi.fn(async () => lines('a'))

    const first = await cachedParagraphLines(paragraph, key(468), compute)
    const second = await cachedParagraphLines(paragraph, key(468), compute)

    expect(compute).toHaveBeenCalledTimes(1)
    expect(second).toBe(first)
  })

  it('recomputes for a different paragraph, a different width, or no key at all', async () => {
    const paragraph = makeParagraph()
    const compute = vi.fn(async () => lines('a'))

    await cachedParagraphLines(paragraph, key(468), compute)
    await cachedParagraphLines(makeParagraph(), key(468), compute)
    await cachedParagraphLines(paragraph, key(300), compute)
    await cachedParagraphLines(paragraph, null, compute)
    await cachedParagraphLines(paragraph, null, compute)

    expect(compute).toHaveBeenCalledTimes(5)
  })

  it('remembers a bounded number of widths per paragraph', async () => {
    const paragraph = makeParagraph()
    const compute = vi.fn(async () => lines('a'))

    for (const width of [100, 200, 300, 400, 500]) {
      await cachedParagraphLines(paragraph, key(width), compute)
    }
    // The oldest entry was evicted, so laying it out again recomputes.
    await cachedParagraphLines(paragraph, key(100), compute)
    expect(compute).toHaveBeenCalledTimes(6)

    // The most recent ones are still cached.
    await cachedParagraphLines(paragraph, key(500), compute)
    expect(compute).toHaveBeenCalledTimes(6)
  })
})
