import { describe, expect, it, vi } from 'vitest'

import {
  buildFallbackNavItems,
  buildOutlineNavItems,
  isPdfRef,
  resolveDestinationPage,
} from '../outline'
import type { PdfOutlineNode, PdfRef } from '../types'

describe('isPdfRef', () => {
  it('returns true for an object with numeric num/gen', () => {
    expect(isPdfRef({ num: 3, gen: 0 })).toBe(true)
  })

  it('returns false for null, primitives, and shapes missing num/gen', () => {
    expect(isPdfRef(null)).toBe(false)
    expect(isPdfRef(42)).toBe(false)
    expect(isPdfRef('ref')).toBe(false)
    expect(isPdfRef({ num: 3 })).toBe(false)
    expect(isPdfRef({ gen: 0 })).toBe(false)
  })
})

describe('resolveDestinationPage', () => {
  const getPageIndex = vi.fn(async (ref: PdfRef) => ref.num - 1)
  const getDestination = vi.fn(async (id: string) =>
    id === 'chapter-2' ? [{ num: 5, gen: 0 }, 'Fit'] : null,
  )

  it('resolves an explicit destination array with a numeric page ref (0-based)', async () => {
    const page = await resolveDestinationPage(getDestination, getPageIndex, [3, 'Fit'])
    expect(page).toBe(4)
  })

  it('resolves an explicit destination array with an indirect page ref', async () => {
    const page = await resolveDestinationPage(getDestination, getPageIndex, [
      { num: 5, gen: 0 },
      'Fit',
    ])
    expect(page).toBe(5)
  })

  it('resolves a named destination string via getDestination', async () => {
    const page = await resolveDestinationPage(getDestination, getPageIndex, 'chapter-2')
    expect(page).toBe(5)
  })

  it('returns null for a named destination that does not resolve', async () => {
    const page = await resolveDestinationPage(getDestination, getPageIndex, 'missing')
    expect(page).toBeNull()
  })

  it('returns null for a null/empty destination', async () => {
    expect(await resolveDestinationPage(getDestination, getPageIndex, null)).toBeNull()
    expect(await resolveDestinationPage(getDestination, getPageIndex, [])).toBeNull()
  })

  it('returns null when the first element is neither a number nor a ref', async () => {
    expect(
      await resolveDestinationPage(getDestination, getPageIndex, [{ weird: true }, 'Fit']),
    ).toBeNull()
  })
})

describe('buildOutlineNavItems', () => {
  const getPageIndex = async (ref: PdfRef) => ref.num - 1
  const getDestination = async () => null

  it('builds one nav item per top-level outline node, in order', async () => {
    const outline: PdfOutlineNode[] = [
      { title: 'Intro', dest: [0, 'Fit'], items: [] },
      { title: 'Chapter 1', dest: [1, 'Fit'], items: [] },
    ]
    const scrollToPage = vi.fn()

    const items = await buildOutlineNavItems(outline, scrollToPage, getDestination, getPageIndex)

    expect(items.map((item) => item.label)).toEqual(['Intro', 'Chapter 1'])
    expect(items[0].id).toBe('page-1-Intro')
    expect(items[1].id).toBe('page-2-Chapter 1')
  })

  it('flattens nested outline items depth-first, incrementing level', async () => {
    const outline: PdfOutlineNode[] = [
      {
        title: 'Part I',
        dest: [0, 'Fit'],
        items: [{ title: 'Section 1.1', dest: [1, 'Fit'], items: [] }],
      },
    ]
    const scrollToPage = vi.fn()

    const items = await buildOutlineNavItems(outline, scrollToPage, getDestination, getPageIndex)

    expect(items).toHaveLength(2)
    expect(items[0]).toMatchObject({ label: 'Part I', level: 1 })
    expect(items[1]).toMatchObject({ label: 'Section 1.1', level: 2 })
  })

  it('falls back to "Untitled" for a blank title and calls scrollToPage on select', async () => {
    const outline: PdfOutlineNode[] = [{ title: '   ', dest: [4, 'Fit'], items: [] }]
    const scrollToPage = vi.fn()

    const items = await buildOutlineNavItems(outline, scrollToPage, getDestination, getPageIndex)

    expect(items[0].label).toBe('Untitled')
    items[0].onSelect()
    expect(scrollToPage).toHaveBeenCalledWith(5)
  })

  it('does not call scrollToPage when the destination could not be resolved', async () => {
    const outline: PdfOutlineNode[] = [{ title: 'Dead link', dest: null, items: [] }]
    const scrollToPage = vi.fn()

    const items = await buildOutlineNavItems(outline, scrollToPage, getDestination, getPageIndex)
    items[0].onSelect()

    expect(scrollToPage).not.toHaveBeenCalled()
  })

  it('treats a bookmark whose destination throws as unresolved instead of aborting the whole outline', async () => {
    // A single malformed/dangling destination (seen in real-world,
    // slightly-corrupted PDFs) must not bubble up and brick loading the
    // rest of the document's outline — or the document itself, since this
    // is awaited directly from the load effect.
    const throwingGetDestination = vi.fn(async () => {
      throw new Error('bad destination')
    })
    const outline: PdfOutlineNode[] = [
      { title: 'Broken bookmark', dest: 'nonexistent', items: [] },
      { title: 'Fine bookmark', dest: [0, 'Fit'], items: [] },
    ]
    const scrollToPage = vi.fn()

    const items = await buildOutlineNavItems(outline, scrollToPage, throwingGetDestination, getPageIndex)

    expect(items).toHaveLength(2)
    items[0].onSelect()
    expect(scrollToPage).not.toHaveBeenCalled()
    items[1].onSelect()
    expect(scrollToPage).toHaveBeenCalledWith(1)
  })
})

describe('buildFallbackNavItems', () => {
  it('builds one item per page, 1-indexed, wired to scrollToPage', () => {
    const scrollToPage = vi.fn()
    const items = buildFallbackNavItems(3, scrollToPage)

    expect(items).toHaveLength(3)
    expect(items.map((item) => item.label)).toEqual(['Page 1', 'Page 2', 'Page 3'])

    items[2].onSelect()
    expect(scrollToPage).toHaveBeenCalledWith(3)
  })

  it('returns an empty array for a zero-page document', () => {
    expect(buildFallbackNavItems(0, vi.fn())).toEqual([])
  })
})
