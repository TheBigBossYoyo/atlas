import { describe, expect, it } from 'vitest'

import {
  countMatchesOnPage,
  extractPageText,
  findMatches,
  localMatchIndexOnPage,
} from '../search'

describe('extractPageText', () => {
  it('joins string `str` fields with a space', () => {
    expect(extractPageText([{ str: 'Hello' }, { str: 'World' }])).toBe('Hello World')
  })

  it('treats missing/non-string str as empty', () => {
    expect(extractPageText([{ str: 'A' }, {}, { str: 'B' }])).toBe('A  B')
  })

  it('returns an empty string for no items', () => {
    expect(extractPageText([])).toBe('')
  })
})

describe('findMatches', () => {
  const pages = [
    { pageNumber: 1, text: 'Atlas is a document viewer' },
    { pageNumber: 2, text: 'The Atlas viewer supports PDF and DOCX' },
    { pageNumber: 3, text: 'No matches on this page' },
  ]

  it('finds matches across every page, case-insensitively', () => {
    const matches = findMatches(pages, 'atlas')
    expect(matches).toHaveLength(2)
    expect(matches[0].pageNumber).toBe(1)
    expect(matches[1].pageNumber).toBe(2)
  })

  it('finds multiple occurrences on the same page', () => {
    const matches = findMatches(pages, 'viewer')
    expect(matches.filter((m) => m.pageNumber === 1)).toHaveLength(1)
    expect(matches.filter((m) => m.pageNumber === 2)).toHaveLength(1)
  })

  it('returns an empty array for a blank query', () => {
    expect(findMatches(pages, '')).toEqual([])
    expect(findMatches(pages, '   ')).toEqual([])
  })

  it('returns an empty array when nothing matches', () => {
    expect(findMatches(pages, 'nonexistent-term')).toEqual([])
  })

  it('reports the correct index and length', () => {
    const matches = findMatches([{ pageNumber: 1, text: 'find me here' }], 'me')
    expect(matches).toEqual([{ pageNumber: 1, index: 5, length: 2 }])
  })
})

describe('localMatchIndexOnPage', () => {
  const matches = [
    { pageNumber: 1, index: 0, length: 3 },
    { pageNumber: 1, index: 10, length: 3 },
    { pageNumber: 2, index: 0, length: 3 },
  ]

  it('returns the 0-based ordinal of a match within its own page', () => {
    expect(localMatchIndexOnPage(matches, 0)).toBe(0)
    expect(localMatchIndexOnPage(matches, 1)).toBe(1)
    expect(localMatchIndexOnPage(matches, 2)).toBe(0)
  })

  it('returns -1 for an out-of-range index', () => {
    expect(localMatchIndexOnPage(matches, -1)).toBe(-1)
    expect(localMatchIndexOnPage(matches, 99)).toBe(-1)
  })
})

describe('countMatchesOnPage', () => {
  it('counts only matches for the given page', () => {
    const matches = [
      { pageNumber: 1, index: 0, length: 1 },
      { pageNumber: 1, index: 5, length: 1 },
      { pageNumber: 2, index: 0, length: 1 },
    ]
    expect(countMatchesOnPage(matches, 1)).toBe(2)
    expect(countMatchesOnPage(matches, 2)).toBe(1)
    expect(countMatchesOnPage(matches, 3)).toBe(0)
  })
})
