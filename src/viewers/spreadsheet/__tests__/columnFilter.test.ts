import { describe, expect, it } from 'vitest'
import {
  BLANK_FILTER_KEY,
  distinctColumnValues,
  filterKeyFor,
  normalizeFilter,
  rowMatchesFilters,
  shiftFiltersForColumnChange,
  withColumnFilter,
} from '../columnFilter'
import { compareForSort } from '../cellOrder'

const ROWS: ReadonlyArray<ReadonlyArray<string>> = [
  ['Region', 'Units'],
  ['North', '10'],
  ['south', '2'],
  ['North', ''],
  ['  ', '10'],
  ['South', '9'],
]

describe('distinctColumnValues', () => {
  it('groups spellings that differ only by case, keeping the first one seen', () => {
    const options = distinctColumnValues(ROWS, 0, 1)
    expect(options.map((o) => o.label)).toEqual(['North', 'south', ''])
    const south = options.find((o) => o.key === filterKeyFor('SOUTH'))
    // Two rows, "south" and "South" — one entry, counted twice. Offering them
    // separately would contradict the sort, which treats them as equal.
    expect(south?.count).toBe(2)
  })

  it('treats a whitespace-only cell as blank, and puts blanks last', () => {
    const options = distinctColumnValues(ROWS, 0, 1)
    expect(options[options.length - 1]).toMatchObject({ key: BLANK_FILTER_KEY, blank: true, count: 1 })
    // Column 1's blank is a different row — the two columns are counted apart.
    expect(distinctColumnValues(ROWS, 1, 1).find((o) => o.blank)?.count).toBe(1)
  })

  it('orders numbers as numbers, not as text', () => {
    const rows = [['9'], ['10'], ['2']]
    expect(distinctColumnValues(rows, 0).map((o) => o.label)).toEqual(['2', '9', '10'])
  })

  it('agrees with the sort comparison, which is the point of sharing cellOrder', () => {
    const rows = [['pear'], [''], ['10'], ['Apple'], ['2']]
    const viaFilter = distinctColumnValues(rows, 0).map((o) => o.label)
    const viaSort = rows.map((r) => r[0]).sort((a, b) => compareForSort(a, b, 'asc'))
    expect(viaFilter).toEqual(viaSort)
  })

  it('skips the header rows it is told about', () => {
    expect(distinctColumnValues(ROWS, 0, 1).some((o) => o.label === 'Region')).toBe(false)
    expect(distinctColumnValues(ROWS, 0, 0).some((o) => o.label === 'Region')).toBe(true)
  })
})

describe('rowMatchesFilters', () => {
  const north = new Set([filterKeyFor('North')])

  it('admits every row when nothing is filtered', () => {
    expect(rowMatchesFilters(['anything'], new Map())).toBe(true)
  })

  it('matches case-insensitively, like the checklist groups', () => {
    expect(rowMatchesFilters(['NORTH', '1'], new Map([[0, north]]))).toBe(true)
    expect(rowMatchesFilters(['South', '1'], new Map([[0, north]]))).toBe(false)
  })

  it('ANDs the columns together', () => {
    const filters = new Map([
      [0, north],
      [1, new Set(['10'])],
    ])
    expect(rowMatchesFilters(['North', '10'], filters)).toBe(true)
    expect(rowMatchesFilters(['North', '2'], filters)).toBe(false)
  })

  it('matches a blank cell only when blanks are admitted', () => {
    const blanks = new Map([[0, new Set([BLANK_FILTER_KEY])]])
    expect(rowMatchesFilters(['   '], blanks)).toBe(true)
    expect(rowMatchesFilters([], blanks)).toBe(true)
    expect(rowMatchesFilters(['x'], blanks)).toBe(false)
  })
})

describe('normalizeFilter', () => {
  const options = distinctColumnValues(ROWS, 0, 1)

  it('clears the filter when everything is ticked', () => {
    expect(normalizeFilter(new Set(options.map((o) => o.key)), options)).toBeNull()
  })

  it('drops ticked values the column no longer has', () => {
    const stale = new Set([filterKeyFor('North'), filterKeyFor('East')])
    expect([...(normalizeFilter(stale, options) ?? [])]).toEqual([filterKeyFor('North')])
  })

  it('keeps an empty selection as an empty filter, not as "no filter"', () => {
    // Untick everything and nothing should show. Returning null would show all
    // rows instead, which reads as the filter silently refusing.
    expect(normalizeFilter(new Set(), options)?.size).toBe(0)
  })
})

describe('withColumnFilter', () => {
  it('sets and clears without mutating the map it was given', () => {
    const original: ReadonlyMap<number, ReadonlySet<string>> = new Map()
    const added = withColumnFilter(original, 2, new Set(['A']))
    expect(original.size).toBe(0)
    expect(added.get(2)?.has('A')).toBe(true)
    expect(withColumnFilter(added, 2, null).has(2)).toBe(false)
  })
})

describe('shiftFiltersForColumnChange', () => {
  const filters = new Map([
    [1, new Set(['a'])],
    [3, new Set(['b'])],
  ])

  it('moves filters right when a column is inserted at or before them', () => {
    const next = shiftFiltersForColumnChange(filters, 1, 'insert')
    expect([...next.keys()].sort()).toEqual([2, 4])
  })

  it('leaves filters on columns before the insertion alone', () => {
    const next = shiftFiltersForColumnChange(filters, 2, 'insert')
    expect([...next.keys()].sort()).toEqual([1, 4])
  })

  it('drops the filter on a deleted column and pulls the later ones left', () => {
    const next = shiftFiltersForColumnChange(filters, 1, 'delete')
    expect([...next.keys()]).toEqual([2])
    expect(next.get(2)?.has('b')).toBe(true)
  })
})
