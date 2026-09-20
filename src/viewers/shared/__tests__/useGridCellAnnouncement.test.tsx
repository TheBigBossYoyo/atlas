import { renderHook } from '@testing-library/react'
import { describe, expect, it } from 'vitest'
import type { GridSelection } from '@glideapps/glide-data-grid'
import { CompactSelection } from '@glideapps/glide-data-grid'

import { cellReference, useGridCellAnnouncement } from '../useGridCellAnnouncement'
import { LocaleProvider } from '../../../i18n'

const EMPTY_SETS = { columns: CompactSelection.empty(), rows: CompactSelection.empty() } as const

function selectionFor(col: number, row: number): GridSelection {
  return {
    current: { cell: [col, row], range: { x: col, y: row, width: 1, height: 1 }, rangeStack: [] },
    ...EMPTY_SETS,
  }
}

const rows: ReadonlyArray<ReadonlyArray<string>> = [
  ['Name', 'Total'],
  ['Widgets', '42'],
  ['Gadgets', ''],
]

function renderAnnouncement(gridSelection: GridSelection | undefined) {
  return renderHook(() => useGridCellAnnouncement(rows, gridSelection), {
    wrapper: ({ children }) => <LocaleProvider>{children}</LocaleProvider>,
  })
}

describe('cellReference', () => {
  it('converts 0-based column/row indices to spreadsheet-style A1 references', () => {
    expect(cellReference(0, 0)).toBe('A1')
    expect(cellReference(1, 2)).toBe('B3')
    expect(cellReference(25, 0)).toBe('Z1')
    expect(cellReference(26, 0)).toBe('AA1')
  })
})

describe('useGridCellAnnouncement', () => {
  it('is empty with no selection', () => {
    const { result } = renderAnnouncement(undefined)
    expect(result.current).toBe('')
  })

  it('announces the reference and value of a selected cell', () => {
    const { result } = renderAnnouncement(selectionFor(1, 1))
    expect(result.current).toBe('B2: 42')
  })

  it('announces a blank cell distinctly from a missing selection', () => {
    const { result } = renderAnnouncement(selectionFor(1, 2))
    expect(result.current).toBe('B3, blank')
  })
})
