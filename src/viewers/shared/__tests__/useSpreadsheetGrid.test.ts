/**
 * T5/DAT-12 (shared grid rendering) + wave 3 editing (DAT-06 closed for
 * real): `allowOverlay`/`onCellEdited` wiring.
 */
import { renderHook } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'
import type { EditableGridCell, Item } from '@glideapps/glide-data-grid'

import { useSpreadsheetGrid } from '../useSpreadsheetGrid'

const ROWS = [
  ['Name', 'Score'],
  ['Alice', '10'],
]

describe('useSpreadsheetGrid — read-only mode (no onCellEdited option)', () => {
  it('renders cells with allowOverlay:false and no onCellEdited handler', () => {
    const { result } = renderHook(() => useSpreadsheetGrid({ rows: ROWS, colCount: 2 }))
    expect(result.current.getCellContent([0, 0]).allowOverlay).toBe(false)
    expect(result.current.onCellEdited).toBeUndefined()
  })
})

describe('useSpreadsheetGrid — editable mode (onCellEdited option provided)', () => {
  it('renders cells with allowOverlay:true', () => {
    const { result } = renderHook(() =>
      useSpreadsheetGrid({ rows: ROWS, colCount: 2, onCellEdited: vi.fn() }),
    )
    expect(result.current.getCellContent([0, 0]).allowOverlay).toBe(true)
  })

  it('translates a committed text-cell edit into (row, col, rawText)', () => {
    const onCellEdited = vi.fn()
    const { result } = renderHook(() => useSpreadsheetGrid({ rows: ROWS, colCount: 2, onCellEdited }))

    const item: Item = [1, 0] // column 1 ("Score"), row 0
    const newValue = { kind: 'text', data: '99', displayData: '99', allowOverlay: true } as EditableGridCell
    result.current.onCellEdited?.(item, newValue)

    expect(onCellEdited).toHaveBeenCalledWith(0, 1, '99')
  })

  it('ignores a non-text committed value (defensive — this grid never opens a non-text editor)', () => {
    const onCellEdited = vi.fn()
    const { result } = renderHook(() => useSpreadsheetGrid({ rows: ROWS, colCount: 2, onCellEdited }))

    const item: Item = [0, 0]
    const newValue = { kind: 'boolean', data: true, allowOverlay: false } as EditableGridCell
    result.current.onCellEdited?.(item, newValue)

    expect(onCellEdited).not.toHaveBeenCalled()
  })
})

describe('useSpreadsheetGrid — formula cells (confirmed data-loss bug: re-editing a formula must not silently delete it)', () => {
  const ROWS_WITH_FORMULA = [
    ['1', '2', '3'],
    ['4', '5', '9'],
  ]
  const FORMULAS = [
    [undefined, undefined, undefined],
    [undefined, undefined, 'A2+B2'],
  ]

  it('seeds the edit overlay (`data`) with the literal formula text, while `displayData` stays the computed value', () => {
    const { result } = renderHook(() =>
      useSpreadsheetGrid({ rows: ROWS_WITH_FORMULA, colCount: 3, formulas: FORMULAS, onCellEdited: vi.fn() }),
    )

    const cell = result.current.getCellContent([2, 1]) as { data: string; displayData: string }
    expect(cell.data).toBe('=A2+B2')
    expect(cell.displayData).toBe('9')
  })

  it('leaves a plain-value cell (no formula) with `data` === `displayData`', () => {
    const { result } = renderHook(() =>
      useSpreadsheetGrid({ rows: ROWS_WITH_FORMULA, colCount: 3, formulas: FORMULAS, onCellEdited: vi.fn() }),
    )

    const cell = result.current.getCellContent([0, 0]) as { data: string; displayData: string }
    expect(cell.data).toBe('1')
    expect(cell.displayData).toBe('1')
  })

  it('re-committing a formula cell unchanged sends back the formula text, not the computed value (the bug this closes)', () => {
    const onCellEdited = vi.fn()
    const { result } = renderHook(() =>
      useSpreadsheetGrid({ rows: ROWS_WITH_FORMULA, colCount: 3, formulas: FORMULAS, onCellEdited }),
    )

    // Simulates glide-data-grid's own text-cell editor: it seeds its input
    // from `data` (see text-cell.js's `provideEditor`), so committing with
    // no changes at all sends `data` straight back as `newValue.data`.
    const seeded = result.current.getCellContent([2, 1]) as { data: string }
    const item: Item = [2, 1]
    const newValue = { kind: 'text', data: seeded.data, displayData: seeded.data, allowOverlay: true } as EditableGridCell
    result.current.onCellEdited?.(item, newValue)

    expect(onCellEdited).toHaveBeenCalledWith(1, 2, '=A2+B2')
  })
})
