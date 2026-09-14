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
