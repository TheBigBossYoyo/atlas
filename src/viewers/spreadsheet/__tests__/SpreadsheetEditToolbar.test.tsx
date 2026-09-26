/**
 * Shared editing toolbar for SpreadsheetViewer and CsvViewer (wave 3).
 *
 * Written to close a coverage gap found by an audit of the tree: at 35%
 * statements / 35% branches it was the least-covered file under
 * `src/viewers/`, and the uncovered half was all behaviour a user can reach
 * with one click — every row/column action, the clipboard paste, and the Save
 * As format picker that only exists for multi-format workbooks.
 *
 * The component is pure plumbing (every action is a callback prop), so these
 * assert the plumbing: which callback fires, with which argument, and when a
 * control is disabled instead of guessing a target.
 */
import { render, screen, waitFor } from '@testing-library/react'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { SpreadsheetEditToolbar, type SpreadsheetEditToolbarProps } from '../SpreadsheetEditToolbar'

const FORMATS = [
  { id: 'xlsx', label: 'Excel Workbook (.xlsx)' },
  { id: 'ods', label: 'OpenDocument Spreadsheet (.ods)' },
]

function handlers() {
  return {
    onUndo: vi.fn(),
    onRedo: vi.fn(),
    onInsertRowAbove: vi.fn(),
    onDeleteRow: vi.fn(),
    onInsertColumnLeft: vi.fn(),
    onDeleteColumn: vi.fn(),
    onPaste: vi.fn(),
    onSave: vi.fn(),
    onSaveAs: vi.fn(),
  }
}

function renderToolbar(overrides: Partial<SpreadsheetEditToolbarProps> = {}) {
  const spies = handlers()
  const props: SpreadsheetEditToolbarProps = {
    canUndo: true,
    canRedo: true,
    selection: { row: 3, col: 2 },
    saveFormats: FORMATS,
    ...spies,
    ...overrides,
  }
  render(<SpreadsheetEditToolbar {...props} />)
  return spies
}

function button(name: string): HTMLButtonElement {
  return screen.getByRole('button', { name }) as HTMLButtonElement
}

let readText: ReturnType<typeof vi.fn>

beforeEach(() => {
  readText = vi.fn().mockResolvedValue('a\tb\nc\td')
  Object.defineProperty(navigator, 'clipboard', {
    value: { readText },
    configurable: true,
    writable: true,
  })
})

afterEach(() => {
  vi.restoreAllMocks()
})

describe('SpreadsheetEditToolbar — undo/redo', () => {
  it('fires onUndo/onRedo and disables each when there is no history in that direction', () => {
    const spies = renderToolbar()

    button('Undo').click()
    button('Redo').click()

    expect(spies.onUndo).toHaveBeenCalledTimes(1)
    expect(spies.onRedo).toHaveBeenCalledTimes(1)
  })

  it('disables Undo when canUndo is false and Redo when canRedo is false', () => {
    renderToolbar({ canUndo: false, canRedo: false })

    expect(button('Undo').disabled).toBe(true)
    expect(button('Redo').disabled).toBe(true)
  })
})

describe('SpreadsheetEditToolbar — row and column actions', () => {
  it("passes the selected cell's row to the row actions", () => {
    const spies = renderToolbar({ selection: { row: 7, col: 4 } })

    button('Insert row above').click()
    button('Delete row').click()

    expect(spies.onInsertRowAbove).toHaveBeenCalledWith(7)
    expect(spies.onDeleteRow).toHaveBeenCalledWith(7)
  })

  it("passes the selected cell's column to the column actions", () => {
    const spies = renderToolbar({ selection: { row: 7, col: 4 } })

    button('Insert column left').click()
    button('Delete column').click()

    expect(spies.onInsertColumnLeft).toHaveBeenCalledWith(4)
    expect(spies.onDeleteColumn).toHaveBeenCalledWith(4)
  })

  it('disables every row/column action with nothing selected, rather than guessing a target', () => {
    const spies = renderToolbar({ selection: null })

    for (const name of ['Insert row above', 'Delete row', 'Insert column left', 'Delete column']) {
      expect(button(name).disabled).toBe(true)
      button(name).click()
    }

    expect(spies.onInsertRowAbove).not.toHaveBeenCalled()
    expect(spies.onDeleteRow).not.toHaveBeenCalled()
    expect(spies.onInsertColumnLeft).not.toHaveBeenCalled()
    expect(spies.onDeleteColumn).not.toHaveBeenCalled()
  })

  it('row 0 / column 0 still act (a falsy index is a real target, not "no selection")', () => {
    const spies = renderToolbar({ selection: { row: 0, col: 0 } })

    button('Insert row above').click()
    button('Insert column left').click()

    expect(spies.onInsertRowAbove).toHaveBeenCalledWith(0)
    expect(spies.onInsertColumnLeft).toHaveBeenCalledWith(0)
  })
})

describe('SpreadsheetEditToolbar — paste', () => {
  it('parses the clipboard as TSV and pastes at the selected cell', async () => {
    const spies = renderToolbar({ selection: { row: 5, col: 1 } })

    button('Paste').click()

    await waitFor(() => expect(spies.onPaste).toHaveBeenCalledTimes(1))
    expect(spies.onPaste).toHaveBeenCalledWith(5, 1, [
      ['a', 'b'],
      ['c', 'd'],
    ])
  })

  it('pastes at the top-left when nothing is selected (Paste is never disabled)', async () => {
    const spies = renderToolbar({ selection: null })

    expect(button('Paste').disabled).toBe(false)
    button('Paste').click()

    await waitFor(() => expect(spies.onPaste).toHaveBeenCalledWith(0, 0, expect.anything()))
  })

  it('treats an empty clipboard as nothing to paste, not as one empty cell', async () => {
    readText.mockResolvedValue('')
    const spies = renderToolbar()

    button('Paste').click()

    await waitFor(() => expect(readText).toHaveBeenCalled())
    expect(spies.onPaste).not.toHaveBeenCalled()
  })

  it('stays silent when the clipboard read is denied (the grid\'s own Ctrl+V still works)', async () => {
    readText.mockRejectedValue(new Error('NotAllowedError: Read permission denied.'))
    const spies = renderToolbar()

    button('Paste').click()

    await waitFor(() => expect(readText).toHaveBeenCalled())
    expect(spies.onPaste).not.toHaveBeenCalled()
  })
})

describe('SpreadsheetEditToolbar — save and Save As', () => {
  it('fires onSave', () => {
    const spies = renderToolbar()

    button('Save').click()

    expect(spies.onSave).toHaveBeenCalledTimes(1)
  })

  it('offers no format picker when the document has only one save format', () => {
    renderToolbar({ saveFormats: [{ id: 'csv', label: 'CSV' }] })

    expect(screen.queryByRole('combobox', { name: 'Save As format' })).toBeNull()
    expect(screen.queryByRole('button', { name: 'Save As' })).toBeNull()
  })

  it('defaults Save As to the first format', () => {
    const spies = renderToolbar()

    button('Save As').click()

    expect(spies.onSaveAs).toHaveBeenCalledWith('xlsx')
  })

  it('saves as whichever format the picker is switched to', async () => {
    const spies = renderToolbar()
    const picker = screen.getByRole('combobox', { name: 'Save As format' }) as HTMLSelectElement

    picker.value = 'ods'
    picker.dispatchEvent(new Event('change', { bubbles: true }))
    await waitFor(() => expect(picker.value).toBe('ods'))
    button('Save As').click()

    expect(spies.onSaveAs).toHaveBeenCalledWith('ods')
  })

  it('lists every offered format as an option', () => {
    renderToolbar()

    const options = screen.getAllByRole('option') as HTMLOptionElement[]
    expect(options.map((o) => o.value)).toEqual(['xlsx', 'ods'])
    expect(options.map((o) => o.textContent)).toEqual([
      'Excel Workbook (.xlsx)',
      'OpenDocument Spreadsheet (.ods)',
    ])
  })

  it('falls back to an empty format id when no formats are offered at all', () => {
    // Defensive: `saveFormats[0]?.id ?? ''`. The picker is hidden in this case,
    // so the only thing that must hold is that it renders without throwing.
    renderToolbar({ saveFormats: [] })

    expect(button('Save')).toBeInTheDocument()
  })
})
