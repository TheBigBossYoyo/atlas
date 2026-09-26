/**
 * Frozen-rows strip (the real `<input>`s that sit above the canvas grid).
 *
 * Written to close a coverage gap found by an audit of the tree — it sat at
 * 63% statements / 37% branches, with the commit-on-blur and Enter-to-commit
 * paths untested — and to pin the two behaviours that are easy to regress: a
 * formula cell must offer its FORMULA for editing rather than its computed
 * value (see `useSpreadsheetGrid.ts`'s header for why that one silently
 * destroys formulas), and each cell needs its own translated accessible name,
 * since a screen reader has nothing else to go on here.
 */
import { fireEvent, render, screen } from '@testing-library/react'
import { afterEach, describe, expect, it, vi } from 'vitest'

import { LocaleProvider } from '../../../i18n/LocaleProvider'
import { LOCALE_KEY } from '../../../i18n/locale'
import { FrozenRowsStrip, type FrozenRowsStripProps } from '../FrozenRowsStrip'

afterEach(() => {
  window.localStorage.clear()
})

function renderStrip(overrides: Partial<FrozenRowsStripProps> = {}, locale?: 'en' | 'fr') {
  const onCommit = vi.fn()
  const props: FrozenRowsStripProps = {
    rows: [['Name', 'Total']],
    formulas: [[undefined, undefined]],
    columnWidthsPx: [90, 110],
    rowMarkerWidthPx: 40,
    translateXPx: 0,
    onCommit,
    ...overrides,
  }
  // LocaleProvider reads the stored preference on mount, so the locale is set
  // the same way a returning user's would already be.
  if (locale !== undefined) window.localStorage.setItem(LOCALE_KEY, locale)
  render(
    <LocaleProvider>
      <FrozenRowsStrip {...props} />
    </LocaleProvider>,
  )
  return onCommit
}

describe('FrozenRowsStrip', () => {
  it('renders nothing at all when no rows are frozen', () => {
    const { container } = render(
      <FrozenRowsStrip
        rows={[]}
        formulas={[]}
        columnWidthsPx={[]}
        rowMarkerWidthPx={40}
        translateXPx={0}
        onCommit={vi.fn()}
      />,
    )

    expect(container).toBeEmptyDOMElement()
  })

  it('renders one editable cell per frozen cell, showing its text', () => {
    renderStrip()

    expect((screen.getByLabelText('Frozen row 1, column 1') as HTMLInputElement).value).toBe('Name')
    expect((screen.getByLabelText('Frozen row 1, column 2') as HTMLInputElement).value).toBe('Total')
  })

  it('offers a formula cell its formula, not its computed value', () => {
    // Committing the computed value back would silently replace the formula
    // with a frozen number — the bug this shape exists to prevent.
    renderStrip({ rows: [['42']], formulas: [['SUM(B1:B9)']], columnWidthsPx: [90] })

    expect((screen.getByLabelText('Frozen row 1, column 1') as HTMLInputElement).value).toBe('=SUM(B1:B9)')
  })

  it('commits an edit on blur, with the cell coordinates', () => {
    const onCommit = renderStrip({ rows: [['a', 'b'], ['c', 'd']], formulas: [[undefined, undefined], [undefined, undefined]] })
    const cell = screen.getByLabelText('Frozen row 2, column 2')

    fireEvent.change(cell, { target: { value: 'edited' } })
    fireEvent.blur(cell)

    expect(onCommit).toHaveBeenCalledWith(1, 1, 'edited')
  })

  it('Enter commits by blurring the cell', () => {
    const onCommit = renderStrip()
    const cell = screen.getByLabelText('Frozen row 1, column 1')

    // Focus first: Enter commits by calling `blur()`, and `blur()` only emits a
    // blur event for the element that actually has focus — which is the state a
    // user pressing Enter is always in.
    cell.focus()
    fireEvent.change(cell, { target: { value: 'typed then Enter' } })
    fireEvent.keyDown(cell, { key: 'Enter' })

    expect(onCommit).toHaveBeenCalledWith(0, 0, 'typed then Enter')
  })

  it('leaves other keys alone, so typing does not commit on every keystroke', () => {
    const onCommit = renderStrip()
    const cell = screen.getByLabelText('Frozen row 1, column 1')

    fireEvent.keyDown(cell, { key: 'a' })
    fireEvent.keyDown(cell, { key: 'Escape' })

    expect(onCommit).not.toHaveBeenCalled()
  })

  it('mirrors the grid\'s horizontal scroll and per-column widths', () => {
    renderStrip({ translateXPx: 120 })

    const track = document.querySelector('.spreadsheet-viewer__frozen-rows-track') as HTMLElement
    expect(track.style.transform).toBe('translateX(-120px)')
    expect((screen.getByLabelText('Frozen row 1, column 1') as HTMLInputElement).style.width).toBe('90px')
  })

  it('falls back to a default width for a column with no measured width', () => {
    renderStrip({ columnWidthsPx: [] })

    expect((screen.getByLabelText('Frozen row 1, column 1') as HTMLInputElement).style.width).toBe('120px')
  })

  it('names each cell in the active locale (I18N-1)', () => {
    renderStrip({}, 'fr')

    expect(screen.getByLabelText('Ligne figée 1, colonne 1')).toBeInTheDocument()
    expect(screen.getByLabelText('Ligne figée 1, colonne 2')).toBeInTheDocument()
  })
})
