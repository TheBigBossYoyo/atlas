import { fireEvent, render, screen } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'

import { FindReplace } from '../FindReplace'
import type { FindReplaceProps } from '../FindReplace'

function makeProps(overrides: Partial<FindReplaceProps> = {}): FindReplaceProps {
  return {
    open: true,
    onClose: vi.fn(),
    onFind: vi.fn(),
    onFindNext: vi.fn(),
    onFindPrev: vi.fn(),
    onReplace: vi.fn(),
    onReplaceAll: vi.fn(),
    matchCount: 0,
    currentMatchIndex: null,
    ...overrides,
  }
}

describe('FindReplace component', () => {
  it('renders when open=true and hides when open=false', () => {
    const { container, rerender } = render(<FindReplace {...makeProps({ open: true })} />)
    const dialog = container.querySelector('.docx-find')
    expect(dialog).not.toBeNull()
    expect(dialog?.classList.contains('docx-find--closed')).toBe(false)

    rerender(<FindReplace {...makeProps({ open: false })} />)
    expect(dialog?.classList.contains('docx-find--closed')).toBe(true)
  })

  it('emits onFindNext when Enter is pressed in the find input', () => {
    const onFindNext = vi.fn()
    render(<FindReplace {...makeProps({ onFindNext })} />)
    const input = screen.getByLabelText('Find')
    fireEvent.keyDown(input, { key: 'Enter' })
    expect(onFindNext).toHaveBeenCalledTimes(1)
  })

  it('emits onClose when Escape is pressed in the find input', () => {
    const onClose = vi.fn()
    render(<FindReplace {...makeProps({ onClose })} />)
    const input = screen.getByLabelText('Find')
    fireEvent.keyDown(input, { key: 'Escape' })
    expect(onClose).toHaveBeenCalledTimes(1)
  })

  it('emits onReplaceAll when Replace All button is clicked', () => {
    const onReplaceAll = vi.fn()
    render(<FindReplace {...makeProps({ onReplaceAll })} />)
    const btn = screen.getByLabelText('Replace all')
    fireEvent.click(btn)
    expect(onReplaceAll).toHaveBeenCalledTimes(1)
  })

  it('shows "X of N" status when matchCount > 0 and currentMatchIndex is set', () => {
    render(<FindReplace {...makeProps({ matchCount: 5, currentMatchIndex: 2 })} />)
    expect(screen.getByText('3 of 5')).toBeTruthy()
  })

  it('shows "No matches" when matchCount is 0', () => {
    render(<FindReplace {...makeProps({ matchCount: 0, currentMatchIndex: null })} />)
    expect(screen.getByText('No matches')).toBeTruthy()
  })
})
