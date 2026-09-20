/** SHELL-17 — the open-documents strip. */
import { fireEvent, render, screen } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'

import { TabBar } from '../TabBar'
import type { DocumentSession } from '../../session/documentSessions'

const sessions: ReadonlyArray<DocumentSession> = [
  { id: '/a.md', name: 'a.md', file: { kind: 'text', content: '', path: '/a.md', format: 'markdown' } },
  { id: '/b.docx', name: 'b.docx', file: { kind: 'text', content: '', path: '/b.docx', format: 'markdown' } },
]

function renderBar(overrides: Partial<React.ComponentProps<typeof TabBar>> = {}) {
  const props = {
    sessions,
    activeId: '/a.md',
    isActiveDirty: false,
    onSelect: vi.fn(),
    onClose: vi.fn(),
    onReorder: vi.fn(),
    ...overrides,
  }
  render(<TabBar {...props} />)
  return props
}

describe('TabBar', () => {
  it('renders nothing when no document is open', () => {
    const { container } = render(
      <TabBar sessions={[]} activeId={null} isActiveDirty={false} onSelect={vi.fn()} onClose={vi.fn()} onReorder={vi.fn()} />,
    )
    expect(container).toBeEmptyDOMElement()
  })

  it('marks the showing document and selects another on click', () => {
    const props = renderBar()
    expect(screen.getByRole('tab', { name: /a\.md/ })).toHaveAttribute('aria-selected', 'true')
    expect(screen.getByRole('tab', { name: /b\.docx/ })).toHaveAttribute('aria-selected', 'false')

    fireEvent.click(screen.getByRole('tab', { name: /b\.docx/ }))
    expect(props.onSelect).toHaveBeenCalledWith('/b.docx')
  })

  it('shows the unsaved dot only on the document being edited', () => {
    renderBar({ isActiveDirty: true })
    const dots = screen.getAllByLabelText('Unsaved changes')
    expect(dots).toHaveLength(1)
    expect(screen.getByRole('tab', { name: /a\.md/ })).toContainElement(dots[0])
  })

  it('closes from the close button and from a middle click', () => {
    const props = renderBar()
    fireEvent.click(screen.getByRole('button', { name: 'Close b.docx' }))
    expect(props.onClose).toHaveBeenCalledWith('/b.docx')

    // Testing Library has no auxClick helper; dispatch the real event.
    const tab = screen.getByRole('tab', { name: /a[.]md/ }).parentElement!
    fireEvent(tab, new MouseEvent('auxclick', { button: 1, bubbles: true, cancelable: true }))
    expect(props.onClose).toHaveBeenCalledWith('/a.md')
  })

  it('reorders on drag and drop', () => {
    const props = renderBar()
    const first = screen.getByRole('tab', { name: /a\.md/ }).parentElement!
    const second = screen.getByRole('tab', { name: /b\.docx/ }).parentElement!

    fireEvent.dragStart(first)
    fireEvent.dragOver(second)
    fireEvent.drop(second)
    expect(props.onReorder).toHaveBeenCalledWith(0, 1)
  })

  describe('keyboard navigation (ARIA APG tablist pattern)', () => {
    it('only the active tab is a Tab stop; the rest are roved out', () => {
      renderBar()
      expect(screen.getByRole('tab', { name: /a\.md/ })).toHaveAttribute('tabindex', '0')
      expect(screen.getByRole('tab', { name: /b\.docx/ })).toHaveAttribute('tabindex', '-1')
    })

    it('ArrowRight/ArrowLeft move focus between tabs without selecting them', () => {
      const props = renderBar()
      const tabA = screen.getByRole('tab', { name: /a\.md/ })
      const tabB = screen.getByRole('tab', { name: /b\.docx/ })
      tabA.focus()

      fireEvent.keyDown(tabA, { key: 'ArrowRight' })
      expect(tabB).toHaveFocus()
      expect(tabB).toHaveAttribute('tabindex', '0')
      expect(tabA).toHaveAttribute('tabindex', '-1')
      // Moving focus alone must never switch the open document.
      expect(props.onSelect).not.toHaveBeenCalled()

      fireEvent.keyDown(tabB, { key: 'ArrowRight' })
      expect(tabA).toHaveFocus()

      fireEvent.keyDown(tabA, { key: 'ArrowLeft' })
      expect(tabB).toHaveFocus()
    })

    it('Home/End jump to the first/last tab', () => {
      renderBar()
      const tabA = screen.getByRole('tab', { name: /a\.md/ })
      const tabB = screen.getByRole('tab', { name: /b\.docx/ })
      tabA.focus()

      fireEvent.keyDown(tabA, { key: 'End' })
      expect(tabB).toHaveFocus()

      fireEvent.keyDown(tabB, { key: 'Home' })
      expect(tabA).toHaveFocus()
    })

    it('Enter/click on a focused tab still selects it', () => {
      const props = renderBar()
      const tabB = screen.getByRole('tab', { name: /b\.docx/ })
      fireEvent.click(tabB)
      expect(props.onSelect).toHaveBeenCalledWith('/b.docx')
    })
  })
})
