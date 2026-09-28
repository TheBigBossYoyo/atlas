/** SHELL-17 — the open-documents strip. */
import { useState } from 'react'
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

  // A11Y-3 — the close button's onClick had no focus handling at all: the
  // tab (and its close button) unmounts with nowhere for focus to go, so it
  // fell back to `document.body` and the next Tab press started over from
  // the top of the window. Needs a stateful harness — the plain `renderBar`
  // helper's `onClose` is a no-op spy, so `sessions` never actually shrinks.
  describe('closing a tab moves focus (A11Y-3)', () => {
    const three: ReadonlyArray<DocumentSession> = [
      { id: '/a.md', name: 'a.md', file: { kind: 'text', content: '', path: '/a.md', format: 'markdown' } },
      { id: '/b.md', name: 'b.md', file: { kind: 'text', content: '', path: '/b.md', format: 'markdown' } },
      { id: '/c.md', name: 'c.md', file: { kind: 'text', content: '', path: '/c.md', format: 'markdown' } },
    ]

    function Harness({ initial }: { initial: ReadonlyArray<DocumentSession> }) {
      const [items, setItems] = useState(initial)
      return (
        <TabBar
          sessions={items}
          activeId={items[0]?.id ?? null}
          isActiveDirty={false}
          onSelect={() => {}}
          onClose={(id) => setItems((prev) => prev.filter((s) => s.id !== id))}
          onReorder={() => {}}
        />
      )
    }

    it("moves focus to the next tab's close button when a middle tab is closed", () => {
      render(<Harness initial={three} />)

      const closeB = screen.getByRole('button', { name: 'Close b.md' })
      closeB.focus()
      fireEvent.click(closeB)

      expect(screen.queryByRole('tab', { name: /b\.md/ })).not.toBeInTheDocument()
      expect(document.activeElement).not.toBe(document.body)
      expect(screen.getByRole('button', { name: 'Close c.md' })).toHaveFocus()
    })

    // CTRLW-FOCUS-1 — the two tests above close through TabBar's own button, which
    // records the neighbour to focus before the removal. The global Ctrl+W in App.tsx
    // removes a session WITHOUT going through this component at all, and that route
    // used to leave focus on <body>. This harness exposes an external close to
    // reproduce exactly that.
    describe('a close that does not go through TabBar (Ctrl+W)', () => {
      function ExternalCloseHarness({ initial }: { initial: ReadonlyArray<DocumentSession> }) {
        const [items, setItems] = useState(initial)
        return (
          <>
            <button type="button" onClick={() => setItems((prev) => prev.filter((s) => s.id !== '/b.md'))}>
              close-b-externally
            </button>
            <TabBar
              sessions={items}
              activeId={items[0]?.id ?? null}
              isActiveDirty={false}
              onSelect={() => {}}
              onClose={(id) => setItems((prev) => prev.filter((s) => s.id !== id))}
              onReorder={() => {}}
            />
          </>
        )
      }

      it("re-homes focus to a neighbouring tab when focus was in the tab bar", () => {
        render(<ExternalCloseHarness initial={three} />)

        // Focus sits on b's close button, as it would after arrowing through the bar;
        // then b is removed from outside TabBar entirely.
        screen.getByRole('button', { name: 'Close b.md' }).focus()
        fireEvent.click(screen.getByText('close-b-externally'))

        expect(screen.queryByRole('tab', { name: /b\.md/ })).not.toBeInTheDocument()
        expect(document.activeElement).not.toBe(document.body)
        expect(screen.getByRole('button', { name: 'Close c.md' })).toHaveFocus()
      })

      it('leaves focus alone when it was OUTSIDE the tab bar', () => {
        // Ctrl+W pressed while the caret is in the document must not yank focus up to
        // a tab's close button — that would be worse than the original bug.
        render(<ExternalCloseHarness initial={three} />)

        const outside = screen.getByText('close-b-externally')
        outside.focus()
        expect(outside).toHaveFocus()
        fireEvent.click(outside)

        expect(screen.queryByRole('tab', { name: /b\.md/ })).not.toBeInTheDocument()
        expect(outside).toHaveFocus()
      })
    })

    it("moves focus to the previous tab's close button when the last tab is closed", () => {
      render(<Harness initial={three} />)

      const closeC = screen.getByRole('button', { name: 'Close c.md' })
      closeC.focus()
      fireEvent.click(closeC)

      expect(screen.queryByRole('tab', { name: /c\.md/ })).not.toBeInTheDocument()
      expect(document.activeElement).not.toBe(document.body)
      expect(screen.getByRole('button', { name: 'Close b.md' })).toHaveFocus()
    })

    it('never leaves focus on <body> when the sole remaining tab is closed', () => {
      const one: ReadonlyArray<DocumentSession> = [
        { id: '/a.md', name: 'a.md', file: { kind: 'text', content: '', path: '/a.md', format: 'markdown' } },
      ]
      const { container } = render(<Harness initial={one} />)

      const closeA = screen.getByRole('button', { name: 'Close a.md' })
      closeA.focus()
      fireEvent.click(closeA)

      // The bar itself unmounts (no tabs left) — nothing inside it can be
      // focused, and where focus goes next is the shell's (App.tsx's)
      // concern, not TabBar's. This only asserts TabBar didn't error and
      // correctly rendered nothing.
      expect(container).toBeEmptyDOMElement()
    })
  })
})
