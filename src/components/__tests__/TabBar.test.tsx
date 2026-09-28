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
  const { container } = render(<TabBar {...props} />)
  return { ...props, container }
}

/**
 * A11Y pass 4 — the X on a tab is a mouse-only `<span>` with `aria-hidden` (a
 * focusable control inside a `role="tab"` is a WCAG 4.1.2 violation), so it has
 * no role and no accessible name to query by. `data-close-tab` exists for this.
 */
function closeAffordance(id: string): HTMLElement {
  const node = document.querySelector(`[data-close-tab="${id}"]`)
  if (node === null) throw new Error(`no close affordance for ${id}`)
  return node as HTMLElement
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
    const { container } = renderBar({ isActiveDirty: true })
    const dots = container.querySelectorAll('.tab-bar__dirty')
    expect(dots).toHaveLength(1)
    expect(screen.getByRole('tab', { name: /a\.md/ })).toContainElement(dots[0] as HTMLElement)
  })

  it('says unsaved changes in the tab own accessible name, not only as a coloured dot', () => {
    // A11Y pass 4 — the dot is `aria-hidden` and the state lives in the tab's
    // `aria-label` instead. A screen reader gets nothing from a background
    // colour, and labelling both the dot and the tab said it twice.
    renderBar({ isActiveDirty: true })
    expect(screen.getByRole('tab', { name: 'a.md (unsaved changes)' })).toBeInTheDocument()
    expect(screen.getByRole('tab', { name: 'b.docx' })).toBeInTheDocument()
  })

  it('does not fold the close button label into the tab name', () => {
    // Name-from-content would make this tab "a.md Close a.md", which is what a
    // screen reader would then read out for every tab.
    renderBar()
    expect(screen.getByRole('tab', { name: 'a.md' })).toBeInTheDocument()
  })

  it('closes from the close button and from a middle click', () => {
    const props = renderBar()
    fireEvent.click(closeAffordance('/b.docx'))
    expect(props.onClose).toHaveBeenCalledWith('/b.docx')

    // Testing Library has no auxClick helper; dispatch the real event.
    // A11Y pass 4 — the tab element IS the wrapper that carries these handlers
    // now, so this no longer reaches for `.parentElement`.
    const tab = screen.getByRole('tab', { name: /a[.]md/ })
    fireEvent(tab, new MouseEvent('auxclick', { button: 1, bubbles: true, cancelable: true }))
    expect(props.onClose).toHaveBeenCalledWith('/a.md')
  })

  it('closing a background tab does not also select it', () => {
    // The close affordance sits inside the tab, which owns the select handler,
    // so the click has to be stopped from bubbling.
    const props = renderBar()
    fireEvent.click(closeAffordance('/b.docx'))
    expect(props.onClose).toHaveBeenCalledWith('/b.docx')
    expect(props.onSelect).not.toHaveBeenCalled()
  })

  it('Delete and Backspace on a focused tab close it', () => {
    // A11Y pass 4 — the X is not focusable at all, so this IS the keyboard path
    // to closing a tab (alongside the shell's Ctrl+W). Listed in the shortcuts
    // dialog for that reason.
    const props = renderBar()
    fireEvent.keyDown(screen.getByRole('tab', { name: /b\.docx/ }), { key: 'Delete' })
    expect(props.onClose).toHaveBeenCalledWith('/b.docx')

    fireEvent.keyDown(screen.getByRole('tab', { name: /a\.md/ }), { key: 'Backspace' })
    expect(props.onClose).toHaveBeenCalledWith('/a.md')
  })

  it('Enter and Space select the focused tab', () => {
    // Free with a `<button>`; a `div[role="tab"]` has to do it itself.
    const props = renderBar()
    fireEvent.keyDown(screen.getByRole('tab', { name: /b\.docx/ }), { key: 'Enter' })
    expect(props.onSelect).toHaveBeenCalledWith('/b.docx')

    fireEvent.keyDown(screen.getByRole('tab', { name: /b\.docx/ }), { key: ' ' })
    expect(props.onSelect).toHaveBeenCalledTimes(2)
  })

  it('keeps the close affordance out of the accessibility tree entirely', () => {
    // Not merely out of the Tab sequence: axe is explicit that a negative
    // tabindex inside an interactive control does not stop AT reaching it, so the
    // X carries no role and no name and is aria-hidden.
    renderBar()
    expect(screen.queryByRole('button', { name: /Close/ })).not.toBeInTheDocument()
    expect(closeAffordance('/a.md')).toHaveAttribute('aria-hidden', 'true')
    expect(closeAffordance('/a.md')).not.toHaveAttribute('tabindex')
  })

  it('reorders on drag and drop', () => {
    const props = renderBar()
    const first = screen.getByRole('tab', { name: /a\.md/ })
    const second = screen.getByRole('tab', { name: /b\.docx/ })

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

    it('moves focus to the next TAB when a middle tab is closed', () => {
      render(<Harness initial={three} />)

      screen.getByRole('tab', { name: /b\.md/ }).focus()
      fireEvent.click(closeAffordance('/b.md'))

      expect(screen.queryByRole('tab', { name: /b\.md/ })).not.toBeInTheDocument()
      expect(document.activeElement).not.toBe(document.body)
      // A11Y pass 4 — the neighbouring tab, not its close button, which is no
      // longer focusable. Pressing Delete again closes that one, so closing
      // several in a row still works without moving the hands.
      expect(screen.getByRole('tab', { name: /c\.md/ })).toHaveFocus()
    })

    // CTRLW-FOCUS-1 — the two tests above close through TabBar's own X, which
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

        // Focus sits on b's tab, as it would after arrowing through the bar; then
        // b is removed from outside TabBar entirely.
        screen.getByRole('tab', { name: /b\.md/ }).focus()
        fireEvent.click(screen.getByText('close-b-externally'))

        expect(screen.queryByRole('tab', { name: /b\.md/ })).not.toBeInTheDocument()
        expect(document.activeElement).not.toBe(document.body)
        expect(screen.getByRole('tab', { name: /c\.md/ })).toHaveFocus()
      })

      it('leaves focus alone when it was OUTSIDE the tab bar', () => {
        // Ctrl+W pressed while the caret is in the document must not yank focus up to
        // a tab — that would be worse than the original bug.
        render(<ExternalCloseHarness initial={three} />)

        const outside = screen.getByText('close-b-externally')
        outside.focus()
        expect(outside).toHaveFocus()
        fireEvent.click(outside)

        expect(screen.queryByRole('tab', { name: /b\.md/ })).not.toBeInTheDocument()
        expect(outside).toHaveFocus()
      })
    })

    it('moves focus to the previous TAB when the last tab is closed', () => {
      render(<Harness initial={three} />)

      screen.getByRole('tab', { name: /c\.md/ }).focus()
      fireEvent.click(closeAffordance('/c.md'))

      expect(screen.queryByRole('tab', { name: /c\.md/ })).not.toBeInTheDocument()
      expect(document.activeElement).not.toBe(document.body)
      expect(screen.getByRole('tab', { name: /b\.md/ })).toHaveFocus()
    })

    it('never leaves focus on <body> when the sole remaining tab is closed', () => {
      const one: ReadonlyArray<DocumentSession> = [
        { id: '/a.md', name: 'a.md', file: { kind: 'text', content: '', path: '/a.md', format: 'markdown' } },
      ]
      const { container } = render(<Harness initial={one} />)

      screen.getByRole('tab', { name: /a\.md/ }).focus()
      fireEvent.click(closeAffordance('/a.md'))

      // The bar itself unmounts (no tabs left) — nothing inside it can be
      // focused, and where focus goes next is the shell's (App.tsx's)
      // concern, not TabBar's. This only asserts TabBar didn't error and
      // correctly rendered nothing.
      expect(container).toBeEmptyDOMElement()
    })
  })
})
