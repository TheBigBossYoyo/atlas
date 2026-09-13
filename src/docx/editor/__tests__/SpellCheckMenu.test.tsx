import { fireEvent, render, screen } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'

import { SpellCheckMenu } from '../SpellCheckMenu'

describe('SpellCheckMenu', () => {
  it('renders the misspelled word and suggestions', () => {
    render(
      <SpellCheckMenu
        word="teh"
        suggestions={['the', 'tech']}
        x={10}
        y={20}
        onReplace={() => {}}
        onAddToDictionary={() => {}}
        onDismiss={() => {}}
      />,
    )
    expect(screen.getByText('teh')).toBeTruthy()
    expect(document.querySelector('.spellcheck-menu')).toBeTruthy()
    expect(screen.getByRole('menuitem', { name: 'the' })).toBeTruthy()
    expect(screen.getByRole('menuitem', { name: 'tech' })).toBeTruthy()
  })

  it('renders an empty hint when no suggestions exist', () => {
    render(
      <SpellCheckMenu
        word="zzqx"
        suggestions={[]}
        x={0}
        y={0}
        onReplace={() => {}}
        onAddToDictionary={() => {}}
        onDismiss={() => {}}
      />,
    )
    expect(screen.getByText('No suggestions')).toBeTruthy()
  })

  it('calls onReplace and dismisses when a suggestion is clicked', () => {
    const onReplace = vi.fn()
    const onDismiss = vi.fn()
    render(
      <SpellCheckMenu
        word="teh"
        suggestions={['the']}
        x={0}
        y={0}
        onReplace={onReplace}
        onAddToDictionary={() => {}}
        onDismiss={onDismiss}
      />,
    )
    fireEvent.click(screen.getByRole('menuitem', { name: 'the' }))
    expect(onReplace).toHaveBeenCalledWith('teh', 'the')
    expect(onDismiss).toHaveBeenCalledTimes(1)
  })

  it('calls onAddToDictionary and dismisses', () => {
    const onAdd = vi.fn()
    const onDismiss = vi.fn()
    render(
      <SpellCheckMenu
        word="atlas"
        suggestions={[]}
        x={0}
        y={0}
        onReplace={() => {}}
        onAddToDictionary={onAdd}
        onDismiss={onDismiss}
      />,
    )
    fireEvent.click(screen.getByRole('menuitem', { name: 'Add to dictionary' }))
    expect(onAdd).toHaveBeenCalledWith('atlas')
    expect(onDismiss).toHaveBeenCalledTimes(1)
  })

  it('dismisses on Escape', () => {
    const onDismiss = vi.fn()
    render(
      <SpellCheckMenu
        word="x"
        suggestions={[]}
        x={0}
        y={0}
        onReplace={() => {}}
        onAddToDictionary={() => {}}
        onDismiss={onDismiss}
      />,
    )
    fireEvent.keyDown(window, { key: 'Escape' })
    expect(onDismiss).toHaveBeenCalledTimes(1)
  })

  it('dismisses on outside mousedown', () => {
    const onDismiss = vi.fn()
    render(
      <div>
        <SpellCheckMenu
          word="x"
          suggestions={[]}
          x={0}
          y={0}
          onReplace={() => {}}
          onAddToDictionary={() => {}}
          onDismiss={onDismiss}
        />
        <button type="button" data-testid="outside">
          outside
        </button>
      </div>,
    )
    fireEvent.mouseDown(screen.getByTestId('outside'))
    expect(onDismiss).toHaveBeenCalledTimes(1)
  })
})
