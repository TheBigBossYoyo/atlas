/**
 * ThemeMenu — accessible labeling (UX-09) and dropdown ARIA pattern (UX-14).
 */
import { fireEvent, render, screen } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { ThemeMenu } from '../ThemeMenu';
import { THEMES } from '../../types';

describe('ThemeMenu', () => {
  it('the icon-only trigger button has an aria-label (UX-09)', () => {
    render(<ThemeMenu current="light" themes={THEMES} onSelect={() => {}} />);
    expect(screen.getByRole('button', { name: 'Theme' })).toBeInTheDocument();
  });

  it('does not use the role="menu" ARIA pattern that had no keyboard support (UX-14)', () => {
    render(<ThemeMenu current="light" themes={THEMES} onSelect={() => {}} />);
    fireEvent.click(screen.getByRole('button', { name: 'Theme' }));

    expect(screen.queryByRole('menu')).not.toBeInTheDocument();
    expect(screen.getByRole('list', { name: 'Theme options' })).toBeInTheDocument();
  });

  it('marks the current theme with aria-current instead of the unsupported aria-checked pattern', () => {
    render(<ThemeMenu current="dark" themes={THEMES} onSelect={() => {}} />);
    fireEvent.click(screen.getByRole('button', { name: 'Theme' }));

    expect(screen.getByRole('button', { name: /^Dark/ })).toHaveAttribute('aria-current', 'true');
    expect(screen.getByRole('button', { name: /^Light/ })).not.toHaveAttribute('aria-current');
  });

  it('calls onSelect with the chosen theme and closes the menu', () => {
    const onSelect = vi.fn();
    render(<ThemeMenu current="light" themes={THEMES} onSelect={onSelect} />);
    fireEvent.click(screen.getByRole('button', { name: 'Theme' }));

    fireEvent.click(screen.getByRole('button', { name: /^Sepia/ }));

    expect(onSelect).toHaveBeenCalledWith('sepia');
    expect(screen.queryByRole('list', { name: 'Theme options' })).not.toBeInTheDocument();
  });

  // UX — picking an item (or Escape) closes the menu by unmounting the
  // `<ul>`; a focused item inside it is then removed from the DOM, and
  // without `useRestoreFocusOnClose` focus fell back to nothing (document
  // .body) instead of returning to the icon button that opened the menu.
  it('restores focus to the trigger button after selecting a theme', () => {
    render(<ThemeMenu current="light" themes={THEMES} onSelect={() => {}} />);
    const trigger = screen.getByRole('button', { name: 'Theme' });
    fireEvent.click(trigger);

    const sepiaOption = screen.getByRole('button', { name: /^Sepia/ });
    sepiaOption.focus();
    fireEvent.click(sepiaOption);

    expect(trigger).toHaveFocus();
  });
});
