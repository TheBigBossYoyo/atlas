/**
 * NewDocumentMenu (NEW-01) — the toolbar's "New document" dropdown had no
 * prior dedicated test file. Covers its basic contract (labeling, item
 * list, create + close, Escape-to-close via the shared shortcut dispatcher)
 * plus the same focus-restore-to-trigger gap fixed for ExportMenu/ThemeMenu.
 */
import { useState } from 'react';
import { fireEvent, render, screen } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { ShortcutManagerProvider } from '../../hooks/ShortcutManagerProvider';
import { NewDocumentMenu } from '../NewDocumentMenu';
import type { NewDocumentFormat } from '../../electron';

function StatefulMenu({ onCreate = vi.fn() }: { onCreate?: (format: NewDocumentFormat) => void }) {
  const [open, setOpen] = useState(true);
  return (
    <ShortcutManagerProvider>
      <NewDocumentMenu onCreate={onCreate} open={open} onOpenChange={setOpen} />
    </ShortcutManagerProvider>
  );
}

describe('NewDocumentMenu', () => {
  it('the trigger button is labeled and reflects open state via aria-expanded', () => {
    render(<NewDocumentMenu onCreate={() => {}} open={false} onOpenChange={() => {}} />);
    const trigger = screen.getByRole('button', { name: /New/ });
    expect(trigger).toHaveAttribute('aria-expanded', 'false');
  });

  it('lists every editable "New document" format when open', () => {
    render(<NewDocumentMenu onCreate={() => {}} open={true} onOpenChange={() => {}} />);
    for (const label of [
      'Markdown (.md)',
      'Word document (.docx)',
      'Excel workbook (.xlsx)',
      'OpenDocument Spreadsheet (.ods)',
      'PowerPoint presentation (.pptx)',
      'OpenDocument Presentation (.odp)',
    ]) {
      expect(screen.getByRole('button', { name: label })).toBeInTheDocument();
    }
  });

  // Regression: the panel inherited the shared right-aligned `.dropdown__menu`
  // geometry, but this trigger sits at the far left of the toolbar, so the
  // panel grew off the left edge of the window (measured at x = -76px — 76px
  // of it unreachable). jsdom has no layout, so the class is what is asserted
  // here; `tests/e2e/new-document.spec.ts` asserts the real geometry.
  it('left-aligns its panel, so it cannot grow off the left edge of the window', () => {
    render(<NewDocumentMenu onCreate={() => {}} open={true} onOpenChange={() => {}} />);
    expect(screen.getByRole('list')).toHaveClass('dropdown__menu', 'dropdown__menu--start');
  });

  it('calls onCreate with the chosen format and closes the menu', () => {
    const onCreate = vi.fn();
    const onOpenChange = vi.fn();
    render(<NewDocumentMenu onCreate={onCreate} open={true} onOpenChange={onOpenChange} />);

    fireEvent.click(screen.getByRole('button', { name: 'Word document (.docx)' }));

    expect(onCreate).toHaveBeenCalledWith('docx');
    expect(onOpenChange).toHaveBeenCalledWith(false);
  });

  it('closes on Escape', () => {
    const onOpenChange = vi.fn();
    render(
      <ShortcutManagerProvider>
        <NewDocumentMenu onCreate={() => {}} open={true} onOpenChange={onOpenChange} />
      </ShortcutManagerProvider>,
    );

    fireEvent.keyDown(window, { key: 'Escape' });

    expect(onOpenChange).toHaveBeenCalledWith(false);
  });

  // UX — picking an item closes the menu by unmounting the `<ul>`; a
  // focused item inside it is then removed from the DOM, and without
  // `useRestoreFocusOnClose` focus fell back to nothing (document.body)
  // instead of returning to the toolbar button that opened the menu.
  it('restores focus to the trigger button after creating a new document', () => {
    render(<StatefulMenu />);
    const trigger = screen.getByRole('button', { name: /New/ });

    const markdownOption = screen.getByRole('button', { name: 'Markdown (.md)' });
    markdownOption.focus();
    fireEvent.click(markdownOption);

    expect(trigger).toHaveFocus();
  });
});
