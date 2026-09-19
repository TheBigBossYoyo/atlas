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
