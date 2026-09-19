/**
 * ExportMenu — accessible labeling (UX-09), format-specific items (UX-12),
 * and the dropdown's ARIA pattern (UX-14).
 */
import { useState } from 'react';
import { fireEvent, render, screen } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { ExportMenu } from '../ExportMenu';

// `open` is externally controlled by every other test below (a static prop
// plus a mock `onOpenChange`), which can't observe the trigger regaining
// focus once the menu actually closes — that needs a real state flip. This
// small stateful harness is only for the focus-restore test.
function StatefulExportMenu() {
  const [open, setOpen] = useState(true);
  return <ExportMenu onExport={() => {}} format="markdown" open={open} onOpenChange={setOpen} />;
}

describe('ExportMenu', () => {
  it('the icon-only trigger button has an aria-label (UX-09)', () => {
    render(<ExportMenu onExport={() => {}} format="markdown" open={false} onOpenChange={() => {}} />);
    expect(screen.getByRole('button', { name: 'Export' })).toBeInTheDocument();
  });

  it('does not use the role="menu" ARIA pattern that had no keyboard support (UX-14)', () => {
    render(<ExportMenu onExport={() => {}} format="markdown" open={true} onOpenChange={() => {}} />);
    expect(screen.queryByRole('menu')).not.toBeInTheDocument();
    expect(screen.getByRole('list', { name: 'Export options' })).toBeInTheDocument();
  });

  it('offers HTML/PDF/DOCX/Markdown for a markdown document', () => {
    render(<ExportMenu onExport={() => {}} format="markdown" open={true} onOpenChange={() => {}} />);
    for (const label of ['HTML', 'PDF', 'DOCX', 'Markdown']) {
      expect(screen.getByRole('button', { name: label })).toBeInTheDocument();
    }
    expect(screen.queryByRole('button', { name: 'Export to CSV' })).not.toBeInTheDocument();
  });

  it('offers "Export to CSV" for a csv document without needing canExportCsv (UX-12)', () => {
    render(<ExportMenu onExport={() => {}} format="csv" open={true} onOpenChange={() => {}} />);
    expect(screen.getByRole('button', { name: 'Export to CSV' })).toBeInTheDocument();
  });

  it('offers "Export to CSV" for another format only when canExportCsv is true (UX-12)', () => {
    // xlsx/ods themselves are exempted from this gate by X1 (see the next
    // test) — pptx has no native CSV-shaped content, so it only offers one
    // if a future viewer registers it via `getExportableContent`.
    const { rerender } = render(
      <ExportMenu onExport={() => {}} format="pptx" open={true} onOpenChange={() => {}} canExportCsv={false} />,
    );
    expect(screen.queryByRole('button', { name: 'Export to CSV' })).not.toBeInTheDocument();

    rerender(<ExportMenu onExport={() => {}} format="pptx" open={true} onOpenChange={() => {}} canExportCsv={true} />);
    expect(screen.getByRole('button', { name: 'Export to CSV' })).toBeInTheDocument();
  });

  it('offers PDF/CSV/"Save a copy" for xlsx/ods unconditionally (X1 — parsed directly from the file\'s own bytes, no viewer registration needed)', () => {
    render(<ExportMenu onExport={() => {}} format="xlsx" open={true} onOpenChange={() => {}} canExportCsv={false} />);
    expect(screen.getByRole('button', { name: 'Export to PDF' })).toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Export to CSV' })).toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Save a copy' })).toBeInTheDocument();
  });

  it('offers a "Save a copy" passthrough (value "copy") for a pdf document', () => {
    const onExport = vi.fn();
    render(<ExportMenu onExport={onExport} format="pdf" open={true} onOpenChange={() => {}} />);
    fireEvent.click(screen.getByRole('button', { name: 'Save a copy' }));
    expect(onExport).toHaveBeenCalledWith('copy');
  });

  it('offers "Export to HTML" alongside "Export to PDF" for text/code', () => {
    render(<ExportMenu onExport={() => {}} format="code" open={true} onOpenChange={() => {}} />);
    expect(screen.getByRole('button', { name: 'Export to PDF' })).toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Export to HTML' })).toBeInTheDocument();
  });

  it('calls onExport with the chosen format and closes the menu', () => {
    const onExport = vi.fn();
    const onOpenChange = vi.fn();
    render(<ExportMenu onExport={onExport} format="markdown" open={true} onOpenChange={onOpenChange} />);

    fireEvent.click(screen.getByRole('button', { name: 'DOCX' }));

    expect(onExport).toHaveBeenCalledWith('docx');
    expect(onOpenChange).toHaveBeenCalledWith(false);
  });

  it('disables the trigger when disabled is true', () => {
    render(<ExportMenu onExport={() => {}} format="markdown" open={false} onOpenChange={() => {}} disabled />);
    expect(screen.getByRole('button', { name: 'Export' })).toBeDisabled();
  });

  // UX — picking an item closes the menu by unmounting the `<ul>`; a
  // focused item inside it is then removed from the DOM, and without
  // `useRestoreFocusOnClose` focus fell back to nothing (document.body)
  // instead of returning to the icon button that opened the menu.
  it('restores focus to the trigger button after choosing an export format', () => {
    render(<StatefulExportMenu />);
    const trigger = screen.getByRole('button', { name: 'Export' });

    const docxOption = screen.getByRole('button', { name: 'DOCX' });
    docxOption.focus();
    fireEvent.click(docxOption);

    expect(trigger).toHaveFocus();
  });
});
