/**
 * ExportMenu — accessible labeling (UX-09), format-specific items (UX-12),
 * and the dropdown's ARIA pattern (UX-14).
 */
import { fireEvent, render, screen } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { ExportMenu } from '../ExportMenu';

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
    const { rerender } = render(
      <ExportMenu onExport={() => {}} format="xlsx" open={true} onOpenChange={() => {}} canExportCsv={false} />,
    );
    expect(screen.queryByRole('button', { name: 'Export to CSV' })).not.toBeInTheDocument();

    rerender(<ExportMenu onExport={() => {}} format="xlsx" open={true} onOpenChange={() => {}} canExportCsv={true} />);
    expect(screen.getByRole('button', { name: 'Export to CSV' })).toBeInTheDocument();
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
});
