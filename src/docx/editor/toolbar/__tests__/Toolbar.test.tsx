import { render, fireEvent } from '@testing-library/react';
import { describe, it, expect, vi } from 'vitest';
import { Toolbar } from '../Toolbar';
import type { ToolbarState } from '../toolbarTypes';

const defaultState: ToolbarState = {
  activeFormats: new Set(),
  alignment: null,
  fontFamily: null,
  fontSizePt: null,
  styleId: null,
  trackChanges: false,
  spellCheck: false,
};

describe('Toolbar', () => {
  it('renders all 4 tabs', () => {
    const { getByText } = render(<Toolbar state={defaultState} onCommand={vi.fn()} />);
    expect(getByText('Home')).toBeInTheDocument();
    expect(getByText('Insert')).toBeInTheDocument();
    expect(getByText('Layout')).toBeInTheDocument();
    expect(getByText('Review')).toBeInTheDocument();
  });

  it('defaults to Home tab', () => {
    const { getByTitle } = render(<Toolbar state={defaultState} onCommand={vi.fn()} />);
    // "Undo" is in Home tab
    expect(getByTitle('Undo')).toBeInTheDocument();
  });

  it('switches tabs in uncontrolled mode', () => {
    const { getByText, getByTitle, queryByTitle } = render(<Toolbar state={defaultState} onCommand={vi.fn()} />);
    
    // Switch to Insert tab
    fireEvent.click(getByText('Insert'));
    
    // Check if Insert-specific button is visible
    expect(getByTitle('Page Break')).toBeInTheDocument();
    expect(queryByTitle('Undo')).not.toBeInTheDocument();
  });

  it('respects controlled tab mode', () => {
    const onTabChange = vi.fn();
    const { getByLabelText } = render(
      <Toolbar state={defaultState} onCommand={vi.fn()} activeTab="layout" onTabChange={onTabChange} />
    );

    // Check Layout tab content — Columns is genuinely unimplemented (D18) so
    // it renders disabled with an explanatory tooltip rather than "Columns".
    const columnsButton = getByLabelText('Columns');
    expect(columnsButton).toBeInTheDocument();
    expect(columnsButton).toBeDisabled();
    expect(columnsButton).toHaveAttribute('title', 'Not yet supported');
  });

  it('calls onTabChange when clicked in controlled mode', () => {
    const onTabChange = vi.fn();
    const { getByText } = render(
      <Toolbar state={defaultState} onCommand={vi.fn()} activeTab="layout" onTabChange={onTabChange} />
    );
    
    fireEvent.click(getByText('Review'));
    expect(onTabChange).toHaveBeenCalledWith('review');
  });

  it('emits toggle-bold command on Bold button click', () => {
    const onCommand = vi.fn();
    const { getByTitle } = render(<Toolbar state={defaultState} onCommand={onCommand} />);
    
    fireEvent.click(getByTitle('toggle-bold'));
    expect(onCommand).toHaveBeenCalledWith({ kind: 'toggle-bold' });
  });

  it('emits toggle-italic command on Italic button click', () => {
    const onCommand = vi.fn();
    const { getByTitle } = render(<Toolbar state={defaultState} onCommand={onCommand} />);
    
    fireEvent.click(getByTitle('toggle-italic'));
    expect(onCommand).toHaveBeenCalledWith({ kind: 'toggle-italic' });
  });

  it('visually distinguishes active state', () => {
    const stateWithBold = { ...defaultState, activeFormats: new Set(['bold'] as const) };
    const { getByTitle } = render(<Toolbar state={stateWithBold} onCommand={vi.fn()} />);
    
    const boldButton = getByTitle('toggle-bold');
    expect(boldButton.className).toContain('docx-toolbar__button--active');
    
    const italicButton = getByTitle('toggle-italic');
    expect(italicButton.className).not.toContain('docx-toolbar__button--active');
  });

  it('emits set-font-family when select changes', () => {
    const onCommand = vi.fn();
    const { getByDisplayValue } = render(
      <Toolbar state={{ ...defaultState, fontFamily: 'Arial' }} onCommand={onCommand} />
    );
    
    const select = getByDisplayValue('Arial');
    fireEvent.change(select, { target: { value: 'Verdana' } });
    
    expect(onCommand).toHaveBeenCalledWith({ kind: 'set-font-family', family: 'Verdana' });
  });

  it('emits set-font-size when select changes', () => {
    const onCommand = vi.fn();
    const { getByDisplayValue } = render(
      <Toolbar state={{ ...defaultState, fontSizePt: 12 }} onCommand={onCommand} />
    );
    
    const select = getByDisplayValue('12');
    fireEvent.change(select, { target: { value: '16' } });
    
    expect(onCommand).toHaveBeenCalledWith({ kind: 'set-font-size', sizePt: 16 });
  });

  it('emits insert-table with correct dimensions', () => {
    const onCommand = vi.fn();
    const { getByTitle, getByText } = render(
      <Toolbar state={defaultState} onCommand={onCommand} activeTab="insert" />
    );
    
    // Open table popover
    fireEvent.click(getByTitle('Insert Table'));
    
    // Label should exist initially
    expect(getByText('1x1 Table')).toBeInTheDocument();
    
    // It should have grid cells. Let's find one by grabbing the grid container and getting its children
    const cells = document.querySelectorAll('.docx-toolbar__table-cell');
    expect(cells.length).toBe(80); // 8 rows * 10 cols
    
    // Hover over row 3 (idx 2), col 4 (idx 3) -> cell idx is 2*10 + 3 = 23
    fireEvent.mouseEnter(cells[23]);
    expect(getByText('3x4 Table')).toBeInTheDocument();
    
    // Click it
    fireEvent.click(cells[23]);

    expect(onCommand).toHaveBeenCalledWith({ kind: 'insert-table', rows: 3, cols: 4 });
  });

  // ---------------------------------------------------------------------------
  // X3/UX-09 — icon-only buttons carry an aria-label mirroring their tooltip
  // ---------------------------------------------------------------------------

  it.each(['home', 'insert', 'layout', 'review'] as const)(
    'every icon-only button on the %s tab has an aria-label',
    (tab) => {
      const { container } = render(<Toolbar state={defaultState} onCommand={vi.fn()} activeTab={tab} />);
      const buttons = Array.from(container.querySelectorAll('button')).filter(
        (button) => button.getAttribute('role') !== 'tab',
      );
      expect(buttons.length).toBeGreaterThan(0);
      for (const button of buttons) {
        expect(button.getAttribute('aria-label')).toBeTruthy();
      }
    },
  );

  it('the font and size selects are labeled for screen readers', () => {
    const { getByLabelText } = render(<Toolbar state={defaultState} onCommand={vi.fn()} />);
    expect(getByLabelText('Font')).toBeInTheDocument();
    expect(getByLabelText('Font size')).toBeInTheDocument();
  });

  // ---------------------------------------------------------------------------
  // D18/DXE-12 — genuinely unimplemented controls are disabled with a tooltip
  // ---------------------------------------------------------------------------

  it('disables Header/Footer insertion with a "not yet supported" tooltip', () => {
    const { getByLabelText } = render(<Toolbar state={defaultState} onCommand={vi.fn()} activeTab="insert" />);
    expect(getByLabelText('Header')).toBeDisabled();
    expect(getByLabelText('Footer')).toBeDisabled();
  });

  it('does not disable the now-implemented Hyperlink/Page Break/Table controls', () => {
    const { getByLabelText } = render(<Toolbar state={defaultState} onCommand={vi.fn()} activeTab="insert" />);
    expect(getByLabelText('Hyperlink')).not.toBeDisabled();
    expect(getByLabelText('Page Break')).not.toBeDisabled();
    expect(getByLabelText('Insert Table')).not.toBeDisabled();
  });

  it('a disabled button never fires its command when clicked', () => {
    const onCommand = vi.fn();
    const { getByLabelText } = render(<Toolbar state={defaultState} onCommand={onCommand} activeTab="insert" />);
    fireEvent.click(getByLabelText('Header'));
    expect(onCommand).not.toHaveBeenCalled();
  });

  // ---------------------------------------------------------------------------
  // DXE-23 — color picker and table-size picker are keyboard-operable
  // ---------------------------------------------------------------------------

  it('color swatches are real buttons reachable by keyboard and labeled', () => {
    const onCommand = vi.fn();
    const { getByTitle } = render(<Toolbar state={defaultState} onCommand={onCommand} />);

    fireEvent.click(getByTitle('Font Color'));
    const swatch = document.querySelector('.docx-toolbar__color-swatch') as HTMLButtonElement;
    expect(swatch.tagName).toBe('BUTTON');
    expect(swatch.getAttribute('aria-label')).toBeTruthy();

    swatch.focus();
    expect(document.activeElement).toBe(swatch);
    fireEvent.click(swatch);
    expect(onCommand).toHaveBeenCalledWith(expect.objectContaining({ kind: 'set-font-color' }));
  });

  it('table-size cells are real buttons that can be focused and activated with Enter/Space semantics', () => {
    const onCommand = vi.fn();
    const { getByTitle } = render(<Toolbar state={defaultState} onCommand={onCommand} activeTab="insert" />);

    fireEvent.click(getByTitle('Insert Table'));
    const cells = document.querySelectorAll<HTMLButtonElement>('.docx-toolbar__table-cell');
    expect(cells[0].tagName).toBe('BUTTON');

    cells[7].focus();
    fireEvent.focus(cells[7]);
    expect(document.querySelector('.docx-toolbar__table-label')?.textContent).toBe('1x8 Table');

    fireEvent.click(cells[7]);
    expect(onCommand).toHaveBeenCalledWith({ kind: 'insert-table', rows: 1, cols: 8 });
  });
});
