/**
 * RawEditor — I18N-1.
 *
 * `RawEditor.tsx` used to render `<span>Markdown Source</span>` and
 * `placeholder="Type or paste markdown here..."` as hard-coded English, and
 * its `<textarea>` had no accessible name at all: no `<label>`, no
 * `aria-label`, no `aria-labelledby`. Its only "name" was the placeholder,
 * which disappears the moment the field has text, so a screen-reader user
 * tabbing into a non-empty editor heard nothing (WCAG 3.3.2).
 *
 * These assertions query by role/name and by the translated text, not by
 * placeholder, precisely because the placeholder is not a reliable name.
 */
import { render, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { RawEditor } from '../RawEditor';

describe('RawEditor (I18N-1)', () => {
  it('renders the translated header text', () => {
    render(<RawEditor markdown="" onChange={() => {}} />);
    expect(screen.getByText('Markdown Source')).toBeInTheDocument();
  });

  it('exposes an accessible name on the textarea, queryable by role', () => {
    render(<RawEditor markdown="" onChange={() => {}} />);
    expect(screen.getByRole('textbox', { name: 'Markdown Source' })).toBeInTheDocument();
  });

  it('keeps its accessible name once the field has content (placeholder is gone, aria-label is not)', () => {
    render(<RawEditor markdown="# Some existing content" onChange={() => {}} />);
    const textarea = screen.getByRole('textbox', { name: 'Markdown Source' });
    expect(textarea).toHaveValue('# Some existing content');
  });

  it('shows the translated placeholder when empty', () => {
    render(<RawEditor markdown="" onChange={() => {}} />);
    expect(screen.getByPlaceholderText('Type or paste markdown here...')).toBeInTheDocument();
  });
});
