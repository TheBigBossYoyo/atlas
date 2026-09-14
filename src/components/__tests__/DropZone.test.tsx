/**
 * DropZone — accurate multi-format copy (UX-16/LOAD-22).
 *
 * Used to claim "Drop your Markdown file" / ".md, .markdown, or .txt" despite
 * full multi-format detection (src/formats/detect.ts) having landed.
 */
import { render, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { DropZone } from '../DropZone';

describe('DropZone (UX-16/LOAD-22)', () => {
  it('renders nothing when not visible', () => {
    const { container } = render(<DropZone isVisible={false} />);
    expect(container).toBeEmptyDOMElement();
  });

  it('no longer claims markdown-only support', () => {
    render(<DropZone isVisible={true} />);
    expect(screen.queryByText(/drop your markdown file/i)).not.toBeInTheDocument();
    expect(screen.queryByText(/\.md, \.markdown, or \.txt/i)).not.toBeInTheDocument();
  });

  it('describes real multi-format support when visible', () => {
    render(<DropZone isVisible={true} />);
    expect(screen.getByText('Drop your document')).toBeInTheDocument();
    expect(screen.getByText(/Word, Excel, PowerPoint/)).toBeInTheDocument();
  });
});
