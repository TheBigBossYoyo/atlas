/**
 * LOAD-22 — DropZone's copy used to say "Drop your Markdown file" / ".md,
 * .markdown, or .txt" despite the app supporting 14 formats. Regression
 * test pinning the corrected, format-agnostic copy.
 */
import { describe, expect, it } from 'vitest'
import { render, screen } from '@testing-library/react'

import { DropZone } from '../DropZone'

describe('DropZone', () => {
  it('renders nothing when not visible', () => {
    const { container } = render(<DropZone isVisible={false} />)
    expect(container).toBeEmptyDOMElement()
  })

  it('does not claim Atlas only supports Markdown (LOAD-22)', () => {
    render(<DropZone isVisible />)
    expect(screen.queryByText(/drop your markdown file/i)).not.toBeInTheDocument()
    expect(screen.queryByText(/\.md, \.markdown, or \.txt/i)).not.toBeInTheDocument()
  })

  it('shows format-agnostic copy when visible', () => {
    render(<DropZone isVisible />)
    expect(screen.getByText(/drop a file to open it/i)).toBeInTheDocument()
  })
})
