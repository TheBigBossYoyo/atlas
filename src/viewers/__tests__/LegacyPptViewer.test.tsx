import { render, waitFor, within } from '@testing-library/react'
import { beforeAll, describe, expect, it } from 'vitest'

import type { LoadedFile } from '../../formats/types'
import { buildMinimalPptBytes } from '../../legacy/__tests__/fixtures'
import { ViewerProvider } from '../shared/ViewerContext'
import { LegacyPptViewer } from '../LegacyPptViewer'

// SlideDeck's thumbnail rail (react-window) measures its container via
// ResizeObserver, which jsdom doesn't implement — stub it, mirroring
// OdpViewer.test.tsx/PptxViewer.test.tsx.
beforeAll(() => {
  class ResizeObserverStub {
    observe(): void {}
    unobserve(): void {}
    disconnect(): void {}
  }
  Object.defineProperty(window, 'ResizeObserver', { writable: true, value: ResizeObserverStub })
})

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return new Uint8Array(bytes).buffer
}

function fixtureFile(): LoadedFile {
  const bytes = buildMinimalPptBytes([
    { title: 'Slide One Title', body: 'First bullet\rSecond bullet' },
    { title: 'Slide Two Title' },
  ])
  return {
    kind: 'binary',
    content: toArrayBuffer(bytes),
    path: '/fixtures/sample.ppt',
    format: 'ppt',
  }
}

describe('LegacyPptViewer', () => {
  it('parses a real .ppt end to end and renders the active slide\'s text', async () => {
    const file = fixtureFile()

    const { container } = render(
      <ViewerProvider filePath={file.path}>
        <LegacyPptViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(container.querySelector('.slide-deck__main')).not.toBeNull()
    })
    const main = within(container.querySelector('.slide-deck__main') as HTMLElement)

    await waitFor(() => {
      expect(main.getByText('Slide One Title')).toBeInTheDocument()
    })
    expect(main.getByText('First bullet')).toBeInTheDocument()
    expect(main.getByText('Second bullet')).toBeInTheDocument()
  })

  it('shows the legacy-format read-only banner naming the modern format', async () => {
    const file = fixtureFile()

    const { getByRole } = render(
      <ViewerProvider filePath={file.path}>
        <LegacyPptViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(getByRole('note')).toHaveTextContent(/PowerPoint 97-2003/)
    })
    expect(getByRole('note')).toHaveTextContent(/\.pptx/)
  })

  it('shows a friendly error instead of crashing on a file that is not a valid .ppt', async () => {
    const file: LoadedFile = { kind: 'binary', content: new ArrayBuffer(3), path: '/fixtures/bad.ppt', format: 'ppt' }

    const { container } = render(
      <ViewerProvider filePath={file.path}>
        <LegacyPptViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(container.querySelector('.legacy-ppt-viewer--error')).not.toBeNull()
    })
  })
})
