import { render, waitFor, within } from '@testing-library/react'
import { beforeAll, describe, expect, it } from 'vitest'

import { ViewerProvider } from '../shared/ViewerContext'
import { PptxViewer } from '../PptxViewer'
import type { LoadedFile } from '../../formats/types'
import { buildPptxFixtureZip } from '../slides/pptx/__tests__/pptxFixture'

// react-window's List measures its container via ResizeObserver, which jsdom
// doesn't implement — stub it so the thumbnail rail mounts instead of silently
// rendering zero rows.
beforeAll(() => {
  class ResizeObserverStub {
    observe(): void {}
    unobserve(): void {}
    disconnect(): void {}
  }

  Object.defineProperty(window, 'ResizeObserver', { writable: true, value: ResizeObserverStub })
})

async function buildFixtureFile(): Promise<LoadedFile> {
  const zip = buildPptxFixtureZip()
  const content = await zip.generateAsync({ type: 'arraybuffer' })
  return { kind: 'binary', content, path: '/fixtures/sample.pptx', format: 'pptx' }
}

describe('PptxViewer', () => {
  it('parses a real archive end to end and renders the active slide\'s text', async () => {
    const file = await buildFixtureFile()

    const { container } = render(
      <ViewerProvider filePath={file.path}>
        <PptxViewer file={file} />
      </ViewerProvider>,
    )

    // The thumbnail rail renders the same slide, so scope queries to the main
    // viewport to avoid ambiguous "found it twice" matches.
    await waitFor(() => {
      expect(container.querySelector('.slide-deck__main')).not.toBeNull()
    })
    const main = within(container.querySelector('.slide-deck__main') as HTMLElement)

    await waitFor(() => {
      expect(main.getByText('Revenue up 12%')).toBeInTheDocument()
    })

    expect(main.getByText('Q3 2024')).toBeInTheDocument()
    // S6 — the table's cells render as real text, not a JSON blob or stray fallback text.
    expect(main.getByText('EMEA')).toBeInTheDocument()
    // S11 — the chart graphicFrame becomes a labeled placeholder, not blank space.
    expect(main.getByText('Chart not supported')).toBeInTheDocument()
  })

  it('S14 — excludes hidden slides from the deck shown to the user', async () => {
    const file = await buildFixtureFile()

    const { container } = render(
      <ViewerProvider filePath={file.path}>
        <PptxViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(container.querySelector('.slide-deck__main')).not.toBeNull()
    })
    const main = within(container.querySelector('.slide-deck__main') as HTMLElement)

    await waitFor(() => {
      expect(main.getByText('Revenue up 12%')).toBeInTheDocument()
    })

    // The fixture has 3 slides: slide 1 (visible), slide 2 (hidden — excluded),
    // and slide 3 (missing archive entry -> S9 error placeholder, still shown).
    expect(container.querySelectorAll('.slide-deck__thumb-row')).toHaveLength(2)
  })
})
