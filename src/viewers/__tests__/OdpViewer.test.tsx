import { render, waitFor, within } from '@testing-library/react'
import { beforeAll, describe, expect, it } from 'vitest'

import { ViewerProvider } from '../shared/ViewerContext'
import { OdpViewer } from '../OdpViewer'
import type { LoadedFile } from '../../formats/types'
import { buildOdpFixtureZip } from '../slides/odp/__tests__/odpFixture'

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
  const zip = buildOdpFixtureZip()
  const content = await zip.generateAsync({ type: 'arraybuffer' })
  return { kind: 'binary', content, path: '/fixtures/sample.odp', format: 'odp' }
}

describe('OdpViewer', () => {
  it('parses a real archive end to end and renders the active slide\'s text', async () => {
    const file = await buildFixtureFile()

    const { container } = render(
      <ViewerProvider filePath={file.path}>
        <OdpViewer file={file} />
      </ViewerProvider>,
    )

    // The thumbnail rail renders the same slide, so scope queries to the main
    // viewport to avoid ambiguous "found it twice" matches.
    await waitFor(() => {
      expect(container.querySelector('.slide-deck__main')).not.toBeNull()
    })
    const main = within(container.querySelector('.slide-deck__main') as HTMLElement)

    await waitFor(() => {
      expect(main.getByText('Driven by EMEA')).toBeInTheDocument()
    })

    // S6 — the table's cells render as real text.
    expect(main.getByText('Growth')).toBeInTheDocument()
    // S11 — the draw:object frame becomes a labeled placeholder.
    expect(main.getByText('Embedded object not supported')).toBeInTheDocument()
  })

  it('S14 — excludes hidden slides from the deck shown to the user', async () => {
    const file = await buildFixtureFile()

    const { container } = render(
      <ViewerProvider filePath={file.path}>
        <OdpViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(container.querySelector('.slide-deck__main')).not.toBeNull()
    })
    const main = within(container.querySelector('.slide-deck__main') as HTMLElement)

    await waitFor(() => {
      expect(main.getByText('Driven by EMEA')).toBeInTheDocument()
    })

    // Only one thumbnail row (the visible slide) should exist in the rail — the
    // hidden slide never reaches SlideDeck at all.
    expect(container.querySelectorAll('.slide-deck__thumb-row')).toHaveLength(1)
  })
})
