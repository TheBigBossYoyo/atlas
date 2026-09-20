import { fireEvent, render, within } from '@testing-library/react'
import { beforeAll, describe, expect, it, vi } from 'vitest'

import type { CancelSignal, ZipArchive } from '../../slides/shared/xmlUtils'
import { parsePptxSlides } from '../../slides/pptx/parser'
import { buildPptxFixtureZip } from '../../slides/pptx/__tests__/pptxFixture'
import { SlideDeck } from '../SlideDeck'
import type { SlideData } from '../SlideDeck.types'

// react-window's List measures its container via ResizeObserver, which jsdom
// doesn't implement.
beforeAll(() => {
  class ResizeObserverStub {
    observe(): void {}
    unobserve(): void {}
    disconnect(): void {}
  }

  Object.defineProperty(window, 'ResizeObserver', { writable: true, value: ResizeObserverStub })
})

async function loadFixtureSlides(): Promise<ReadonlyArray<SlideData>> {
  const zip = buildPptxFixtureZip()
  const signal: CancelSignal = { cancelled: false }
  const slides = await parsePptxSlides(zip as unknown as ZipArchive, signal)
  return slides.filter(slide => !slide.hidden)
}

function mainSlideEl(container: HTMLElement): HTMLElement {
  return container.querySelector('.slide-deck__main .slide-deck__slide') as HTMLElement
}

describe('SlideDeck', () => {
  it('S21 — zoom in/out/reset controls scale the active slide beyond the auto-fit scale', async () => {
    const slides = await loadFixtureSlides()
    const { container, getByTitle } = render(<SlideDeck slides={slides} activeIndex={0} onSelect={vi.fn()} />)

    const before = mainSlideEl(container).style.transform

    fireEvent.click(getByTitle('Zoom in'))
    const afterZoomIn = mainSlideEl(container).style.transform
    expect(afterZoomIn).not.toBe(before)

    fireEvent.click(getByTitle('Reset zoom'))
    expect(mainSlideEl(container).style.transform).toBe(before)
  })

  it('S12 — the notes drawer toggles open/closed for a slide that has speaker notes', async () => {
    const slides = await loadFixtureSlides()
    const { container, getByTitle, queryByText, getByText } = render(
      <SlideDeck slides={slides} activeIndex={0} onSelect={vi.fn()} />,
    )

    const notesText = /Remember to mention EMEA growth drivers\./

    expect(queryByText(notesText)).not.toBeInTheDocument()

    fireEvent.click(getByTitle('Speaker notes'))
    expect(getByText(notesText)).toBeInTheDocument()

    fireEvent.click(getByTitle('Close notes'))
    expect(queryByText(notesText)).not.toBeInTheDocument()
    expect(container).toBeTruthy()
  })

  it('renders each thumbnail rail row scoped from the main viewport', async () => {
    const slides = await loadFixtureSlides()
    const { container } = render(<SlideDeck slides={slides} activeIndex={0} onSelect={vi.fn()} />)

    const rail = within(container.querySelector('.slide-deck__rail') as HTMLElement)
    expect(rail.getAllByRole('button').length).toBeGreaterThan(0)
  })

  it('A11Y — the thumbnail rail is a labeled list, and each thumbnail has an accessible name', async () => {
    const slides = await loadFixtureSlides()
    const { getByRole } = render(<SlideDeck slides={slides} activeIndex={0} onSelect={vi.fn()} />)

    const rail = getByRole('list', { name: 'Slide thumbnails' })
    const railButtons = within(rail).getAllByRole('button')
    expect(within(rail).getAllByRole('listitem').length).toBe(slides.length)
    expect(railButtons.length).toBe(slides.length)
    for (const button of railButtons) {
      expect(button).toHaveAccessibleName()
    }
  })
})
