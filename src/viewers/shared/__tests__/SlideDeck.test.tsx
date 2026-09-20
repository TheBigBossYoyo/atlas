import { act, fireEvent, render, within } from '@testing-library/react'
import { beforeAll, describe, expect, it, vi } from 'vitest'

import type { CancelSignal, ZipArchive } from '../../slides/shared/xmlUtils'
import { parsePptxSlides } from '../../slides/pptx/parser'
import { buildPptxFixtureZip } from '../../slides/pptx/__tests__/pptxFixture'
import { SlideDeck } from '../SlideDeck'
import type { SlideData, SlideTextBox } from '../SlideDeck.types'
import type { SlideDeckEditor } from '../SlideEditToolbar'

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

/** A no-op `SlideDeckEditor`, with individual handlers overridable per test. */
function makeEditor(overrides: Partial<SlideDeckEditor> = {}): SlideDeckEditor {
  return {
    canUndo: false,
    canRedo: false,
    onUndo: vi.fn(),
    onRedo: vi.fn(),
    onAddSlide: vi.fn(),
    onDuplicateSlide: vi.fn(),
    onDeleteSlide: vi.fn(),
    onMoveSlide: vi.fn(),
    onInsertTextBox: vi.fn(async () => null),
    onSave: vi.fn(),
    onSaveAs: vi.fn(),
    canEditNotes: false,
    onNotesChange: vi.fn(),
    onShapeText: vi.fn(),
    onShapeBox: vi.fn(async () => {}),
    onDeleteShape: vi.fn(),
    saveError: null,
    ...overrides,
  }
}

const NEW_TEXT_BOX_SOURCE_ID = '999'

function newEmptyTextBoxShape(): SlideTextBox {
  return {
    id: 'shape-new',
    sourceId: NEW_TEXT_BOX_SOURCE_ID,
    kind: 'text',
    transform: { x: 100, y: 100, w: 400, h: 50 },
    paragraphs: [{ runs: [], level: 0 }],
    text: '',
    movable: true,
  }
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

  // USR-16 — root cause of the "type immediately after Insert Text Box"
  // lost-text bug: the toolbar's `onInsertTextBox()` promise resolves with
  // the new shape's `sourceId` BEFORE the re-parsed `slides` prop (the one
  // actually containing that shape) has been committed and passed back down
  // — the previous code looked the shape up in the `activeSlide` closure
  // captured at click time, which is that same *stale* pre-insert value, so
  // the lookup always failed and nothing was ever selected or opened for
  // editing. This reproduces that exact ordering (resolve first, `slides`
  // update second) and asserts editing still opens once the shape actually
  // exists.
  it('USR-16 — opens the newly inserted text box for editing once it exists on the slide, even when onInsertTextBox resolves first', async () => {
    const slides = await loadFixtureSlides()
    let resolveInsert: (sourceId: string | null) => void = () => {}
    const onInsertTextBox = vi.fn(
      () => new Promise<string | null>((resolve) => {
        resolveInsert = resolve
      }),
    )
    const editor = makeEditor({ onInsertTextBox })

    const { getByRole, queryByRole, rerender } = render(
      <SlideDeck slides={slides} activeIndex={0} onSelect={vi.fn()} editor={editor} />,
    )

    fireEvent.click(getByRole('button', { name: 'Insert text box' }))
    expect(onInsertTextBox).toHaveBeenCalledTimes(1)

    // Resolve with the new shape's id while `slides` still doesn't have it —
    // the race that broke selection before the fix.
    await act(async () => {
      resolveInsert(NEW_TEXT_BOX_SOURCE_ID)
    })
    expect(queryByRole('textbox', { name: 'Edit slide text' })).not.toBeInTheDocument()

    // The deck's slides now actually gain the new shape, as a real re-parse
    // eventually delivers.
    const updatedSlides = slides.map((slide, i) =>
      i === 0 ? { ...slide, shapes: [...slide.shapes, newEmptyTextBoxShape()] } : slide,
    )
    rerender(<SlideDeck slides={updatedSlides} activeIndex={0} onSelect={vi.fn()} editor={editor} />)

    expect(getByRole('textbox', { name: 'Edit slide text' })).toBeInTheDocument()
  })

  // USR-16 — the toolbar button keeps focus after the click that triggers
  // it (moving focus elsewhere mid-insert opens a different bug: see
  // SlideEditToolbar's comment on why). Typing "immediately" routinely
  // includes a space or Enter, which the browser turns into a second native
  // click on a still-focused button — without a guard, that calls
  // `onInsertTextBox()` again and creates a duplicate, empty shape.
  it('USR-16 — a second click while the first Insert Text Box is still pending does not insert twice', async () => {
    const slides = await loadFixtureSlides()
    let resolveInsert: (sourceId: string | null) => void = () => {}
    const onInsertTextBox = vi.fn(
      () => new Promise<string | null>((resolve) => {
        resolveInsert = resolve
      }),
    )
    const editor = makeEditor({ onInsertTextBox })

    const { getByRole } = render(<SlideDeck slides={slides} activeIndex={0} onSelect={vi.fn()} editor={editor} />)

    const button = getByRole('button', { name: 'Insert text box' })
    fireEvent.click(button)
    // A space typed while the button is still focused re-fires a native
    // click before the first insert's promise has resolved.
    fireEvent.click(button)
    expect(onInsertTextBox).toHaveBeenCalledTimes(1)

    await act(async () => {
      resolveInsert(NEW_TEXT_BOX_SOURCE_ID)
    })

    // Once settled, a further click is a genuinely new, separate insert.
    fireEvent.click(button)
    expect(onInsertTextBox).toHaveBeenCalledTimes(2)
  })

  // USR-16 — belt-and-suspenders for the same bug: the button also stops a
  // Space/Enter keydown from ever reaching the browser's native "activate
  // the focused button" behavior in the first place, so it never gets the
  // chance to fire a second click at all.
  it('USR-16 — Space and Enter on the focused Insert Text Box button do not trigger its default (native click) behavior', () => {
    const slides = loadFixtureSlides()
    return slides.then((resolvedSlides) => {
      const editor = makeEditor()
      const { getByRole } = render(
        <SlideDeck slides={resolvedSlides} activeIndex={0} onSelect={vi.fn()} editor={editor} />,
      )
      const button = getByRole('button', { name: 'Insert text box' })

      // fireEvent.* returns false when the event's default was prevented.
      expect(fireEvent.keyDown(button, { key: ' ' })).toBe(false)
      expect(fireEvent.keyDown(button, { key: 'Enter' })).toBe(false)
      // Unrelated keys are left alone.
      expect(fireEvent.keyDown(button, { key: 'a' })).toBe(true)
    })
  })
})
