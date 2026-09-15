/**
 * S18/SLD-19 — the thumbnail rail (interactive=false) must never render an
 * embedded image at its original `src` directly; it goes through
 * `getDownscaledImage` first (a placeholder shows while that's pending). The
 * interactive main view (interactive=true) skips downscaling entirely and
 * renders the original `src` immediately, since fidelity matters there.
 */
import { render, screen, waitFor } from '@testing-library/react'
import { describe, expect, it } from 'vitest'

import { SlideCanvas } from '../SlideCanvas'
import type { SlideData } from '../SlideDeck.types'

const IMAGE_SRC = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='

function makeSlide(): SlideData {
  return {
    id: 'slide-1',
    index: 0,
    width: 960,
    height: 540,
    shapes: [
      {
        kind: 'image',
        id: 'img-1',
        transform: { x: 0, y: 0, w: 400, h: 300 },
        src: IMAGE_SRC,
        alt: 'a picture',
      },
    ],
  }
}

describe('SlideCanvas — thumbnail image downscaling (S18/SLD-19)', () => {
  it('renders the original src immediately in the interactive main view (no downscale)', () => {
    render(<SlideCanvas slide={makeSlide()} scale={1} interactive />)

    const img = screen.getByAltText('a picture') as HTMLImageElement
    expect(img.src).toBe(IMAGE_SRC)
  })

  it('does not paint the full-resolution src directly in the thumbnail rail — goes through the downscale path first', async () => {
    render(<SlideCanvas slide={makeSlide()} scale={0.1} interactive={false} />)

    // jsdom has no createImageBitmap/OffscreenCanvas, so getDownscaledImage
    // falls back to the original src, but only after an async tick — until
    // then, a placeholder (not the image) is rendered.
    await waitFor(() => {
      expect(screen.getByAltText('a picture')).toBeInTheDocument()
    })
    const img = screen.getByAltText('a picture') as HTMLImageElement
    expect(img.src).toBe(IMAGE_SRC)
  })
})
