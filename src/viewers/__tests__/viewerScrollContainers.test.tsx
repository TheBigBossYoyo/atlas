/**
 * P1.10 (DAT-02 / UX-04) regression test.
 *
 * `.content--viewer .preview-panel` clips overflow (`overflow: hidden`), so
 * every non-markdown viewer must own its own scroll container or content
 * past one screen becomes unreachable. Code/RTF/ODT/Unknown previously had
 * no such container at all. This test locks in that each viewer's root
 * element carries the dedicated class styled with `overflow: auto` in
 * `__styles__/viewer-{code,rtf,odt,unknown}.css`, mirroring the working
 * `.docx-viewer` pattern.
 */
import type { ComponentType } from 'react'
import { render } from '@testing-library/react'
import { beforeAll, describe, expect, it } from 'vitest'

import type { LoadedFile } from '../../formats/types'
import { ViewerProvider } from '../shared/ViewerContext'
import { CodeViewer } from '../CodeViewer'
import { RtfViewer } from '../RtfViewer'
import { OdtViewer } from '../OdtViewer'
import { UnknownViewer } from '../UnknownViewer'
import { LegacyDocViewer } from '../LegacyDocViewer'
import { LegacyPptViewer } from '../LegacyPptViewer'

// CodeViewer calls useTheme(), which reads window.matchMedia on mount; jsdom
// doesn't implement it.
beforeAll(() => {
  if (typeof window.matchMedia !== 'function') {
    window.matchMedia = ((query: string) => ({
      matches: false,
      media: query,
      onchange: null,
      addListener: () => {},
      removeListener: () => {},
      addEventListener: () => {},
      removeEventListener: () => {},
      dispatchEvent: () => false,
    })) as unknown as typeof window.matchMedia
  }
})

function renderWithProvider(
  file: LoadedFile,
  Viewer: ComponentType<{ readonly file: LoadedFile }>,
) {
  return render(
    <ViewerProvider filePath={file.path}>
      <Viewer file={file} />
    </ViewerProvider>,
  )
}

describe('viewer scroll containers (DAT-02 / UX-04)', () => {
  it('CodeViewer root carries the scrollable .code-viewer class', () => {
    const file: LoadedFile = {
      kind: 'text',
      content: 'const x = 1\n'.repeat(200),
      path: '/tmp/example.ts',
      format: 'code',
    }

    const { container } = renderWithProvider(file, CodeViewer)
    expect(container.querySelector('.code-viewer')).not.toBeNull()
  })

  it('RtfViewer root carries the scrollable .rtf-viewer class', () => {
    const file: LoadedFile = {
      kind: 'binary',
      content: new ArrayBuffer(0),
      path: '/tmp/example.rtf',
      format: 'rtf',
    }

    const { container } = renderWithProvider(file, RtfViewer)
    expect(container.querySelector('.rtf-viewer')).not.toBeNull()
  })

  it('OdtViewer root carries the scrollable .odt-viewer class', () => {
    const file: LoadedFile = {
      kind: 'binary',
      content: new ArrayBuffer(0),
      path: '/tmp/example.odt',
      format: 'odt',
    }

    const { container } = renderWithProvider(file, OdtViewer)
    expect(container.querySelector('.odt-viewer')).not.toBeNull()
  })

  it('UnknownViewer root carries the scrollable .unknown-viewer class', () => {
    const file: LoadedFile = {
      kind: 'binary',
      content: new ArrayBuffer(0),
      path: '/tmp/example.xyz',
      format: 'unknown',
    }

    const { container } = renderWithProvider(file, UnknownViewer)
    expect(container.querySelector('.unknown-viewer')).not.toBeNull()
  })

  it('LegacyDocViewer root carries the scrollable .legacy-doc-viewer class', () => {
    const file: LoadedFile = {
      kind: 'binary',
      content: new ArrayBuffer(0),
      path: '/tmp/example.doc',
      format: 'doc',
    }

    const { container } = renderWithProvider(file, LegacyDocViewer)
    expect(container.querySelector('.legacy-doc-viewer')).not.toBeNull()
  })

  it('LegacyPptViewer root carries the .legacy-ppt-viewer class', () => {
    const file: LoadedFile = {
      kind: 'binary',
      content: new ArrayBuffer(0),
      path: '/tmp/example.ppt',
      format: 'ppt',
    }

    const { container } = renderWithProvider(file, LegacyPptViewer)
    expect(container.querySelector('.legacy-ppt-viewer')).not.toBeNull()
  })
})
