/**
 * Security regression (found during Wave 3-T review): rtf.js assigns a
 * rendered hyperlink's `href` directly from the RTF file's own
 * `\fldinst HYPERLINK "..."` string with no scheme validation at all
 * (`Renderer.buildHyperlinkElement`), so a crafted `.rtf` could previously
 * render a clickable `javascript:`-URI link that would run arbitrary
 * script in this renderer the moment a reader clicked it — RtfViewer
 * appended rtf.js's DOM nodes straight into the document with no
 * sanitization, unlike OdtViewer's existing DOMPurify pass. This test
 * mocks rtf.js to hand back exactly that shape of malicious node and
 * asserts RtfViewer neutralizes it end to end, in its own file so mocking
 * `rtf.js` doesn't affect RtfViewer.test.tsx's real-fixture coverage.
 */
import { render, waitFor } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'

import type { LoadedFile } from '../../formats/types'
import { ViewerProvider } from '../shared/ViewerContext'

vi.mock('rtf.js', () => {
  class FakeDocument {
    render(): Promise<Node[]> {
      const link = document.createElement('a')
      link.href = 'javascript:alert(1)'
      link.textContent = 'Click me'

      const para = document.createElement('p')
      para.textContent = 'Safe paragraph text'

      return Promise.resolve([link, para])
    }
  }

  return {
    RTFJS: { Document: FakeDocument, loggingEnabled: () => {} },
    WMFJS: { loggingEnabled: () => {} },
    EMFJS: { loggingEnabled: () => {} },
  }
})

describe('RtfViewer — hyperlink sanitization (security)', () => {
  it('neutralizes a javascript: URI hyperlink instead of rendering it live', async () => {
    const { RtfViewer } = await import('../RtfViewer')
    const file: LoadedFile = {
      kind: 'binary',
      content: new ArrayBuffer(8),
      path: '/tmp/malicious.rtf',
      format: 'rtf',
    }

    const { container } = render(
      <ViewerProvider filePath={file.path}>
        <RtfViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(container.textContent).toContain('Safe paragraph text'))

    const anchor = container.querySelector('a')
    // DOMPurify neutralizes a disallowed URI scheme by dropping the `href`
    // attribute entirely rather than leaving a live javascript: URI behind.
    const href = anchor?.getAttribute('href')
    expect(href == null || !/javascript:/i.test(href)).toBe(true)
  })
})
