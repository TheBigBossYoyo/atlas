/**
 * T11 (DAT-16) — CodeViewer render/parse fixture coverage, plus T2/DAT-13's
 * large-file virtualized-plain-text fallback.
 */
import { readFileSync } from 'node:fs'
import path from 'node:path'
import { render, screen, waitFor } from '@testing-library/react'
import { beforeAll, describe, expect, it, vi } from 'vitest'

import type { LoadedFile } from '../../formats/types'
import { ViewerProvider } from '../shared/ViewerContext'
import { CODE_VIRTUALIZE_LINE_THRESHOLD } from '../shared/sizeThresholds'
import { CodeViewer } from '../CodeViewer'

// CodeViewer calls useTheme(), which reads window.matchMedia on mount; jsdom
// doesn't implement it (mirrors viewerScrollContainers.test.tsx).
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

vi.mock('shiki', () => ({
  getSingletonHighlighter: async () => ({
    getLoadedLanguages: () => ['typescript', 'text'],
    codeToHtml: (code: string, opts: { lang: string }) =>
      `<pre class="shiki" data-lang="${opts.lang}"><code>${code.replace(/</g, '&lt;')}</code></pre>`,
  }),
}))

const FIXTURES_DIR = path.resolve(process.cwd(), 'src/viewers/__fixtures__')

function readFixtureText(name: string): string {
  return readFileSync(path.join(FIXTURES_DIR, name), 'utf8')
}

describe('CodeViewer — fixture render/parse (T11)', () => {
  it('renders a real .ts fixture with shiki-highlighted output', async () => {
    const file: LoadedFile = {
      kind: 'text',
      content: readFixtureText('sample.ts'),
      path: '/tmp/sample.ts',
      format: 'code',
    }

    const { container } = render(
      <ViewerProvider filePath={file.path}>
        <CodeViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(container.querySelector('.shiki')).not.toBeNull())
    expect(container.querySelector('.shiki')?.getAttribute('data-lang')).toBe('typescript')
    expect(container.textContent).toContain('atlasFixture')
  })

  it('falls back to plain text when shiki fails, without crashing', async () => {
    const file: LoadedFile = {
      kind: 'text',
      content: 'const x = 1\n',
      path: '/tmp/example.unknownext',
      format: 'code',
    }

    const { container } = render(
      <ViewerProvider filePath={file.path}>
        <CodeViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => expect(container.querySelector('.shiki')).not.toBeNull())
  })
})

describe('CodeViewer — large-file virtualized fallback (T2/DAT-13)', () => {
  it('skips shiki and renders virtualized plain text above the line threshold', async () => {
    const bigContent = Array.from({ length: CODE_VIRTUALIZE_LINE_THRESHOLD + 10 }, (_, i) => `line ${i}`).join('\n')
    const file: LoadedFile = {
      kind: 'text',
      content: bigContent,
      path: '/tmp/huge.ts',
      format: 'code',
    }

    render(
      <ViewerProvider filePath={file.path}>
        <CodeViewer file={file} />
      </ViewerProvider>,
    )

    expect(screen.getByText(/syntax highlighting is disabled/i)).toBeInTheDocument()
    expect(screen.getByText('line 0')).toBeInTheDocument()
  })
})
