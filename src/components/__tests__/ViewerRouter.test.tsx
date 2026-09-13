import React from 'react'
import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import { describe, it, expect, vi, beforeEach } from 'vitest'
import { ViewerRouter } from '../ViewerRouter'
import type { LoadedFile } from '../../formats/types'

// ---------------------------------------------------------------------------
// Registry mock — overridden per test
// ---------------------------------------------------------------------------
vi.mock('../../formats/registry', () => ({
  viewerRegistry: new Proxy({} as Record<string, () => Promise<unknown>>, {
    get(_target, prop: string) {
      return () => (registryMocks[prop] ?? registryMocks['__default__'])()
    },
  }),
}))

const registryMocks: Record<string, () => Promise<unknown>> = {}

function setRegistryMock(format: string, loader: () => Promise<unknown>): void {
  registryMocks[format] = loader
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
const mdFile: LoadedFile = {
  kind: 'text',
  content: '# hi',
  path: '/a.md',
  format: 'markdown',
}

const mdFile2: LoadedFile = {
  kind: 'text',
  content: '# bye',
  path: '/b.md',
  format: 'markdown',
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------
describe('ViewerRouter', () => {
  beforeEach(() => {
    // Clear all mocks between tests
    for (const key of Object.keys(registryMocks)) {
      delete registryMocks[key]
    }
  })

  it('1. shows loading state while registry promise is pending', () => {
    // Never-resolving promise
    setRegistryMock('markdown', () => new Promise(() => {}))

    render(<ViewerRouter file={mdFile} />)

    expect(screen.getByRole('status')).toBeInTheDocument()
    expect(screen.getByRole('status').textContent).toContain('Loading markdown')
  })

  it('2. renders viewer component after registry resolves', async () => {
    const FakeViewer = () => <div>VIEWER_OK</div>
    setRegistryMock('markdown', () => Promise.resolve(FakeViewer))

    render(<ViewerRouter file={mdFile} />)

    await waitFor(() => {
      expect(screen.getByText('VIEWER_OK')).toBeInTheDocument()
    })
  })

  it('3. shows error boundary when viewer throws on render', async () => {
    const consoleSpy = vi.spyOn(console, 'error').mockImplementation(() => {})

    const CrashingViewer = (): React.ReactElement => {
      throw new Error('boom from viewer')
    }
    setRegistryMock('markdown', () => Promise.resolve(CrashingViewer))

    render(<ViewerRouter file={mdFile} />)

    await waitFor(() => {
      expect(screen.getByRole('alert')).toBeInTheDocument()
    })
    expect(screen.getByRole('alert').textContent).toContain('Viewer crashed')
    expect(screen.getByRole('alert').textContent).toContain('boom from viewer')

    consoleSpy.mockRestore()
  })

  it('4. switching file.path creates a new lazy component', async () => {
    let callCount = 0

    const FakeViewerA = () => <div>VIEWER_A</div>
    const FakeViewerB = () => <div>VIEWER_B</div>

    setRegistryMock('markdown', () => {
      callCount++
      return callCount === 1 ? Promise.resolve(FakeViewerA) : Promise.resolve(FakeViewerB)
    })

    const { rerender } = render(<ViewerRouter file={mdFile} />)

    await waitFor(() => {
      expect(screen.getByText('VIEWER_A')).toBeInTheDocument()
    })

    rerender(<ViewerRouter file={mdFile2} />)

    await waitFor(() => {
      expect(screen.getByText('VIEWER_B')).toBeInTheDocument()
    })

    expect(callCount).toBe(2)
  })

  it('5. a crash on one file does not brick the viewer for a subsequently opened valid file (LOAD-08/RUN-09)', async () => {
    const consoleSpy = vi.spyOn(console, 'error').mockImplementation(() => {})

    const CrashingViewer = (): React.ReactElement => {
      throw new Error('boom from crashing file')
    }
    const FakeViewerB = () => <div>VIEWER_B_OK</div>

    setRegistryMock('markdown', () => Promise.resolve(CrashingViewer))

    const { rerender } = render(<ViewerRouter file={mdFile} />)

    await waitFor(() => {
      expect(screen.getByRole('alert')).toBeInTheDocument()
    })

    // A different, valid file is opened next — the crash screen must clear.
    setRegistryMock('markdown', () => Promise.resolve(FakeViewerB))
    rerender(<ViewerRouter file={mdFile2} />)

    await waitFor(() => {
      expect(screen.getByText('VIEWER_B_OK')).toBeInTheDocument()
    })
    expect(screen.queryByRole('alert')).not.toBeInTheDocument()

    consoleSpy.mockRestore()
  })

  it('6. "Try again" recovers from a transient chunk-load failure (LOAD-09/DAT-21)', async () => {
    const consoleSpy = vi.spyOn(console, 'error').mockImplementation(() => {})

    let attempts = 0
    const RecoveredViewer = () => <div>RECOVERED</div>

    setRegistryMock('markdown', () => {
      attempts += 1
      if (attempts === 1) {
        return Promise.reject(new Error('Failed to fetch dynamically imported module'))
      }
      return Promise.resolve(RecoveredViewer)
    })

    render(<ViewerRouter file={mdFile} />)

    await waitFor(() => {
      expect(screen.getByRole('alert')).toBeInTheDocument()
    })

    fireEvent.click(screen.getByRole('button', { name: /try again/i }))

    await waitFor(() => {
      expect(screen.getByText('RECOVERED')).toBeInTheDocument()
    })

    expect(attempts).toBe(2)
    consoleSpy.mockRestore()
  })
})
