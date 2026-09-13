import React from 'react'
import { render, screen, waitFor } from '@testing-library/react'
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
})
