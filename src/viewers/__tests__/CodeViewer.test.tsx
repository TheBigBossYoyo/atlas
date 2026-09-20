/**
 * USR-18/USR-19 — code files open in a real editor (CodeMirror 6) with save,
 * word wrap and an explicit Run action. T11's fixture coverage is kept: a
 * real .ts fixture still renders through the editor.
 */
import { readFileSync } from 'node:fs'
import path from 'node:path'
import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import { beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'

import type { LoadedFile } from '../../formats/types'
import { ViewerProvider } from '../shared/ViewerContext'
import { CodeViewer } from '../CodeViewer'
import { isRunnable } from '../code/useCodeRun'

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

const FIXTURES_DIR = path.resolve(process.cwd(), 'src/viewers/__fixtures__')

function codeFile(content: string, filePath = '/tmp/sample.ts'): LoadedFile {
  return { kind: 'text', content, path: filePath, format: 'code' }
}

function renderViewer(file: LoadedFile) {
  return render(
    <ViewerProvider filePath={file.path}>
      <CodeViewer file={file} />
    </ViewerProvider>,
  )
}

beforeEach(() => {
  window.localStorage.clear()
  window.electronAPI = {
    saveFile: vi.fn().mockResolvedValue({ saved: true, path: '/tmp/sample.ts' }),
  } as unknown as typeof window.electronAPI
})

describe('CodeViewer — editor (USR-18)', () => {
  it('renders a real .ts fixture in the editor with its language and line count', async () => {
    const content = readFileSync(path.join(FIXTURES_DIR, 'sample.ts'), 'utf8')
    const { container } = renderViewer(codeFile(content))

    await waitFor(() => expect(container.querySelector('.cm-content')).not.toBeNull())
    expect(container.querySelector('.cm-content')?.textContent).toContain('atlasFixture')
    expect(screen.getByText('typescript')).toBeInTheDocument()
    // Line numbers, not a plain <pre>.
    expect(container.querySelector('.cm-gutters')).not.toBeNull()
  })

  it('saves the edited document through the shared save contract', async () => {
    const { container } = renderViewer(codeFile('const a = 1\n'))
    await waitFor(() => expect(container.querySelector('.cm-content')).not.toBeNull())

    fireEvent.click(screen.getByRole('button', { name: 'Save' }))

    await waitFor(() => expect(window.electronAPI?.saveFile).toHaveBeenCalled())
    const request = vi.mocked(window.electronAPI!.saveFile!).mock.calls[0][0]
    expect(request).toMatchObject({ content: 'const a = 1\n', existingPath: '/tmp/sample.ts' })
  })

  it('remembers the word wrap preference', async () => {
    const { container, unmount } = renderViewer(codeFile('const a = 1\n'))
    await waitFor(() => expect(container.querySelector('.cm-content')).not.toBeNull())

    const wrapButton = screen.getByRole('button', { name: 'Word wrap' })
    expect(wrapButton).toHaveAttribute('aria-pressed', 'false')
    fireEvent.click(wrapButton)
    expect(wrapButton).toHaveAttribute('aria-pressed', 'true')
    unmount()

    renderViewer(codeFile('const a = 1\n'))
    await waitFor(() => expect(screen.getByRole('button', { name: 'Word wrap' })).toHaveAttribute('aria-pressed', 'true'))
  })
})

describe('CodeViewer — run (USR-19)', () => {
  it('knows which files can be run', () => {
    expect(isRunnable('/tmp/a.js')).toBe(true)
    expect(isRunnable('/tmp/a.PY')).toBe(true)
    expect(isRunnable('/tmp/a.ts')).toBe(true)
    expect(isRunnable('/tmp/a.rs')).toBe(false)
    expect(isRunnable('/tmp/Makefile')).toBe(false)
  })

  it('offers Run only when the main process exposes it', async () => {
    const { container, unmount } = renderViewer(codeFile('print(1)\n', '/tmp/a.py'))
    await waitFor(() => expect(container.querySelector('.cm-content')).not.toBeNull())
    expect(screen.queryByRole('button', { name: 'Run' })).toBeNull()
    unmount()

    window.electronAPI = {
      ...window.electronAPI,
      codeRun: { start: vi.fn(), stop: vi.fn(), onOutput: () => () => {}, onExit: () => () => {} },
    } as unknown as typeof window.electronAPI

    renderViewer(codeFile('print(1)\n', '/tmp/a.py'))
    await waitFor(() => expect(screen.getByRole('button', { name: 'Run' })).toBeInTheDocument())
  })

  it('saves before running, and streams the program output into the panel', async () => {
    const listeners: { output?: (p: unknown) => void; exit?: (p: unknown) => void } = {}
    const start = vi.fn().mockResolvedValue({ ok: true, runId: 3 })
    window.electronAPI = {
      saveFile: vi.fn().mockResolvedValue({ saved: true, path: '/tmp/a.js' }),
      codeRun: {
        start,
        stop: vi.fn(),
        onOutput: (callback: (p: unknown) => void) => {
          listeners.output = callback
          return () => {}
        },
        onExit: (callback: (p: unknown) => void) => {
          listeners.exit = callback
          return () => {}
        },
      },
    } as unknown as typeof window.electronAPI

    const { container } = renderViewer(codeFile('console.log(1)\n', '/tmp/a.js'))
    await waitFor(() => expect(container.querySelector('.cm-content')).not.toBeNull())

    fireEvent.click(screen.getByRole('button', { name: 'Run' }))
    await waitFor(() => expect(start).toHaveBeenCalledWith('/tmp/a.js'))
    expect(window.electronAPI?.saveFile).toHaveBeenCalled()
    await screen.findByRole('button', { name: 'Stop' })

    listeners.output?.({ runId: 3, stream: 'stdout', text: 'hello world' })
    await screen.findByText('hello world')

    listeners.exit?.({ runId: 3, code: 0, timedOut: false, stopped: false })
    await screen.findByText('Finished (exit code 0).')
  })

  // Found by driving the real app: switching to another tab (or closing this
  // one) while a program was running unmounted CodeViewer along with the
  // only Stop button that could ever reach that run. Main allows exactly one
  // run at a time, so the child kept running orphaned with no way to stop it
  // short of its 60s timeout or quitting the whole app, and every other
  // file's Run just failed with "a program is already running".
  it('stops the run when the viewer is closed mid-run, instead of leaving it orphaned', async () => {
    const stop = vi.fn()
    const start = vi.fn().mockResolvedValue({ ok: true, runId: 9 })
    window.electronAPI = {
      saveFile: vi.fn().mockResolvedValue({ saved: true, path: '/tmp/a.js' }),
      codeRun: { start, stop, onOutput: () => () => {}, onExit: () => () => {} },
    } as unknown as typeof window.electronAPI

    const { container, unmount } = renderViewer(codeFile('console.log(1)\n', '/tmp/a.js'))
    await waitFor(() => expect(container.querySelector('.cm-content')).not.toBeNull())

    fireEvent.click(screen.getByRole('button', { name: 'Run' }))
    await waitFor(() => expect(start).toHaveBeenCalledWith('/tmp/a.js'))
    await screen.findByRole('button', { name: 'Stop' })

    expect(stop).not.toHaveBeenCalled()
    unmount()
    expect(stop).toHaveBeenCalledWith(9)
  })
})
