/**
 * UnknownViewer — P2.11/LOAD-18/LOAD-11
 *
 * Was a bare `<span>{format}</span>` + path with zero actionable UX. Now a
 * real empty state (icon, file name/size, explanation, actions) that also
 * gives legacy binary Office files (.doc/.xls/.ppt) a specific message.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { render, screen, fireEvent, waitFor } from '@testing-library/react'

import { UnknownViewer } from '../UnknownViewer'
import type { LoadedFile } from '../../formats/types'

function makeBinaryFile(bytes: ReadonlyArray<number>, path = 'C:/Users/test/mystery.bin'): LoadedFile {
  return {
    kind: 'binary',
    format: 'unknown',
    path,
    content: new Uint8Array(bytes).buffer,
  }
}

function cfbBytes(extra: ReadonlyArray<number> = []): number[] {
  return [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, ...extra]
}

afterEach(() => {
  vi.restoreAllMocks()
  delete (window as { electronAPI?: unknown }).electronAPI
})

describe('UnknownViewer', () => {
  it('shows the file name and a formatted size for a genuinely unknown binary', () => {
    render(<UnknownViewer file={makeBinaryFile([1, 2, 3, 4, 5], 'C:/docs/weird.xyz')} />)

    expect(screen.getByText('weird.xyz')).toBeInTheDocument()
    expect(screen.getByText('5 B')).toBeInTheDocument()
    expect(screen.getByText(/doesn't recognize this file's format/i)).toBeInTheDocument()
  })

  it('gives a specific message for a legacy .doc (CFB magic) file', () => {
    render(<UnknownViewer file={makeBinaryFile(cfbBytes(), 'C:/docs/old-report.doc')} />)

    expect(screen.getByText(/Word 97-2003 document/i)).toBeInTheDocument()
    expect(screen.getByText(/\.docx/)).toBeInTheDocument()
  })

  it('gives a specific message for a legacy .xls (CFB magic) file', () => {
    render(<UnknownViewer file={makeBinaryFile(cfbBytes(), 'C:/docs/budget.xls')} />)

    expect(screen.getByText(/Excel 97-2003 workbook/i)).toBeInTheDocument()
  })

  it('toggles to a decoded text view and back', () => {
    const text = 'Hello from an extensionless file'
    const bytes = Array.from(new TextEncoder().encode(text))
    render(<UnknownViewer file={makeBinaryFile(bytes, 'C:/docs/README')} />)

    fireEvent.click(screen.getByRole('button', { name: /open as text/i }))
    expect(screen.getByText(text)).toBeInTheDocument()

    fireEvent.click(screen.getByRole('button', { name: /back/i }))
    expect(screen.getByText('README')).toBeInTheDocument()
  })

  it('reveal in folder calls the IPC bridge with the file path', async () => {
    const revealInFolder = vi.fn().mockResolvedValue({ ok: true })
    window.electronAPI = { revealInFolder } as unknown as typeof window.electronAPI

    render(<UnknownViewer file={makeBinaryFile([1, 2, 3], 'C:/docs/weird.xyz')} />)
    fireEvent.click(screen.getByRole('button', { name: /reveal in folder/i }))

    await waitFor(() => {
      expect(revealInFolder).toHaveBeenCalledWith('C:/docs/weird.xyz')
    })
  })

  it('shows a friendly error when reveal in folder fails', async () => {
    const revealInFolder = vi.fn().mockResolvedValue({ ok: false })
    window.electronAPI = { revealInFolder } as unknown as typeof window.electronAPI

    render(<UnknownViewer file={makeBinaryFile([1, 2, 3], 'C:/docs/weird.xyz')} />)
    fireEvent.click(screen.getByRole('button', { name: /reveal in folder/i }))

    await waitFor(() => {
      expect(screen.getByRole('alert')).toHaveTextContent(/could not reveal/i)
    })
  })

  it('shows a friendly error when reveal in folder is used outside the desktop app', async () => {
    render(<UnknownViewer file={makeBinaryFile([1, 2, 3], 'C:/docs/weird.xyz')} />)
    fireEvent.click(screen.getByRole('button', { name: /reveal in folder/i }))

    await waitFor(() => {
      expect(screen.getByRole('alert')).toHaveTextContent(/desktop app/i)
    })
  })
})
