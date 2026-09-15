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
    expect(screen.queryByText(/showing the first/i)).not.toBeInTheDocument()

    fireEvent.click(screen.getByRole('button', { name: /back/i }))
    expect(screen.getByText('README')).toBeInTheDocument()
  })

  it('caps decoded bytes for an oversized file and shows a truncation notice (wave-3 shell-polish follow-up)', () => {
    const TEXT_PREVIEW_MAX_BYTES = 5 * 1024 * 1024
    // One byte over the cap — an "a" repeated, followed by a single
    // recognizable marker character placed right at the cap boundary so the
    // test can assert it is NOT present in the decoded preview.
    const bytes = new Uint8Array(TEXT_PREVIEW_MAX_BYTES + 1).fill(0x61) // 'a'
    bytes[TEXT_PREVIEW_MAX_BYTES] = 0x5a // 'Z' — one byte past the cap

    render(<UnknownViewer file={makeBinaryFile(Array.from(bytes), 'C:/docs/huge.bin')} />)

    fireEvent.click(screen.getByRole('button', { name: /open as text/i }))

    expect(screen.getByText(/showing the first 5(\.0)? MB only/i)).toBeInTheDocument()
    // The marker byte just past the cap must never have been decoded.
    expect(screen.queryByText(/Z/)).not.toBeInTheDocument()
  })

  it('does not show a truncation notice for a file at or under the cap', () => {
    const bytes = new Uint8Array(1024).fill(0x61)
    render(<UnknownViewer file={makeBinaryFile(Array.from(bytes), 'C:/docs/small.bin')} />)

    fireEvent.click(screen.getByRole('button', { name: /open as text/i }))

    expect(screen.queryByText(/showing the first/i)).not.toBeInTheDocument()
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
