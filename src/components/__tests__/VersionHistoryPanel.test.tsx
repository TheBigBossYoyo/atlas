/**
 * VERSIONS-1 — the version history panel.
 *
 * The behaviour worth pinning is mostly about what the panel REFUSES to do: it
 * must not offer to restore over unsaved work, and it must not throw a history
 * away on one click. Both are one-way losses of something the user cannot get
 * back, which is exactly the kind of thing a later refactor quietly relaxes.
 */
import { fireEvent, render, screen } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'

import type { DocumentVersion } from '../../electron'
import type { VersionHistoryState } from '../../hooks/useVersionHistory'
import { VersionHistoryPanel } from '../VersionHistoryPanel'

function version(overrides: Partial<DocumentVersion> = {}): DocumentVersion {
  return { id: 'a'.repeat(64), at: Date.now() - 60_000, bytes: 2048, label: null, ...overrides }
}

function makeHistory(overrides: Partial<VersionHistoryState> = {}): VersionHistoryState {
  return {
    available: true,
    versions: [version()],
    loading: false,
    error: null,
    refresh: vi.fn(async () => {}),
    captureNow: vi.fn(async () => true),
    readVersion: vi.fn(async () => new Uint8Array()),
    clear: vi.fn(async () => {}),
    ...overrides,
  }
}

function renderPanel(overrides: Partial<React.ComponentProps<typeof VersionHistoryPanel>> = {}) {
  const props = {
    open: true,
    onClose: vi.fn(),
    history: makeHistory(),
    onRestore: vi.fn(async () => {}),
    documentDirty: false,
    onCaptureNow: vi.fn(async () => {}),
    ...overrides,
  }
  render(<VersionHistoryPanel {...props} />)
  return props
}

describe('VersionHistoryPanel (VERSIONS-1)', () => {
  it('renders nothing when closed', () => {
    renderPanel({ open: false })
    expect(screen.queryByRole('complementary')).not.toBeInTheDocument()
  })

  it('lists a version with when it was taken and how big it was', () => {
    renderPanel()
    // The shared formatter renders "1m ago" — terse on purpose, matching the
            // recent-files list it was extracted from.
    expect(screen.getByText('1m ago')).toBeInTheDocument()
    expect(screen.getByText('2 KB')).toBeInTheDocument()
  })

  it('marks only the newest version as the latest', () => {
    renderPanel({
      history: makeHistory({
        versions: [version({ id: 'a'.repeat(64) }), version({ id: 'b'.repeat(64), at: Date.now() - 500_000 })],
      }),
    })
    expect(screen.getAllByText('Latest')).toHaveLength(1)
  })

  it('restores the version whose button was pressed', async () => {
    const newest = version({ id: 'a'.repeat(64) })
    const older = version({ id: 'b'.repeat(64), at: Date.now() - 500_000 })
    const props = renderPanel({ history: makeHistory({ versions: [newest, older] }) })

    fireEvent.click(screen.getAllByRole('button', { name: /Restore/ })[1]!)

    expect(props.onRestore).toHaveBeenCalledWith(older)
  })

  describe('with unsaved changes', () => {
    it('refuses to restore, and says why', () => {
      // Restoring writes over the document; main records what it writes, so the
      // state being replaced survives only if it was already written. Offering
      // this over unsaved edits would destroy work that exists nowhere else.
      renderPanel({ documentDirty: true })

      expect(screen.getByRole('button', { name: /Restore/ })).toBeDisabled()
      expect(screen.getByText(/Save your changes before restoring/i)).toBeInTheDocument()
    })
  })

  it('asks before throwing a history away', () => {
    const history = makeHistory()
    renderPanel({ history })

    fireEvent.click(screen.getByRole('button', { name: /Delete all versions/i }))
    // Still nothing deleted: the first click only asks.
    expect(history.clear).not.toHaveBeenCalled()

    fireEvent.click(screen.getByRole('button', { name: 'Delete' }))
    expect(history.clear).toHaveBeenCalledTimes(1)
  })

  it('lets the user back out of deleting', () => {
    const history = makeHistory()
    renderPanel({ history })

    fireEvent.click(screen.getByRole('button', { name: /Delete all versions/i }))
    fireEvent.click(screen.getByRole('button', { name: /Keep them/i }))

    expect(history.clear).not.toHaveBeenCalled()
    expect(screen.getByRole('button', { name: /Delete all versions/i })).toBeInTheDocument()
  })

  it('explains an empty history rather than showing a blank panel', () => {
    renderPanel({ history: makeHistory({ versions: [] }) })
    expect(screen.getByText(/No versions yet/i)).toBeInTheDocument()
  })

  it('surfaces an error from the store', () => {
    renderPanel({ history: makeHistory({ error: 'the disk is full' }) })
    expect(screen.getByText('the disk is full')).toBeInTheDocument()
  })

  it('offers a manual capture only where one is possible', () => {
    renderPanel()
    expect(screen.getByRole('button', { name: /Save a version now/i })).toBeInTheDocument()

    // VERSIONS-2 — every editable format can be captured now (each viewer
    // registers its own serializer), but a format with no editor at all — a PDF,
    // an image — still cannot. There the action is absent rather than present
    // and broken.
    renderPanel({ onCaptureNow: null })
    expect(screen.queryAllByRole('button', { name: /Save a version now/i })).toHaveLength(1)
  })

  it('shows a version label when one was recorded', () => {
    renderPanel({ history: makeHistory({ versions: [version({ label: 'before restoring an earlier draft' })] }) })
    expect(screen.getByText('before restoring an earlier draft')).toBeInTheDocument()
  })

  it('closes on the close button', () => {
    const props = renderPanel()
    fireEvent.click(screen.getByRole('button', { name: /Close version history/i }))
    expect(props.onClose).toHaveBeenCalled()
  })
})
