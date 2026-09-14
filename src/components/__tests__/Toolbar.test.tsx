/**
 * P2.7 (SHELL-14 / SHELL-15 / UX-07) regression test.
 *
 * The font-size button group and the global Save button used to be always
 * rendered — disabled or silently inert — for every non-markdown format,
 * which is misleading (SHELL-14/UX-07: font-size controls that visibly do
 * nothing) and confusing (SHELL-15: a permanently-disabled Save button next
 * to DocxViewer's own working one). Both must be hidden outright rather than
 * shown-but-inert, driven by dedicated props (`canChangeFontSize`, `canSave`)
 * so a capability-gated format (e.g. DOCX via a future flag) can still show
 * them by flipping those props.
 */
import type { ComponentProps } from 'react'
import { render, screen } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'

import { Toolbar } from '../Toolbar'
import { THEMES } from '../../types'

function noop() {
  /* no-op */
}

type ToolbarOverrides = Partial<ComponentProps<typeof Toolbar>>

function renderToolbar(overrides: ToolbarOverrides = {}) {
  return render(
    <Toolbar
      theme="light"
      themes={THEMES}
      viewMode="preview"
      sidebarOpen={true}
      fileName="example.txt"
      isDirty={false}
      hasContent={true}
      canSave={false}
      canSearch={false}
      canChangeFontSize={false}
      isMarkdown={false}
      isElectron={false}
      exportFormat="text"
      exportMenuOpen={false}
      canExportCsv={false}
      onSelectTheme={noop}
      onViewModeChange={noop}
      onToggleSidebar={noop}
      onOpenFile={noop}
      onSave={vi.fn()}
      onExport={noop}
      onOpenSearch={noop}
      onShowShortcuts={noop}
      onIncreaseFont={noop}
      onDecreaseFont={noop}
      onResetFont={noop}
      onExportMenuOpenChange={noop}
      {...overrides}
    />,
  )
}

describe('Toolbar format-gated chrome (SHELL-14 / SHELL-15 / UX-07)', () => {
  it('hides the font-size group when canChangeFontSize is false', () => {
    renderToolbar({ canChangeFontSize: false })
    expect(screen.queryByRole('group', { name: 'Font size' })).not.toBeInTheDocument()
  })

  it('shows the font-size group when canChangeFontSize is true', () => {
    renderToolbar({ canChangeFontSize: true })
    expect(screen.getByRole('group', { name: 'Font size' })).toBeInTheDocument()
  })

  it('hides the global Save button when canSave is false', () => {
    renderToolbar({ canSave: false })
    expect(screen.queryByTitle('Save (Ctrl+S)')).not.toBeInTheDocument()
  })

  it('shows an enabled global Save button when canSave is true', () => {
    renderToolbar({ canSave: true })
    const saveButton = screen.getByTitle('Save (Ctrl+S)')
    expect(saveButton).toBeInTheDocument()
    expect(saveButton).not.toBeDisabled()
  })
})
