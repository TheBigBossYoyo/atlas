import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execFileSync } from 'node:child_process'

import { _electron as electron, expect, test } from '@playwright/test'
import type { ElectronApplication, Page } from '@playwright/test'

import { openMarkdownEditorWithShortcut } from './helpers/markdownViewMode'

/**
 * P4.6 (part) / P2.1 (SHELL-08, SHELL-09, UX-03) — shortcut-collision
 * regression coverage against the real app, on top of the inline unit-level
 * regression test P2.1 already gates its own merge on and the wave-3
 * shell-polish follow-up's own dispatcher-registration unit tests
 * (usePdfFind, useGridFind, useSlideKeyboardNav). Those exercise the
 * dispatcher's precedence logic directly; this exercises the SAME keys
 * through a real Electron renderer end to end, across every format the
 * dispatcher now covers (DOCX, markdown, PDF, spreadsheet, slides) —
 * catching anything a jsdom-based unit test's simulated event dispatch
 * could miss.
 *
 * Ctrl+B/E/P/S/F/W are the exact combos SHELL-08/SHELL-09/UX-03 originally
 * found racing between a viewer's own editing shortcuts and the shell's
 * global sidebar-toggle/export-menu/print/save/close-file/find actions.
 */

const projectRoot = process.cwd()
const fixtureDir = path.join(projectRoot, 'tests', 'e2e', 'fixtures')

test.beforeAll(() => {
  execFileSync(process.execPath, ['tests/e2e/fixtures/generate.mjs'], {
    cwd: projectRoot,
    stdio: 'inherit',
  })
})

async function launch(filePath: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({
    args: ['.', filePath],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  const page = await app.firstWindow()
  return { app, page }
}

/** Copies a fixture into a scratch temp dir so a test that exercises a real
 * Ctrl+S save (DocxViewer's own save always writes, with no dirty-check
 * short-circuit) overwrites a throwaway file instead of the git-tracked
 * fixture. */
function copyFixtureToTemp(fileName: string): string {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-e2e-shortcut-'))
  const dest = path.join(tempDir, fileName)
  fs.copyFileSync(path.join(fixtureDir, fileName), dest)
  return dest
}

test('DOCX editor: Ctrl+B/Ctrl+E/Ctrl+S do not leak to the shell (sidebar/export menu), Ctrl+F opens the document find bar, and Ctrl+W (once the document is clean) closes it', async () => {
  const tempPath = copyFixtureToTemp('sample.docx')
  const { app, page } = await launch(tempPath)
  try {
    const editor = page.getByRole('textbox', { name: 'Document editor' })
    await expect(editor).toBeVisible({ timeout: 15_000 })
    await editor.click()

    const exportButton = page.getByRole('button', { name: 'Export', exact: true })
    const sidebar = page.locator('.sidebar')

    await expect(sidebar).toHaveCount(1)
    await expect(exportButton).toHaveAttribute('aria-expanded', 'false')

    // Ctrl+B (bold) is reserved by DocxViewer's own editor — must not also
    // toggle the shell's sidebar (SHELL-08).
    await page.keyboard.press('Control+b')
    await expect(sidebar).toHaveCount(1)
    await expect(exportButton).toHaveAttribute('aria-expanded', 'false')

    // Ctrl+E is reserved for center-alignment — must not also open the
    // shell's Export menu (SHELL-09).
    await page.keyboard.press('Control+e')
    await expect(exportButton).toHaveAttribute('aria-expanded', 'false')
    await expect(page.getByLabel('Export options')).toHaveCount(0)

    // Ctrl+S saves via the viewer's own registered save — must not also pop
    // the shell's Save As / export flow.
    await page.keyboard.press('Control+s')
    await expect(page.getByLabel('Export options')).toHaveCount(0)

    // Ctrl+F opens DocxViewer's own find bar, not markdown's shared
    // SearchOverlay (a different component entirely).
    await page.keyboard.press('Control+f')
    await expect(page.locator('.docx-find__input[aria-label="Find"]')).toBeVisible({ timeout: 5_000 })
    await expect(page.locator('.search-overlay')).toHaveCount(0)
    await page.keyboard.press('Escape')
    // FindReplace stays mounted (its own `open` prop only toggles the
    // `docx-find--closed` CSS class, matching FindReplace.test.tsx's own
    // assertions) rather than unmounting — assert hidden, not absent.
    await expect(page.locator('.docx-find__input[aria-label="Find"]')).not.toBeVisible()

    // Ctrl+W (close file) is NOT one of DocxViewer's reserved combos, so it
    // falls through to the shell. FIELD-01 — the shell no longer skips Ctrl+W
    // while focus is in a plain-text/contentEditable field (it's carved out
    // of `ctx.inPlainField` the same way Ctrl+S/Ctrl+P already were: a
    // shell/window-level convention, not a field-editing command), so the
    // contentEditable document surface being focused here does not block it.
    // The document was just saved via Ctrl+S above, so it's clean — closeFile
    // finds nothing to confirm and closes straight through to the Welcome
    // screen with no unsaved-changes prompt.
    await editor.click()
    await page.keyboard.press('Control+w')
    await expect(page.locator('.confirm-dialog')).toHaveCount(0)
    await expect(page.getByRole('button', { name: 'Load Sample Document' })).toBeVisible({ timeout: 10_000 })
    await expect(page.locator('[data-viewer="docx"]')).toHaveCount(0)
  } finally {
    await app.close()
    fs.rmSync(path.dirname(tempPath), { recursive: true, force: true })
  }
})

test('Markdown editor (RawEditor): Ctrl+B does not toggle the sidebar while typing, and Ctrl+F opens the search overlay', async () => {
  const { app, page } = await launch(path.join(fixtureDir, 'sample.md'))
  try {
    await expect(page.locator('[data-viewer="markdown"]')).toHaveCount(1)

    // Switch to a view mode that mounts RawEditor (a plain <textarea> with
    // no key handling of its own — SHELL-08's fix must hold even though
    // this component makes zero effort to claim the combo itself).
    await openMarkdownEditorWithShortcut(page)
    const textarea = page.locator('.editor-panel__textarea')
    await textarea.click()

    const sidebar = page.locator('.sidebar')
    await expect(sidebar).toHaveCount(1)

    await page.keyboard.press('Control+b')
    await expect(sidebar).toHaveCount(1)
    await expect(page.getByLabel('Export options')).toHaveCount(0)

    await page.keyboard.press('Control+f')
    await expect(page.locator('.search-overlay')).toBeVisible({ timeout: 5_000 })
    await page.keyboard.press('Escape')
  } finally {
    await app.close()
  }
})

test('PDF: Ctrl+F opens the PDF find bar (not the shared search overlay); Ctrl+B, which PDF does not reserve, still toggles the shell sidebar', async () => {
  const { app, page } = await launch(path.join(fixtureDir, 'sample.pdf'))
  try {
    await expect(page.locator('[data-viewer="pdf"]')).toHaveCount(1)
    await expect(page.locator('.pdf-viewer__page-wrapper')).toHaveCount(1, { timeout: 15_000 })

    await page.keyboard.press('Control+f')
    await expect(page.getByPlaceholder('Find in document…')).toBeVisible({ timeout: 5_000 })
    await expect(page.locator('.search-overlay')).toHaveCount(0)
    await page.keyboard.press('Escape')
    // Wait for the find bar to actually close: Ctrl+B pressed while its input
    // still has focus goes to the input, not the shell (flaky on slow CI).
    await expect(page.getByPlaceholder('Find in document…')).toHaveCount(0, { timeout: 5_000 })

    // PDF's own shortcut set (zoom/print/page-nav/find) never reserves
    // Ctrl+B — confirms the new viewer-tier find registration didn't
    // accidentally start swallowing OTHER shell shortcuts too.
    // Shell shortcuts deliberately stand back while a text input has focus
    // (focused-input -> active-viewer -> shell-global). After the find bar
    // closes, focus can still sit on one of the PDF viewer's own controls on
    // a slow runner, which swallowed this Ctrl+B in CI run 35469314616.
    await page.evaluate(() => {
      ;(document.activeElement as HTMLElement | null)?.blur()
      // Record the keydown as the window sees it, so a failure here says
      // whether the key never arrived or the shell ignored it (this assertion
      // failed twice on CI and never locally).
      const win = window as unknown as { __ctrlB?: string[] }
      win.__ctrlB = []
      window.addEventListener(
        'keydown',
        (event) => {
          if (event.key === 'b') win.__ctrlB!.push(`ctrl=${event.ctrlKey} target=${(event.target as Element)?.className || (event.target as Node)?.nodeName}`)
        },
        true,
      )
    })

    const sidebar = page.locator('.sidebar')
    const wasOpen = (await sidebar.count()) === 1
    await page.keyboard.press('Control+b')
    const seen = await page.evaluate(() => (window as unknown as { __ctrlB?: string[] }).__ctrlB ?? [])
    expect(seen, 'the Ctrl+B keydown never reached the window').not.toEqual([])
    await expect(sidebar, `Ctrl+B reached the window (${seen.join('; ')}) but the sidebar did not toggle`).toHaveCount(
      wasOpen ? 0 : 1,
      { timeout: 10_000 },
    )
  } finally {
    await app.close()
  }
})

test('Spreadsheet: Ctrl+F opens the grid find bar and finds a real cell match', async () => {
  const { app, page } = await launch(path.join(fixtureDir, 'sample.xlsx'))
  try {
    await expect(page.locator('[data-viewer="xlsx"]')).toHaveCount(1, { timeout: 15_000 })

    await page.keyboard.press('Control+f')
    const searchOverlay = page.locator('.search-overlay')
    await expect(searchOverlay).toBeVisible({ timeout: 5_000 })

    await searchOverlay.locator('.search-overlay__input').fill('Atlas')
    await expect(searchOverlay.locator('.search-overlay__count')).toHaveText('1 of 1', { timeout: 5_000 })

    await page.keyboard.press('Escape')
    await expect(searchOverlay).toHaveCount(0)
  } finally {
    await app.close()
  }
})

test('Slides: Space on a focused thumbnail activates the button (not "next slide"); Ctrl+F is a safe no-op (not searchable)', async () => {
  const { app, page } = await launch(path.join(fixtureDir, 'sample.pptx'))
  try {
    await expect(page.locator('[data-viewer="pptx"]')).toHaveCount(1, { timeout: 15_000 })

    const thumbnails = page.locator('.slide-deck__thumb')
    await expect(thumbnails).toHaveCount(2, { timeout: 15_000 })

    // Move the active slide to index 1 via an ordinary click (unrelated to
    // Space handling), then focus the FIRST thumbnail and press Space. The
    // pre-fix bug advanced the keyboard-nav hook's own index instead of
    // letting the focused button handle Space itself — with the active
    // index already at the last slide (1), that bug would clamp right back
    // to 1 (no visible change), indistinguishable from doing nothing. The
    // fix instead lets the focused button's native activation fire,
    // selecting slide 1 (index 0) — a clearly different, and correct,
    // result.
    await thumbnails.nth(1).click()
    await expect(thumbnails.nth(1)).toHaveClass(/slide-deck__thumb--active/)

    await thumbnails.nth(0).focus()
    await page.keyboard.press(' ')
    await expect(thumbnails.nth(0)).toHaveClass(/slide-deck__thumb--active/)
    await expect(thumbnails.nth(1)).not.toHaveClass(/slide-deck__thumb--active/)

    const consoleErrors: string[] = []
    page.on('console', (message) => {
      if (message.type() === 'error') consoleErrors.push(message.text())
    })

    // Slides are excluded from the TreeWalker-based search (canvas-shaped
    // content) — Ctrl+F must never be swallowed with no feature behind it,
    // and must not crash or open any find UI.
    await page.keyboard.press('Control+f')
    await page.waitForTimeout(300)
    await expect(page.locator('.search-overlay')).toHaveCount(0)
    await expect(page.locator('.pdf-viewer__find-bar')).toHaveCount(0)
    expect(consoleErrors, consoleErrors.join('\n')).toEqual([])
  } finally {
    await app.close()
  }
})
