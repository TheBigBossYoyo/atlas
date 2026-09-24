import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication } from '@playwright/test'

import { openMarkdownEditorWithShortcut } from './helpers/markdownViewMode'

/**
 * SHELL-6 — found by the night sweep in the real app: after a crash,
 * relaunching Atlas straight into the file that was being edited (a
 * double-click / "Open with", the usual way back) never offered the draft,
 * and autosave overwrote the draft with the file's on-disk content ~800ms
 * later. Only a relaunch with no file showed the prompt.
 *
 * Two launches share a profile via `ATLAS_E2E_PROFILE_DIR` (see
 * recent-files.spec.ts for why the kill must be awaited, and why storage is
 * flushed before it).
 */

const projectRoot = process.cwd()

async function killAndWait(app: ElectronApplication, timeoutMs = 10_000): Promise<void> {
  const proc = app.process()
  if (proc.exitCode !== null || proc.signalCode !== null) return
  const exited = new Promise<void>((resolve) => proc.once('exit', () => resolve()))
  try {
    execSync(`taskkill /PID ${proc.pid} /T /F`, { stdio: 'ignore' })
  } catch {
    try {
      proc.kill()
    } catch {
      // already gone
    }
  }
  await Promise.race([exited, new Promise<void>((resolve) => setTimeout(resolve, timeoutMs))])
}

function launch(profileDir: string, file: string): Promise<ElectronApplication> {
  return electron.launch({
    args: ['.', file],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1', ATLAS_HIDDEN_WINDOW: '1', ATLAS_E2E_PROFILE_DIR: profileDir },
  })
}

test('SHELL-6: relaunching straight into the crashed file offers its draft, and Restore brings the edits back', async () => {
  const profileDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-e2e-draft-profile-'))
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-e2e-draft-')), 'notes.md')
  fs.writeFileSync(file, '# Notes\n\nOn disk.\n')

  // Session 1: edit without saving, let autosave write the draft, then "crash".
  const first = await launch(profileDir, file)
  try {
    const page = await first.firstWindow()
    await expect(page.locator('[data-viewer="markdown"]')).toHaveCount(1, { timeout: 15_000 })
    await openMarkdownEditorWithShortcut(page)
    const textarea = page.locator('.editor-panel__textarea')
    await textarea.click()
    await page.keyboard.press('Control+End')
    await page.keyboard.type('Unsaved edit before the crash.')
    await expect
      .poll(() => page.evaluate(() => localStorage.getItem('atlas-draft')), { timeout: 5_000 })
      .toContain('Unsaved edit before the crash.')
    await first.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].webContents.session.flushStorageData())
  } finally {
    await killAndWait(first)
  }
  expect(fs.readFileSync(file, 'utf-8'), 'the crash must not have saved anything').toBe('# Notes\n\nOn disk.\n')

  // Session 2: relaunch INTO the same file.
  const second = await launch(profileDir, file)
  try {
    const page = await second.firstWindow()
    await expect(page.locator('[data-viewer="markdown"]')).toHaveCount(1, { timeout: 15_000 })
    // Past the 800ms autosave debounce: the draft must still be the crashed one.
    await page.waitForTimeout(1500)
    expect(await page.evaluate(() => localStorage.getItem('atlas-draft'))).toContain('Unsaved edit before the crash.')

    await expect(page.getByText(/unsaved changes to "notes\.md"/i)).toBeVisible({ timeout: 5_000 })
    await page.getByRole('button', { name: 'Restore' }).click()
    await openMarkdownEditorWithShortcut(page)
    await expect(page.locator('.editor-panel__textarea')).toHaveValue(/Unsaved edit before the crash\./)

    // Restored into THIS file: saving writes the recovered text to it.
    await page.keyboard.press('Control+s')
    await expect.poll(() => fs.readFileSync(file, 'utf-8'), { timeout: 10_000 }).toContain('Unsaved edit before the crash.')
  } finally {
    await killAndWait(second)
  }
})
