/**
 * SHELL-17's invariant ("only the showing document can be dirty" —
 * `App.tsx`'s `selectSession`) is what the close-confirmation machinery
 * relies on to only ever need to save ONE document before quitting. This
 * confirms that invariant actually holds through the real UI: switching away
 * from a dirty tab must always prompt first, so two tabs can never both be
 * dirty when the window (or the whole app) is closed.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication } from '@playwright/test'

const projectRoot = process.cwd()

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

test('switching away from a dirty tab always prompts first — two tabs can never both be dirty at once', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-multi-dirty-'))
  const fileA = path.join(dir, 'A.md')
  const fileB = path.join(dir, 'B.md')
  fs.writeFileSync(fileA, '# A')
  fs.writeFileSync(fileB, '# B')

  const app = await electron.launch({
    args: ['.', fileA],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  const pageErrors: string[] = []
  try {
    const page = await app.firstWindow()
    page.on('pageerror', (e) => pageErrors.push(String(e)))
    await page.waitForTimeout(1000)

    // Open B as a second tab via the native Open dialog (stubbed, same way
    // export.spec.ts/new-document.spec.ts already do for `showSaveDialog`).
    await app.evaluate(({ dialog }, targetPath) => {
      dialog.showOpenDialog = (() => Promise.resolve({ canceled: false, filePaths: [targetPath] })) as typeof dialog.showOpenDialog
    }, fileB)
    await page.getByRole('button', { name: 'Open', exact: true }).click()
    await page.waitForTimeout(1000)

    const tabCount = await page.locator('[role="tab"]').count()
    console.log('tab count after opening B:', tabCount)

    // Dirty tab B (the active one).
    await page.getByRole('button', { name: 'Editor', exact: true }).click()
    await page.waitForTimeout(300)
    const editable = page.locator('[contenteditable="true"], textarea, .cm-content').first()
    await editable.click()
    await page.keyboard.type(' dirty-edit')
    await page.waitForTimeout(300)

    // Try to switch to tab A — must show the unsaved-changes prompt, not switch silently.
    const tabs = page.locator('[role="tab"]')
    const firstTab = tabs.first()
    await firstTab.click()
    await page.waitForTimeout(500)

    const dialogVisible = await page.locator('text=Unsaved changes').first().isVisible().catch(() => false)
    console.log('unsaved-changes dialog visible after trying to switch tabs?', dialogVisible)
    expect(dialogVisible).toBe(true)

    // Discard, to leave the app in a clean state.
    const discardButton = page.getByRole('button', { name: 'Discard', exact: true })
    if (await discardButton.count()) {
      await discardButton.click()
      await page.waitForTimeout(500)
    }

    expect(pageErrors, pageErrors.join('\n')).toEqual([])
  } finally {
    kill(app)
  }
})
