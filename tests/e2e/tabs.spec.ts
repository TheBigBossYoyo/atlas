import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication } from '@playwright/test'

/** SHELL-17 — several documents open at once. */

const projectRoot = process.cwd()

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

test('opens documents in tabs, switches between them and closes them', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-tabs-'))
  const first = path.join(dir, 'first.md')
  const second = path.join(dir, 'second.md')
  fs.writeFileSync(first, '# First document\n\nHello from the first file.\n')
  fs.writeFileSync(second, '# Second document\n\nHello from the second file.\n')

  const app = await electron.launch({
    args: ['.', first],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  try {
    const page = await app.firstWindow()
    await expect(page.getByRole('tab', { name: 'first.md' })).toBeVisible({ timeout: 20_000 })

    // Opening a second file adds a tab and shows it — through the real Open
    // flow (the dialog is answered in the main process), so the path goes
    // through the same allowlisting a user's own open would.
    await app.evaluate(async ({ dialog }, filePath) => {
      dialog.showOpenDialog = (async () => ({ canceled: false, filePaths: [filePath] })) as typeof dialog.showOpenDialog
    }, second)
    await page.getByRole('button', { name: 'Open', exact: true }).click()
    const secondTab = page.getByRole('tab', { name: 'second.md' })
    await expect(secondTab).toBeVisible({ timeout: 20_000 })
    await expect(secondTab).toHaveAttribute('aria-selected', 'true')
    await expect(page.locator('.markdown-body')).toContainText('second file')

    // Switching back shows the first document again, without re-reading it.
    await page.getByRole('tab', { name: 'first.md' }).click()
    await expect(page.locator('.markdown-body')).toContainText('first file')
    await expect(page.getByRole('tab', { name: 'first.md' })).toHaveAttribute('aria-selected', 'true')

    // Ctrl+Tab walks to the other document.
    await page.keyboard.press('Control+Tab')
    await expect(page.getByRole('tab', { name: 'second.md' })).toHaveAttribute('aria-selected', 'true')

    // Closing a tab leaves the other one showing.
    await page.getByRole('button', { name: 'Close second.md' }).click()
    await expect(page.getByRole('tab', { name: 'second.md' })).toHaveCount(0)
    await expect(page.getByRole('tab', { name: 'first.md' })).toHaveAttribute('aria-selected', 'true')
    await expect(page.locator('.markdown-body')).toContainText('first file')

    // Ctrl+Shift+T brings the closed one back.
    await page.keyboard.press('Control+Shift+T')
    await expect(page.getByRole('tab', { name: 'second.md' })).toBeVisible()
    await expect(page.locator('.markdown-body')).toContainText('second file')
  } finally {
    kill(app)
  }
})
