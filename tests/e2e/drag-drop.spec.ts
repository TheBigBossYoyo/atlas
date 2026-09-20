import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'

/**
 * Drag-and-drop onto the window (P1.2/P1.4 — ELEC-01/SHELL-01/LOAD-02/RUN-04/
 * SHELL-24). Real Electron `File` objects backed by an OS path only come
 * from an actual native drag, or from a real `<input type="file">` selection
 * (Playwright's `setInputFiles` populates one with genuinely disk-backed
 * `File`s) — a hand-built `File`/`DataTransfer` in-page has no backing path,
 * so `webUtils.getPathForFile` (which Atlas's drop handler calls through
 * `window.electronAPI.getPathForFile`, exposed — and frozen — by
 * contextBridge) can't resolve it. These tests inject a throwaway
 * `<input type="file" multiple>`, populate it with `setInputFiles`, then
 * build the drop `DataTransfer` from that input's own real `.files` and
 * dispatch a genuine `drop` DragEvent — exercising the app's real
 * `getPathForFile` -> `registerDroppedPath` -> open pipeline with no mocking.
 */

const projectRoot = process.cwd()

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

async function launch(file: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({ args: ['.', file], cwd: projectRoot, env: { ...process.env, CI: '1', PLAYWRIGHT: '1' } })
  const page = await app.firstWindow()
  await page.waitForSelector('[role="tab"]', { timeout: 20_000 })
  await page.waitForTimeout(500)
  return { app, page }
}

async function dropRealFiles(page: Page, paths: string[]): Promise<void> {
  await page.evaluate(() => {
    let input = document.getElementById('__test_file_input__') as HTMLInputElement | null
    if (!input) {
      input = document.createElement('input')
      input.type = 'file'
      input.multiple = true
      input.id = '__test_file_input__'
      input.style.position = 'fixed'
      input.style.top = '0'
      input.style.left = '0'
      document.body.appendChild(input)
    }
  })
  await page.locator('#__test_file_input__').setInputFiles(paths)
  await page.evaluate(() => {
    const input = document.getElementById('__test_file_input__') as HTMLInputElement
    const dt = new DataTransfer()
    for (const f of Array.from(input.files ?? [])) dt.items.add(f)
    const target = document.querySelector('.app') as HTMLElement
    const evt = new DragEvent('drop', { bubbles: true, cancelable: true, dataTransfer: dt })
    target.dispatchEvent(evt)
    input.remove()
  })
}

test('dropping several files at once opens every one of them, not just the first', async () => {
  // Found by driving the real app: `handleDrop` only ever resolved
  // `e.dataTransfer.files[0]`, so dropping a multi-file selection silently
  // opened just the first file and discarded the rest with no feedback.
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-dragdrop-'))
  const first = path.join(dir, 'first.md')
  const second = path.join(dir, 'second.md')
  const third = path.join(dir, 'third.md')
  fs.writeFileSync(first, '# First\n')
  fs.writeFileSync(second, '# Second\n')
  fs.writeFileSync(third, '# Third\n')

  const { app, page } = await launch(first)
  try {
    await dropRealFiles(page, [second, third])
    await expect(page.getByRole('tab', { name: 'second.md' })).toBeVisible({ timeout: 10_000 })
    await expect(page.getByRole('tab', { name: 'third.md' })).toBeVisible({ timeout: 10_000 })
    await expect(page.getByRole('tab', { name: 'first.md' })).toBeVisible()
  } finally {
    kill(app)
  }
})

test('a drop while the active document is dirty asks before discarding the edit', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-dragdrop-dirty-'))
  const first = path.join(dir, 'first.md')
  const second = path.join(dir, 'second.md')
  fs.writeFileSync(first, '# First\n')
  fs.writeFileSync(second, '# Second\n')

  const { app, page } = await launch(first)
  try {
    await page.getByRole('button', { name: 'Editor' }).click()
    const textarea = page.locator('.editor-panel__textarea')
    await expect(textarea).toBeVisible({ timeout: 10_000 })
    await textarea.click()
    await page.keyboard.type(' unsaved edit')
    await expect(page.getByText(/unsaved changes/i)).toHaveCount(0)

    await dropRealFiles(page, [second])
    await expect(page.getByText(/unsaved changes/i).first()).toBeVisible({ timeout: 10_000 })

    // Discard and confirm the drop still completes.
    await page.getByRole('button', { name: /discard/i }).click()
    await expect(page.getByRole('tab', { name: 'second.md' })).toBeVisible({ timeout: 10_000 })
  } finally {
    kill(app)
  }
})
