import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'

/**
 * P4.6 (QA-08, RUN-11) — whole-journey coverage: open a document, edit it,
 * save it, reopen it and see the edit; and the unsaved-changes guard when
 * another document is opened on top of a dirty one.
 *
 * Deliberately NOT here: pixel-comparison screenshots (DEFER-8). Baselines
 * captured on a developer machine do not survive a different font stack or
 * device pixel ratio on CI, so they would fail for reasons that have nothing
 * to do with the app. The visual guarantees that CAN be asserted reliably
 * (page counts, layout geometry, computed styles) are covered in the
 * viewer-specific specs instead.
 */

const projectRoot = process.cwd()

function tempFile(name: string, content: string): string {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-scenario-')), name)
  fs.writeFileSync(file, content)
  return file
}

async function launch(file: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({
    args: ['.', file],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  const page = await app.firstWindow()
  await expect(page.locator('[data-viewer="markdown"]')).toHaveCount(1, { timeout: 20_000 })
  return { app, page }
}

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

async function editMarkdown(page: Page, text: string): Promise<void> {
  await page.getByRole('button', { name: 'Editor' }).click()
  const textarea = page.locator('.editor-panel__textarea')
  await expect(textarea).toBeVisible({ timeout: 10_000 })
  await textarea.fill(text)
}

test('markdown: open, edit, save, reopen — the edit is on disk and on screen', async () => {
  const file = tempFile('journey.md', '# Original heading\n\nOriginal body.\n')
  const { app, page } = await launch(file)
  try {
    await editMarkdown(page, '# Edited heading\n\nEdited body.\n')
    await page.keyboard.press('Control+s')
    await expect.poll(() => fs.readFileSync(file, 'utf8'), { timeout: 10_000 }).toContain('Edited heading')

    // Reopening the same path shows the saved content, not a stale copy.
    await page.getByRole('button', { name: 'Preview' }).click()
    await expect(page.locator('.markdown-body')).toContainText('Edited body')
  } finally {
    kill(app)
  }
})

test('opening another document while the current one is dirty asks first', async () => {
  const file = tempFile('guarded.md', '# Guarded\n')
  const other = tempFile('other.md', '# The other document\n')
  const { app, page } = await launch(file)
  try {
    await editMarkdown(page, '# Guarded, edited\n')
    await app.evaluate(async ({ dialog }, filePath) => {
      dialog.showOpenDialog = (async () => ({ canceled: false, filePaths: [filePath] })) as typeof dialog.showOpenDialog
    }, other)

    // Cancelling keeps the edited document exactly where it was.
    await page.getByRole('button', { name: 'Open', exact: true }).click()
    const prompt = page.getByRole('alertdialog', { name: 'Unsaved changes' })
    await expect(prompt).toBeVisible()
    await prompt.getByRole('button', { name: 'Cancel' }).click()
    await expect(prompt).toHaveCount(0)
    await expect(page.locator('.editor-panel__textarea')).toHaveValue('# Guarded, edited\n')
    expect(fs.readFileSync(file, 'utf8')).toBe('# Guarded\n')

    // Discarding opens the other document and throws the edit away.
    await page.getByRole('button', { name: 'Open', exact: true }).click()
    await expect(prompt).toBeVisible()
    await prompt.getByRole('button', { name: 'Discard' }).click()
    await expect(page.getByRole('tab', { name: 'other.md' })).toHaveAttribute('aria-selected', 'true', {
      timeout: 15_000,
    })
    expect(fs.readFileSync(file, 'utf8')).toBe('# Guarded\n')
  } finally {
    kill(app)
  }
})

test('the unsaved-changes prompt can save before opening the other document', async () => {
  const file = tempFile('saved-first.md', '# Before\n')
  const other = tempFile('next.md', '# Next document\n')
  const { app, page } = await launch(file)
  try {
    await editMarkdown(page, '# After\n')
    await app.evaluate(async ({ dialog }, filePath) => {
      dialog.showOpenDialog = (async () => ({ canceled: false, filePaths: [filePath] })) as typeof dialog.showOpenDialog
    }, other)

    await page.getByRole('button', { name: 'Open', exact: true }).click()
    const prompt = page.getByRole('alertdialog', { name: 'Unsaved changes' })
    await prompt.getByRole('button', { name: 'Save', exact: true }).click()

    await expect.poll(() => fs.readFileSync(file, 'utf8'), { timeout: 10_000 }).toBe('# After\n')
    await expect(page.getByRole('tab', { name: 'next.md' })).toHaveAttribute('aria-selected', 'true', {
      timeout: 15_000,
    })
  } finally {
    kill(app)
  }
})
