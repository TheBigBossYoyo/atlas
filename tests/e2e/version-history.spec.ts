/**
 * VERSIONS-1 — version history, driven through the real app.
 *
 * The unit tests cover the store, the IPC and the panel separately; this is the
 * one that answers "does the feature work". It edits a document, saves it twice,
 * and checks the panel shows both states and can put the earlier one back —
 * through the same code path a person uses.
 *
 * Versions are recorded in MAIN when a document is written, so this also
 * exercises the design decision that makes the feature format-agnostic: nothing
 * in the markdown viewer asks for a version, yet saving produces one.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'

import { openMarkdownEditorWithShortcut } from './helpers/markdownViewMode'

const projectRoot = process.cwd()

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

async function launch(file: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({
    args: ['.', file],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  const page = await app.firstWindow()
  await expect(page.locator('[data-viewer="markdown"]')).toHaveCount(1, { timeout: 30_000 })
  // The markdown editor is a split-view pane, not the default preview — same
  // entry point `markdown-ctrlw-plain-field.spec.ts` uses.
  await openMarkdownEditorWithShortcut(page)
  return { app, page }
}

function tmpMarkdown(contents: string): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-versions-'))
  const file = path.join(dir, 'draft.md')
  fs.writeFileSync(file, contents, 'utf8')
  return file
}

/** Types into the markdown editor and saves, which is what records a version. */
async function editAndSave(page: Page, text: string): Promise<void> {
  const editor = page.locator('.editor-panel__textarea')
  await editor.click()
  await editor.fill(text)
  await page.keyboard.press('Control+s')
  await page.waitForTimeout(1500)
}

test('a save records a version, and an earlier version can be restored', async () => {
  const file = tmpMarkdown('# First draft\n')
  const { app, page } = await launch(file)
  try {
    await editAndSave(page, '# First draft\n\nThe opening paragraph.')
    await editAndSave(page, '# First draft\n\nA completely rewritten opening.')

    // Open the history panel from the toolbar.
    await page.getByRole('button', { name: /Version history/i }).click()
    const panel = page.locator('.version-history')
    await expect(panel).toBeVisible({ timeout: 10_000 })

    // Two saves, two versions — recorded by main without the viewer asking.
    const items = panel.locator('.version-history__item')
    await expect
      .poll(async () => items.count(), { message: 'the saves should have produced versions', timeout: 10_000 })
      .toBeGreaterThanOrEqual(2)

    // Restore the older one: the last item in a newest-first list.
    const count = await items.count()
    await items.nth(count - 1).getByRole('button', { name: /Restore/i }).click()
    await page.waitForTimeout(2500)

    // The file on disk is what proves it — the panel could show anything.
    const onDisk = fs.readFileSync(file, 'utf8')
    expect(onDisk, `file after restore: ${JSON.stringify(onDisk)}`).toContain('First draft')
    expect(onDisk).not.toContain('completely rewritten')
  } finally {
    kill(app)
  }
})

test('restoring is refused while there are unsaved changes', async () => {
  // The safety rule this feature rests on: restoring writes over the document,
  // and the state being replaced survives only if it was already written. The
  // button must be unavailable rather than destructive.
  const file = tmpMarkdown('# Draft\n')
  const { app, page } = await launch(file)
  try {
    await editAndSave(page, '# Draft\n\nSaved once.')

    // Now type WITHOUT saving.
    const editor = page.locator('textarea, .cm-content').first()
    await editor.click()
    await page.keyboard.type('\n\nUnsaved thoughts.', { delay: 5 })
    await page.waitForTimeout(600)

    await page.getByRole('button', { name: /Version history/i }).click()
    const panel = page.locator('.version-history')
    await expect(panel).toBeVisible({ timeout: 10_000 })

    await expect(panel.getByRole('button', { name: /Restore/i }).first()).toBeDisabled()
    await expect(panel.locator('.version-history__notice')).toBeVisible()
  } finally {
    kill(app)
  }
})
