import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'

/**
 * SEC-1 — verified allowlist-bypass hole, end to end.
 *
 * `path:register-dropped` (drag-drop) cannot verify a path actually came
 * from a real OS drag — the main process only checks `fs.existsSync`. Before
 * this fix, `registerDroppedPath` added the path to the SAME allowlist that
 * `save-file`/`save-binary-file` check for `existingPath`, so a compromised
 * renderer (an XSS-class bug in a viewer, this app's own stated threat
 * model) could call `registerDroppedPath('C:\\...\\important.docx')`
 * followed by a save with that `existingPath` and silently overwrite an
 * arbitrary file that exists on disk, with no dialog and no prompt.
 *
 * The fix splits the allowlist into a READ tier (drag-drop lands here only)
 * and a WRITE tier (only dialog/argv/recent-files-revalidated paths).
 * `save-file`/`save-binary-file` now only skip the save dialog for a WRITE-
 * tier `existingPath`.
 *
 * This spec exercises the REAL renderer-to-main IPC round trip (no faking
 * `path:register-dropped` or the allowlist) using the same real-drag
 * technique as drag-drop.spec.ts: a `<input type="file">` populated via
 * Playwright's `setInputFiles` (genuinely disk-backed `File`s, so
 * `webUtils.getPathForFile` resolves a real path), whose `.files` seed a
 * dispatched `drop` DragEvent.
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

async function editMarkdown(page: Page, text: string): Promise<void> {
  await page.getByRole('button', { name: 'Editor' }).click()
  const textarea = page.locator('.editor-panel__textarea')
  await expect(textarea).toBeVisible({ timeout: 10_000 })
  await textarea.fill(text)
}

/** Installs a spy on the main-process save dialog and returns a call counter getter. */
async function spyOnSaveDialog(app: ElectronApplication, resolveTo: { canceled: boolean; filePath?: string }): Promise<void> {
  await app.evaluate(({ dialog }, resolution) => {
    dialog.showSaveDialog = (async () => {
      ;(globalThis as { __saveDialogCalls?: number }).__saveDialogCalls =
        ((globalThis as { __saveDialogCalls?: number }).__saveDialogCalls ?? 0) + 1
      return resolution
    }) as typeof dialog.showSaveDialog
  }, resolveTo)
}

async function saveDialogCallCount(app: ElectronApplication): Promise<number> {
  return app.evaluate(() => (globalThis as { __saveDialogCalls?: number }).__saveDialogCalls ?? 0)
}

test('a drag-dropped file is NOT silently overwritten on Save — the save dialog is shown, and the save still completes', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-sec1-'))
  const argvFile = path.join(dir, 'argv-opened.md')
  const droppedFile = path.join(dir, 'dropped.md')
  fs.writeFileSync(argvFile, '# Opened via argv\n')
  fs.writeFileSync(droppedFile, '# Original dropped content\n')

  const { app, page } = await launch(argvFile)
  try {
    // Drag-drop the second file in — this is the registerDroppedPath path
    // (READ tier only, per the fix), NOT an Open-dialog/argv path.
    await dropRealFiles(page, [droppedFile])
    await expect(page.getByRole('tab', { name: 'dropped.md' })).toBeVisible({ timeout: 10_000 })
    await page.getByRole('tab', { name: 'dropped.md' }).click()

    await editMarkdown(page, '# Edited after drop\n\nThis must not be a silent overwrite.\n')

    // Simulate the user completing the native save dialog by picking the
    // SAME path back (the common "just re-save it" choice).
    await spyOnSaveDialog(app, { canceled: false, filePath: droppedFile })

    const before = fs.readFileSync(droppedFile, 'utf-8')
    await page.keyboard.press('Control+s')

    // The regression this test exists to catch: if the fix silently
    // overwrote the dropped file via the existingPath fast path, the save
    // dialog would never be invoked. If the fix instead broke saving
    // entirely for a dropped file, the dialog would show but the write
    // would never land. Both must be false — the dialog fires exactly
    // once AND the edit lands on disk.
    await expect.poll(() => fs.readFileSync(droppedFile, 'utf-8'), { timeout: 10_000 }).not.toBe(before)
    expect(await saveDialogCallCount(app)).toBe(1)
    expect(fs.readFileSync(droppedFile, 'utf-8')).toContain('Edited after drop')

    // Acceptance criterion: Save As (here, the dialog completing) on a
    // dropped file makes that path write-eligible afterward — a SECOND
    // save to the same path must now go silent (no second dialog call).
    await editMarkdown(page, '# Edited again, should save silently\n')
    await page.keyboard.press('Control+s')
    await expect.poll(() => fs.readFileSync(droppedFile, 'utf-8'), { timeout: 10_000 }).toContain(
      'Edited again, should save silently',
    )
    expect(await saveDialogCallCount(app)).toBe(1)
  } finally {
    kill(app)
  }
})

test('a normal Save on a file opened through the Open dialog still writes silently (no regression on the everyday save path)', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-sec1-everyday-'))
  const file = path.join(dir, 'everyday.md')
  fs.writeFileSync(file, '# Everyday save\n')

  const { app, page } = await launch(file)
  try {
    await spyOnSaveDialog(app, { canceled: true })

    await editMarkdown(page, '# Everyday save, edited\n')
    await page.keyboard.press('Control+s')

    await expect.poll(() => fs.readFileSync(file, 'utf-8'), { timeout: 10_000 }).toContain('edited')
    // argv-opened paths are trusted (WRITE tier) — the save dialog must
    // never even be consulted for the everyday save path.
    expect(await saveDialogCallCount(app)).toBe(0)
  } finally {
    kill(app)
  }
})
