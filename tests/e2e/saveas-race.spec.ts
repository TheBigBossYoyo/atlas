import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import { Document, Packer, Paragraph, TextRun } from 'docx'

/**
 * SAVE-1 — a non-markdown Save As (DOCX/PPTX/XLSX/ODP/ODS all write through
 * the same `save-binary-file` IPC channel) that resolves AFTER the user has
 * already switched to a different tab.
 *
 * Before the fix, `handleViewerSavedPath` (App.tsx) trusted the *currently
 * showing* tab to be the one that was just saved. Since the write is a
 * `window.electronAPI.saveBinaryFile(...)` `await` the user is free to
 * out-race by switching tabs, a save that finished late would rename the
 * NEW tab to the saved document's path and silently re-read that path over
 * whatever the user had switched to.
 *
 * Making the race deterministic: rather than guessing at timing (which the
 * task this fixes explicitly warns against — a flaky race test is worse than
 * none), this reaches into the running main process (the same
 * `app.evaluate(({ dialog }) => ...)` pattern `new-document.spec.ts` already
 * uses to stub `dialog.showSaveDialog`) and wraps the real
 * `ipcMain.handle('save-binary-file', ...)` listener with a gate that only
 * calls through to the ORIGINAL handler once the test explicitly releases
 * it. `ipcMain._invokeHandlers` is Electron's own internal map backing
 * `ipcMain.handle`/`removeHandler` (verified directly against the installed
 * Electron 44.4.1 build in this repo) — reading the handler already
 * registered for the channel and swapping in a wrapper this way delays
 * exactly the moment `saveBinaryFile()`'s promise resolves in the renderer,
 * with nothing about the dialog stub, the write itself, or `main.cjs`
 * changed. The test only ever proceeds past the gate by polling for a real
 * state change (the gate variable existing, then the file actually
 * appearing on disk) — never a fixed sleep standing in for either race edge.
 */

const projectRoot = process.cwd()

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

async function createDocxFixture(dir: string, text: string): Promise<string> {
  const doc = new Document({
    sections: [{ children: [new Paragraph({ children: [new TextRun(text)] })] }],
  })
  const file = path.join(dir, 'a.docx')
  fs.writeFileSync(file, await Packer.toBuffer(doc))
  return file
}

/**
 * Installs the save-binary-file gate in the main process and points the
 * (stubbed) native Save dialog at `targetPath`. Returns nothing — poll
 * `waitForGateArmed` to know when a save is blocked on it, and call
 * `releaseGate` to let it through.
 */
async function installSaveGate(app: ElectronApplication, targetPath: string): Promise<void> {
  await app.evaluate(({ ipcMain, dialog }, targetPathArg) => {
    dialog.showSaveDialog = (() =>
      Promise.resolve({ canceled: false, filePath: targetPathArg })) as typeof dialog.showSaveDialog

    const channel = 'save-binary-file'
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const invokeHandlers = (ipcMain as any)._invokeHandlers as Map<string, (...args: unknown[]) => unknown>
    const original = invokeHandlers.get(channel)
    if (!original) throw new Error(`no handler registered for ${channel}`)
    invokeHandlers.set(channel, async (...args: unknown[]) => {
      await new Promise<void>((resolve) => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        ;(global as any).__atlasReleaseSave = resolve
      })
      return original(...args)
    })
  }, targetPath)
}

/** Waits until a save is actually blocked on the gate (not a fixed sleep — polls the real state). */
async function waitForGateArmed(app: ElectronApplication): Promise<void> {
  await app.evaluate(async () => {
    const start = Date.now()
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    while (typeof (global as any).__atlasReleaseSave !== 'function') {
      if (Date.now() - start > 10_000) throw new Error('save-binary-file gate never armed')
      await new Promise((resolve) => setTimeout(resolve, 20))
    }
  })
}

async function releaseGate(app: ElectronApplication): Promise<void> {
  await app.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    ;(global as any).__atlasReleaseSave()
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    ;(global as any).__atlasReleaseSave = undefined
  })
}

test('a DOCX Save As that resolves after switching tabs updates the saved document\'s own tab, not the one now showing', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-saveas-race-'))
  const docxPath = await createDocxFixture(dir, 'Original DOCX content')
  const mdPath = path.join(dir, 'second.md')
  fs.writeFileSync(mdPath, '# Second document\n\nUntouched markdown tab.\n')
  const renamedDocxPath = path.join(dir, 'a-renamed.docx')

  const app = await electron.launch({
    args: ['.', docxPath],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1', ATLAS_HIDDEN_WINDOW: '1' },
  })
  const pageErrors: string[] = []
  try {
    const page: Page = await app.firstWindow()
    page.on('pageerror', (e) => pageErrors.push(String(e)))

    await page.waitForSelector('[data-paragraph-path]', { state: 'attached', timeout: 20_000 })
    await expect(page.getByRole('tab', { name: 'a.docx' })).toHaveAttribute('aria-selected', 'true')
    // Let save/saveAs registration settle (mirrors docx-editor.spec.ts's own launch helper).
    await page.waitForTimeout(500)

    // Arm the gate: the next Save As will not resolve until the test says so.
    await installSaveGate(app, renamedDocxPath)

    // Start Save As on a.docx.
    await page.keyboard.press('Control+Shift+S')
    await waitForGateArmed(app)

    // Nothing written yet — the gate really is blocking the write, not just
    // the renderer's view of it.
    expect(fs.existsSync(renamedDocxPath)).toBe(false)

    // Switch to a different (markdown) tab while that Save As is still stuck
    // on the gate — the exact race SAVE-1 is about.
    await app.evaluate(async ({ dialog }, filePath) => {
      dialog.showOpenDialog = (async () => ({ canceled: false, filePaths: [filePath] })) as typeof dialog.showOpenDialog
    }, mdPath)
    await page.getByRole('button', { name: 'Open', exact: true }).click()
    const mdTab = page.getByRole('tab', { name: 'second.md' })
    await expect(mdTab).toBeVisible({ timeout: 20_000 })
    await expect(mdTab).toHaveAttribute('aria-selected', 'true')
    await expect(page.locator('.markdown-body')).toContainText('Untouched markdown tab')

    // NOW let the stale a.docx Save As finish.
    await releaseGate(app)
    await expect.poll(() => fs.existsSync(renamedDocxPath), { timeout: 10_000 }).toBe(true)
    // Give the renderer a moment to process the IPC resolution and act on it.
    await page.waitForTimeout(300)

    // The showing tab (second.md) is completely untouched: still selected,
    // still markdown, content unchanged, no docx viewer mounted over it.
    await expect(mdTab).toHaveAttribute('aria-selected', 'true')
    await expect(page.locator('.markdown-body')).toContainText('Untouched markdown tab')
    await expect(page.locator('[data-viewer="docx"]')).toHaveCount(0)

    // a.docx's own (background) tab followed its new path instead.
    await expect(page.getByRole('tab', { name: 'a-renamed.docx' })).toBeVisible({ timeout: 5_000 })
    await expect(page.getByRole('tab', { name: 'a.docx' })).toHaveCount(0)

    // Reactivating the saved document shows it at its NEW path, with its
    // content intact — not a re-read of the (now-abandoned) original path.
    await page.getByRole('tab', { name: 'a-renamed.docx' }).click()
    await page.waitForSelector('[data-paragraph-path]', { state: 'attached', timeout: 20_000 })
    await expect(page.locator('[data-paragraph-path="0"]').first()).toContainText('Original DOCX content')

    expect(pageErrors, pageErrors.join('\n')).toEqual([])
  } finally {
    kill(app)
  }
})

test('a DOCX Save As that resolves while still showing follows the tab exactly as before (no regression)', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-saveas-race-control-'))
  const docxPath = await createDocxFixture(dir, 'Control-case content')
  const renamedDocxPath = path.join(dir, 'still-showing-renamed.docx')

  const app = await electron.launch({
    args: ['.', docxPath],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1', ATLAS_HIDDEN_WINDOW: '1' },
  })
  try {
    const page: Page = await app.firstWindow()
    await page.waitForSelector('[data-paragraph-path]', { state: 'attached', timeout: 20_000 })
    await page.waitForTimeout(500)

    await app.evaluate(({ dialog }, targetPathArg) => {
      dialog.showSaveDialog = (() =>
        Promise.resolve({ canceled: false, filePath: targetPathArg })) as typeof dialog.showSaveDialog
    }, renamedDocxPath)

    await page.keyboard.press('Control+Shift+S')
    await expect.poll(() => fs.existsSync(renamedDocxPath), { timeout: 10_000 }).toBe(true)

    await expect(page.getByRole('tab', { name: 'still-showing-renamed.docx' })).toHaveAttribute(
      'aria-selected',
      'true',
      { timeout: 10_000 },
    )
    await expect(page.getByRole('tab', { name: 'a.docx' })).toHaveCount(0)
  } finally {
    kill(app)
  }
})
