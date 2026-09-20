import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execFileSync } from 'node:child_process'

import { _electron as electron, expect, test } from '@playwright/test'
import type { ElectronApplication, Page } from '@playwright/test'

/**
 * FIELD-01 — live repro for the bug found driving the real app: typing in
 * the markdown editor's plain `<textarea>` (RawEditor), then pressing
 * Ctrl+W, did nothing at all — no close, no unsaved-changes prompt —
 * because `useUniversalShortcuts`'s `ctx.inPlainField` gate covered `w`
 * (Ctrl+S/Ctrl+P were already carved out of that gate; Ctrl+W was not). That
 * silently ate the exact path a user takes to close a document they just
 * edited, which also meant the unsaved-changes prompt could never fire on
 * that path. This only shows up against the real app/real DOM focus
 * handling — jsdom-based unit coverage lives in
 * `src/hooks/__tests__/useUniversalShortcuts.test.tsx`.
 */

const projectRoot = process.cwd()
const fixtureDir = path.join(projectRoot, 'tests', 'e2e', 'fixtures')

test.beforeAll(() => {
  execFileSync(process.execPath, ['tests/e2e/fixtures/generate.mjs'], {
    cwd: projectRoot,
    stdio: 'inherit',
  })
})

function copyMarkdownFixtureToTemp(): string {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-e2e-ctrlw-field-'))
  const dest = path.join(tempDir, 'sample.md')
  fs.copyFileSync(path.join(fixtureDir, 'sample.md'), dest)
  return dest
}

async function launch(filePath: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({
    args: ['.', filePath],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1', ATLAS_HIDDEN_WINDOW: '1' },
  })
  const page = await app.firstWindow()
  return { app, page }
}

function kill(app: ElectronApplication): void {
  const pid = app.process().pid
  try {
    if (pid) execFileSync('taskkill', ['/PID', String(pid), '/T', '/F'], { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

test('Ctrl+W with focus inside the markdown editor textarea raises the unsaved-changes prompt (and Discard actually closes)', async () => {
  const tempPath = copyMarkdownFixtureToTemp()
  const { app, page } = await launch(tempPath)
  try {
    await expect(page.locator('[data-viewer="markdown"]')).toHaveCount(1, { timeout: 15_000 })

    // Mount RawEditor — a plain <textarea> with no key handling of its own,
    // exactly the surface the bug report reproduced against.
    await page.keyboard.press('Control+2')
    const textarea = page.locator('.editor-panel__textarea')
    await expect(textarea).toBeVisible({ timeout: 10_000 })
    await textarea.click()
    await textarea.fill('# Edited by the FIELD-01 e2e spec\n')

    // Focus never leaves the textarea for this keypress — this is the exact
    // "type, then Ctrl+W" flow the bug report describes.
    await page.keyboard.press('Control+w')

    // Pre-fix: nothing happened at all (no dialog, no close) because the
    // shell handler bailed out on `ctx.inPlainField` before ever reaching
    // `closeFile`. Post-fix: closeFile() runs, finds the document dirty, and
    // raises the same in-app unsaved-changes prompt Ctrl+S/other close paths
    // use.
    const dialog = page.getByRole('alertdialog')
    await expect(dialog).toBeVisible({ timeout: 5_000 })

    // Follow the close path all the way through: Discard actually closes
    // the document back to the Welcome screen, proving this isn't just a
    // dialog appearing in isolation.
    await dialog.getByRole('button', { name: 'Discard' }).click()
    await expect(page.getByRole('button', { name: 'Load Sample Document' })).toBeVisible({ timeout: 10_000 })
  } finally {
    kill(app)
    fs.rmSync(path.dirname(tempPath), { recursive: true, force: true })
  }
})
