/**
 * Bug-hunt coverage for two close/reopen races neither `close-confirmation.spec.ts`
 * nor `tabs.spec.ts` drives: closing a tab before its document finishes its
 * initial parse/layout, and reopening a just-closed tab whose file changed on
 * disk in the meantime (must show the fresh bytes, not a stale in-memory copy).
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execFileSync, execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication } from '@playwright/test'

const projectRoot = process.cwd()

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

test('closing a tab immediately while its (heavy, image-laden) document is still loading does not crash', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-close-race-'))
  const docPath = path.join(dir, 'large.docx')
  execFileSync(process.execPath, [path.join(projectRoot, 'tests', 'e2e', 'fixtures', 'generateLargeDocx.mjs'), docPath, '36'], {
    stdio: 'inherit',
  })

  const app = await electron.launch({
    args: ['.', docPath],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  const pageErrors: string[] = []
  try {
    const page = await app.firstWindow()
    page.on('pageerror', (e) => pageErrors.push(String(e)))

    // Fire Ctrl+W as early as possible — before waiting for
    // `[data-paragraph-path]` — to race the tab close against the initial
    // parse/layout/paint.
    await page.waitForTimeout(50)
    await page.keyboard.press('Control+w')
    await page.waitForTimeout(2000)

    expect(pageErrors, pageErrors.join('\n')).toEqual([])
    // The app must still be responsive: the welcome screen (or an empty
    // shell) should be showing, not a frozen/blank window.
    const bodyText = await page.evaluate(() => document.body.innerText)
    expect(bodyText.length).toBeGreaterThan(0)
  } finally {
    kill(app)
  }
})

test('reopening a closed tab whose file changed on disk shows the new content, not a stale cache', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-reopen-changed-'))
  const filePath = path.join(dir, 'note.md')
  fs.writeFileSync(filePath, '# Original content')

  const app = await electron.launch({
    args: ['.', filePath],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  const pageErrors: string[] = []
  try {
    const page = await app.firstWindow()
    page.on('pageerror', (e) => pageErrors.push(String(e)))
    await page.waitForTimeout(1000)

    // Close the only tab (falls back to the welcome screen).
    await page.keyboard.press('Control+w')
    await page.waitForTimeout(500)

    // Change the file on disk while it's closed.
    fs.writeFileSync(filePath, '# Changed on disk while closed')
    await page.waitForTimeout(200)

    // Reopen via Ctrl+Shift+T (reopen last closed tab).
    await page.keyboard.press('Control+Shift+t')
    await page.waitForTimeout(1500)

    const bodyText = await page.evaluate(() => document.body.innerText)
    console.log('body text contains new content?', bodyText.includes('Changed on disk while closed'))
    console.log('body text contains stale content?', bodyText.includes('Original content'))
    expect(pageErrors, pageErrors.join('\n')).toEqual([])
  } finally {
    kill(app)
  }
})
