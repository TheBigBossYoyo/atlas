import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'

/** USR-18/USR-19 — the code editor and the explicit Run action in the real app. */

const projectRoot = process.cwd()

function createFile(name: string, content: string): string {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-code-')), name)
  fs.writeFileSync(file, content)
  return file
}

async function launch(file: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({ args: ['.', file], cwd: projectRoot, env: { ...process.env, CI: '1', PLAYWRIGHT: '1' } })
  const page = await app.firstWindow()
  await page.waitForSelector('.cm-content', { timeout: 20_000 })
  await page.waitForTimeout(400)
  return { app, page }
}

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

test('USR-18: edits, finds and saves a source file', async () => {
  const file = createFile('script.js', 'const greeting = "hi"\n')
  const { app, page } = await launch(file)
  try {
    // Line numbers and highlighting, not a read-only <pre>.
    await expect(page.locator('.cm-gutters')).toBeVisible()
    await expect(page.locator('.cm-content')).toContainText('greeting')

    await page.locator('.cm-content').click()
    await page.keyboard.press('Control+End')
    await page.keyboard.type('const answer = 42\n')
    await expect(page.locator('.cm-content')).toContainText('answer')

    await page.keyboard.press('Control+f')
    await expect(page.locator('.cm-search')).toBeVisible()
    await page.keyboard.press('Escape')

    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect.poll(() => fs.readFileSync(file, 'utf8'), { timeout: 10_000 }).toContain('const answer = 42')
  } finally {
    kill(app)
  }
})

test('USR-19: runs the file only after confirmation, and streams its output', async () => {
  const file = createFile('hello.js', 'console.log("atlas-run-ok")\n')
  const { app, page } = await launch(file)
  try {
    // Decline first: nothing runs.
    await app.evaluate(({ dialog }) => {
      dialog.showMessageBox = (() => Promise.resolve({ response: 1, checkboxChecked: false })) as typeof dialog.showMessageBox
    })
    await page.getByRole('button', { name: 'Run' }).click()
    await expect(page.getByRole('region', { name: 'Program output' })).toHaveCount(0)

    // Accept: the program runs and its output is shown.
    await app.evaluate(({ dialog }) => {
      dialog.showMessageBox = (() => Promise.resolve({ response: 0, checkboxChecked: false })) as typeof dialog.showMessageBox
    })
    await page.getByRole('button', { name: 'Run' }).click()
    const output = page.getByRole('region', { name: 'Program output' })
    await expect(output).toBeVisible()
    await expect(output).toContainText('atlas-run-ok', { timeout: 20_000 })
    await expect(output).toContainText('Finished (exit code 0).')
  } finally {
    kill(app)
  }
})

test('USR-19: a long-running program can be stopped', async () => {
  const file = createFile('loop.js', 'setInterval(() => console.log("tick"), 50)\n')
  const { app, page } = await launch(file)
  try {
    await app.evaluate(({ dialog }) => {
      dialog.showMessageBox = (() => Promise.resolve({ response: 0, checkboxChecked: false })) as typeof dialog.showMessageBox
    })
    await page.getByRole('button', { name: 'Run' }).click()
    const output = page.getByRole('region', { name: 'Program output' })
    await expect(output).toContainText('tick', { timeout: 20_000 })

    await page.getByRole('button', { name: 'Stop' }).click()
    await expect(output).toContainText('Stopped.', { timeout: 20_000 })
    await expect(page.getByRole('button', { name: 'Run' })).toBeVisible()
  } finally {
    kill(app)
  }
})
