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

test('USR-19: switching away from a running file stops it instead of orphaning it', async () => {
  // Found by driving the real app: leaving (or closing) a running file's tab
  // unmounted the only Stop button that could ever reach that run. Main
  // allows exactly one run at a time, so the child kept running with no way
  // to stop it short of its 60s timeout or quitting the app, and every other
  // file's Run just failed with "a program is already running".
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-code-orphan-'))
  const marker = path.join(dir, 'marker.txt')
  const loop = path.join(dir, 'loop.js')
  const other = path.join(dir, 'other.js')
  fs.writeFileSync(
    loop,
    `const fs=require('fs');const p=${JSON.stringify(marker)};` +
      `setInterval(()=>fs.writeFileSync(p,String(process.pid)),100);\n`,
  )
  fs.writeFileSync(other, 'console.log("other-ran")\n')

  const { app, page } = await launch(loop)
  try {
    await app.evaluate(({ dialog }) => {
      dialog.showMessageBox = (() => Promise.resolve({ response: 0, checkboxChecked: false })) as typeof dialog.showMessageBox
    })
    await page.getByRole('button', { name: 'Run' }).click()
    await expect(page.getByRole('region', { name: 'Program output' })).toBeVisible()
    await expect.poll(() => fs.existsSync(marker), { timeout: 10_000 }).toBe(true)
    const pid = fs.readFileSync(marker, 'utf8').trim()

    // Switch to a different file — the running one's tab (and its Stop
    // button) goes away, but the child must go with it.
    await app.evaluate(async ({ dialog }, filePath) => {
      dialog.showOpenDialog = (async () => ({ canceled: false, filePaths: [filePath] })) as typeof dialog.showOpenDialog
    }, other)
    await page.getByRole('button', { name: 'Open', exact: true }).click()
    await expect(page.getByRole('tab', { name: 'other.js' })).toBeVisible({ timeout: 10_000 })

    await expect
      .poll(
        () => {
          try {
            execSync(`tasklist /FI "PID eq ${pid}"`).toString()
            return execSync(`tasklist /FI "PID eq ${pid}"`).toString().includes(pid)
          } catch {
            return false
          }
        },
        { timeout: 10_000 },
      )
      .toBe(false)

    // The global "one run at a time" slot is free again.
    await page.getByRole('button', { name: 'Run' }).click()
    await expect(page.getByRole('region', { name: 'Program output' })).toContainText('other-ran', { timeout: 20_000 })
  } finally {
    kill(app)
  }
})
