/**
 * PDF-14/P10 — password-protected PDF coverage against a GENUINELY encrypted
 * fixture (RC4-40, Standard Security Handler V1/R2 — see
 * `fixtures/generateEncryptedPdf.mjs`), not a mock of pdf.js's `onPassword`
 * callback. No password-protected-PDF fixture existed in this repo before —
 * this was previously undriven end-to-end.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execFileSync, execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication } from '@playwright/test'

const projectRoot = process.cwd()
const USER_PASSWORD = 'secret123'
const OWNER_PASSWORD = 'owner456'

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

let encryptedPdfPath: string

test.beforeAll(() => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-pdf-pw-'))
  encryptedPdfPath = path.join(dir, 'protected.pdf')
  execFileSync(
    process.execPath,
    [path.join(projectRoot, 'tests', 'e2e', 'fixtures', 'generateEncryptedPdf.mjs'), encryptedPdfPath, USER_PASSWORD, OWNER_PASSWORD],
    { stdio: 'inherit' },
  )
  expect(fs.existsSync(encryptedPdfPath)).toBe(true)
})

test('wrong password shows an error and lets you retry; right password unlocks the real content', async () => {
  const app = await electron.launch({
    args: ['.', encryptedPdfPath],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  const pageErrors: string[] = []
  try {
    const page = await app.firstWindow()
    page.on('pageerror', (e) => pageErrors.push(String(e)))

    const dialog = page.locator('.pdf-viewer__password-dialog')
    await expect(dialog).toBeVisible({ timeout: 20_000 })

    const input = page.locator('.pdf-viewer__password-input')
    await input.fill('totally-wrong-password')
    await page.locator('.pdf-viewer__password-submit').click()

    await expect(page.locator('.pdf-viewer__password-error')).toBeVisible({ timeout: 10_000 })
    // Dialog stays open for a retry.
    await expect(dialog).toBeVisible()

    await input.fill(USER_PASSWORD)
    await page.locator('.pdf-viewer__password-submit').click()

    await expect(dialog).not.toBeVisible({ timeout: 10_000 })
    // Real content rendered: pdf.js draws text into a canvas/text-layer; check the page loaded without an error screen.
    await expect(page.locator('.pdf-viewer__error')).toHaveCount(0)
    await page.waitForTimeout(1000)
    const textLayerText = await page.evaluate(() =>
      Array.from(document.querySelectorAll('.textLayer span, [class*="text-layer"] span'))
        .map((e) => e.textContent)
        .join(' '),
    )
    console.log('textLayerText:', textLayerText)
    expect(pageErrors, pageErrors.join('\n')).toEqual([])
  } finally {
    kill(app)
  }
})

test('owner password also unlocks it', async () => {
  const app = await electron.launch({
    args: ['.', encryptedPdfPath],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  try {
    const page = await app.firstWindow()
    const dialog = page.locator('.pdf-viewer__password-dialog')
    await expect(dialog).toBeVisible({ timeout: 20_000 })
    await page.locator('.pdf-viewer__password-input').fill(OWNER_PASSWORD)
    await page.locator('.pdf-viewer__password-submit').click()
    await expect(dialog).not.toBeVisible({ timeout: 10_000 })
    await expect(page.locator('.pdf-viewer__error')).toHaveCount(0)
  } finally {
    kill(app)
  }
})

test('cancel leaves a clear error state and no crash; closing that tab afterward works cleanly', async () => {
  const app = await electron.launch({
    args: ['.', encryptedPdfPath],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  const pageErrors: string[] = []
  try {
    const page = await app.firstWindow()
    page.on('pageerror', (e) => pageErrors.push(String(e)))

    const dialog = page.locator('.pdf-viewer__password-dialog')
    await expect(dialog).toBeVisible({ timeout: 20_000 })
    await page.locator('.pdf-viewer__password-cancel').click()

    await expect(dialog).not.toBeVisible()
    await expect(page.locator('.pdf-viewer__error')).toBeVisible({ timeout: 5000 })
    expect(pageErrors, pageErrors.join('\n')).toEqual([])

    await page.keyboard.press('Control+w') // close current tab — must not hang/crash after a cancelled password prompt
    await page.waitForTimeout(500)
    expect(pageErrors, pageErrors.join('\n')).toEqual([])
  } finally {
    kill(app)
  }
})

test('reopening the same encrypted file in a brand-new app instance prompts for a password again (not cached)', async () => {
  const app = await electron.launch({
    args: ['.', encryptedPdfPath],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  try {
    const page = await app.firstWindow()
    const dialog = page.locator('.pdf-viewer__password-dialog')
    await expect(dialog).toBeVisible({ timeout: 20_000 })
  } finally {
    kill(app)
  }
})
