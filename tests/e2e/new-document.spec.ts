import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'

/**
 * NEW-01 — end-to-end coverage for the owner's two reported gaps:
 *   (a) a file made outside Atlas (Windows Explorer's "New > Word Document"
 *       on a PC without Office) is a genuine 0-byte `.docx` on disk — opening
 *       it must show a real, editable, blank document instead of "Failed to
 *       unzip DOCX: End of data reached…", and a normal Save must turn it
 *       into a real non-empty `.docx` at that same path.
 *   (b) the toolbar's "New" menu creates a brand-new document itself, via
 *       main's native Save dialog (stubbed here the same way
 *       `export.spec.ts` stubs `dialog.showSaveDialog` — this only replaces
 *       the DIALOG, not any of the write/allowlist logic downstream of it).
 */

const projectRoot = process.cwd()

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

async function stubSaveDialog(app: ElectronApplication, targetPath: string): Promise<void> {
  await app.evaluate(({ dialog }, targetPathArg) => {
    dialog.showSaveDialog = (() =>
      Promise.resolve({ canceled: false, filePath: targetPathArg })) as typeof dialog.showSaveDialog
  }, targetPath)
}

test('opening a 0-byte .docx shows a real blank editor, and Save writes a real document', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-new-doc-'))
  const file = path.join(dir, 'empty.docx')
  fs.writeFileSync(file, Buffer.alloc(0)) // genuinely 0 bytes on disk, like Explorer's "New > Word Document"
  expect(fs.statSync(file).size).toBe(0)

  const app = await electron.launch({
    args: ['.', file],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  try {
    const page: Page = await app.firstWindow()

    // The blank template opened, fully editable — not a parse-error message.
    // The one (empty) paragraph's own line box is legitimately 0-width (no
    // content yet), so it never counts as Playwright-"visible" — only wait
    // for it to be attached, and click the editable column that contains it
    // instead of the zero-width line itself.
    await page.waitForSelector('[data-paragraph-path]', { state: 'attached', timeout: 20_000 })
    await expect(page.locator('.docx-viewer').first()).toBeVisible()

    const paragraph = page.locator('[data-paragraph-path="0"]').first()
    await page.locator('.docx-page__column').first().click({ position: { x: 5, y: 5 } })
    await page.keyboard.type('Hello from a template that used to be 0 bytes', { delay: 20 })
    await expect.poll(() => paragraph.textContent()).toContain('Hello from a template')

    // Still 0 bytes on disk — typing alone never writes anything.
    expect(fs.statSync(file).size).toBe(0)

    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect.poll(() => fs.statSync(file).size, { timeout: 10_000 }).toBeGreaterThan(0)

    const zip = await (await import('jszip')).default.loadAsync(fs.readFileSync(file))
    const documentXml = await zip.file('word/document.xml')!.async('string')
    expect(documentXml).toContain('Hello from a template')
  } finally {
    kill(app)
  }
})

/**
 * Regression — the "New" panel reused the shared `.dropdown__menu` geometry,
 * which is right-aligned because every other toolbar dropdown (Export, Theme,
 * Language) lives at the far RIGHT of the toolbar and grows leftwards into it.
 * "New" is the second control from the LEFT, so right-aligning grew the panel
 * off the left edge of the window: measured at x = -76px, i.e. 76px of the
 * panel (its icons and the start of every label) was unreachable. Four of the
 * six labels also wrapped onto two lines, because an absolutely-positioned box
 * shrink-to-fits against its trigger's width and so never grew past its 200px
 * `min-width`.
 *
 * Neither unit nor e2e coverage caught it: Playwright still clicks an item
 * whose visible remainder is on-screen, and jsdom has no layout at all. This
 * asserts the real rendered box instead.
 */
test('the New menu opens fully inside the window, with no wrapped labels', async () => {
  const app = await electron.launch({
    args: ['.'],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  try {
    const page: Page = await app.firstWindow()
    await expect(page.getByRole('button', { name: 'Load Sample Document' })).toBeVisible({ timeout: 20_000 })

    await page.getByRole('button', { name: 'New', exact: true }).click()
    const menu = page.locator('.dropdown__menu')
    await expect(menu).toBeVisible()

    const viewportWidth = await page.evaluate(() => window.innerWidth)
    const box = (await menu.boundingBox())!
    expect(box).not.toBeNull()
    expect(box.x).toBeGreaterThanOrEqual(0)
    expect(box.x + box.width).toBeLessThanOrEqual(viewportWidth)

    // Every label on a single line: a wrapped one is taller than one line box.
    const labelHeights = await page.$$eval('.dropdown__menu .dropdown__item-label', (els) =>
      els.map((el) => el.getBoundingClientRect().height),
    )
    expect(labelHeights).toHaveLength(6)
    for (const height of labelHeights) expect(height).toBeLessThan(24)
  } finally {
    kill(app)
  }
})

test('New → Word document creates a real file and opens it in a new tab', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-new-doc-'))
  const targetPath = path.join(dir, 'NewDoc.docx')

  const app = await electron.launch({
    args: ['.'],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  try {
    const page: Page = await app.firstWindow()
    await expect(page.getByRole('button', { name: 'Load Sample Document' })).toBeVisible({ timeout: 20_000 })

    await stubSaveDialog(app, targetPath)
    await page.getByRole('button', { name: 'New', exact: true }).click()
    await page.getByRole('button', { name: 'Word document (.docx)' }).click()

    await expect(page.getByRole('tab', { name: 'NewDoc.docx' })).toBeVisible({ timeout: 20_000 })
    // See the other test's comment — the blank paragraph's line box is
    // legitimately 0-width, so only "attached" (not "visible") is checked.
    await page.waitForSelector('[data-paragraph-path]', { state: 'attached', timeout: 20_000 })

    expect(fs.existsSync(targetPath)).toBe(true)
    expect(fs.statSync(targetPath).size).toBeGreaterThan(0)
  } finally {
    kill(app)
  }
})
