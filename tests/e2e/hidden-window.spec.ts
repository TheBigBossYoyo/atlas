import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'

/**
 * `ATLAS_HIDDEN_WINDOW=1` — the test-only invisible-window mode used when
 * driving the real app for a UI sweep, so the run does not cover the screen or
 * steal focus for however long it takes.
 *
 * It is NOT headless and must never be reported as such: the window is created,
 * shown and composited exactly as in a normal run, then made fully transparent
 * and click-through. Leaving it genuinely shown is load-bearing — a
 * `show: false` window is marked hidden by Chromium, which stops compositing
 * and `requestAnimationFrame`, and Playwright's actionability checks (which
 * wait for an element's box to be stable across two animation frames) then
 * hang forever, so every click times out.
 *
 * This asserts both halves: invisible to the user (opacity 0), yet a fully
 * working, laid-out, clickable, painting window.
 */

const projectRoot = process.cwd()

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

async function windowState(app: ElectronApplication) {
  return app.evaluate(({ BrowserWindow }) => {
    const win = BrowserWindow.getAllWindows()[0]
    return { isVisible: win.isVisible(), opacity: win.getOpacity() }
  })
}

test('ATLAS_HIDDEN_WINDOW=1 shows a fully transparent window that still lays out and clicks', async () => {
  const app = await electron.launch({
    args: ['.'],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1', ATLAS_HIDDEN_WINDOW: '1' },
  })
  try {
    const page: Page = await app.firstWindow()
    await expect(page.getByRole('button', { name: 'Load Sample Document' })).toBeVisible({ timeout: 20_000 })
    await expect.poll(async () => (await windowState(app)).isVisible, { timeout: 15_000 }).toBe(true)

    // Shown (so Chromium keeps compositing) but invisible to the user.
    expect((await windowState(app)).opacity).toBe(0)

    // ...and still a real, interactive window: a click that depends on
    // Playwright's animation-frame stability check must succeed, and the
    // resulting panel must report real geometry.
    await page.getByRole('button', { name: 'New', exact: true }).click()
    const menu = page.locator('.dropdown__menu')
    await expect(menu).toBeVisible()
    const box = (await menu.boundingBox())!
    expect(box.width).toBeGreaterThan(0)
    expect(box.height).toBeGreaterThan(0)

    // The compositor is genuinely producing pixels, not a blank surface.
    const shot = await page.screenshot()
    expect(shot.byteLength).toBeGreaterThan(1000)
  } finally {
    kill(app)
  }
})

test('without the flag the window is shown normally, fully opaque', async () => {
  const app = await electron.launch({
    args: ['.'],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  try {
    const page: Page = await app.firstWindow()
    await expect(page.getByRole('button', { name: 'Load Sample Document' })).toBeVisible({ timeout: 20_000 })
    await expect.poll(async () => (await windowState(app)).isVisible, { timeout: 15_000 }).toBe(true)
    expect((await windowState(app)).opacity).toBe(1)
  } finally {
    kill(app)
  }
})
