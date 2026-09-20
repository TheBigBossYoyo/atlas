/**
 * A11Y focus-management pass 4 — two gaps the earlier passes (focus traps,
 * WCAG AA contrast, the keyboard-only journey) left:
 *
 * - A11Y-2: `PresenterView` only captured/restored focus on mount/unmount,
 *   with no Tab handling at all — Tab walked straight out into the document
 *   (here, the slide toolbar/editor) behind it. That was worse than it
 *   sounds because `requestFullscreen()` can be silently refused
 *   (`.catch(() => undefined)`), which this test forces deterministically
 *   rather than relying on whatever this CI environment's real window
 *   manager happens to grant.
 * - A11Y-3: `TabBar`'s close button had no focus handling — closing the
 *   focused tab unmounted it with nowhere for focus to go, dropping it to
 *   `<body>` so the next Tab press started over from the top of the window.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'

import { buildEditableDeck } from '../../src/viewers/slides/pptx/editing/__tests__/editableDeck'

const projectRoot = process.cwd()

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

async function activeElementDescription(page: Page): Promise<string> {
  return page.evaluate(() => {
    const el = document.activeElement
    if (!el || el === document.body) return 'body'
    const role = el.getAttribute('role') ?? el.tagName.toLowerCase()
    const name = el.getAttribute('aria-label') ?? el.getAttribute('title') ?? el.textContent?.trim().slice(0, 30) ?? ''
    return `${role}:${name}`
  })
}

test('presenter view traps Tab within itself even when requestFullscreen is refused', async () => {
  const deckFile = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-focus-presenter-')), 'deck.pptx')
  fs.writeFileSync(deckFile, Buffer.from(await buildEditableDeck()))

  const app = await electron.launch({
    args: ['.', deckFile],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1', ATLAS_HIDDEN_WINDOW: '1' },
  })
  try {
    const page = await app.firstWindow()
    await page.waitForSelector('.slide-deck__main .slide-edit__layer', { timeout: 20_000 })
    await page.waitForTimeout(500)

    // Force every fullscreen request to be refused, deterministically —
    // regardless of whether this CI window manager would actually grant one.
    // This is exactly the condition A11Y-2 fixed: PresenterView used to have
    // no Tab handling at all, so a refused request silently left Tab free to
    // walk out of the dialog.
    await page.evaluate(() => {
      HTMLElement.prototype.requestFullscreen = () =>
        Promise.reject(new DOMException('denied for test', 'NotAllowedError'))
    })

    const presenterTrigger = page.getByRole('button', { name: 'Presenter view' })
    await presenterTrigger.click()
    const dialog = page.getByRole('dialog', { name: 'Presenter view' })
    await expect(dialog).toBeVisible({ timeout: 10_000 })

    // The refusal actually took effect.
    await expect
      .poll(() => page.evaluate(() => document.fullscreenElement === null), { timeout: 5_000 })
      .toBe(true)

    const nextButton = page.getByRole('button', { name: 'Next slide' })
    const prevButton = page.getByRole('button', { name: 'Previous slide' })
    const exitButton = page.getByRole('button', { name: 'Exit presenter view' })
    // First slide of a 2-slide deck: Previous is genuinely disabled — the
    // same disabled-control case A11Y-1 was about, here for free.
    await expect(prevButton).toBeDisabled()
    await expect(nextButton).toBeEnabled()
    await expect(exitButton).toBeEnabled()

    // A real Tab walk (native browser focus movement, unlike jsdom) must
    // never land outside `.presenter-view` — that would be Tab escaping
    // into the slide toolbar/editor sitting right behind this overlay.
    await nextButton.focus()
    for (let i = 0; i < 8; i++) {
      await page.keyboard.press('Tab')
      const insideDialog = await page.evaluate(
        () => !!document.activeElement && !!document.activeElement.closest('.presenter-view'),
      )
      expect(insideDialog, `Tab press ${i + 1} escaped the presenter view (now on ${await activeElementDescription(page)})`).toBe(true)
    }

    // Escape exits back to the deck, restoring focus to the trigger.
    await page.keyboard.press('Escape')
    await expect(dialog).toHaveCount(0)
    await expect(presenterTrigger).toBeFocused()
  } finally {
    kill(app)
  }
})

test('closing a tab by keyboard moves focus to a neighbouring tab, never to <body>', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-focus-tabclose-'))
  const first = path.join(dir, 'first.md')
  const second = path.join(dir, 'second.md')
  const third = path.join(dir, 'third.md')
  fs.writeFileSync(first, '# First\n')
  fs.writeFileSync(second, '# Second\n')
  fs.writeFileSync(third, '# Third\n')

  const app = await electron.launch({
    args: ['.', first],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1', ATLAS_HIDDEN_WINDOW: '1' },
  })
  try {
    const page = await app.firstWindow()
    await expect(page.getByRole('tab', { name: 'first.md' })).toBeVisible({ timeout: 20_000 })

    // Open second.md and third.md through the real Open flow (dialog
    // answered in the main process), same as tabs.spec.ts.
    await app.evaluate(async ({ dialog }, filePath) => {
      dialog.showOpenDialog = (async () => ({ canceled: false, filePaths: [filePath] })) as typeof dialog.showOpenDialog
    }, second)
    await page.getByRole('button', { name: 'Open', exact: true }).click()
    await expect(page.getByRole('tab', { name: 'second.md' })).toBeVisible({ timeout: 20_000 })

    await app.evaluate(async ({ dialog }, filePath) => {
      dialog.showOpenDialog = (async () => ({ canceled: false, filePaths: [filePath] })) as typeof dialog.showOpenDialog
    }, third)
    await page.getByRole('button', { name: 'Open', exact: true }).click()
    await expect(page.getByRole('tab', { name: 'third.md' })).toBeVisible({ timeout: 20_000 })

    // ---- Close the MIDDLE tab (second.md) purely by keyboard: Tab lands on
    // its close button (simulated by focusing it directly, the same way a
    // real Tab walk would — see keyboard-only.spec.ts for the equivalent
    // pattern), then Enter activates it. No mouse click anywhere.
    const closeSecond = page.getByRole('button', { name: 'Close second.md' })
    await closeSecond.focus()
    await page.keyboard.press('Enter')

    await expect(page.getByRole('tab', { name: 'second.md' })).toHaveCount(0)
    // Never <body> — the previous close button unmounted right under focus.
    expect(await activeElementDescription(page)).not.toBe('body')
    // Lands on the next tab's own close button (third.md), so a keyboard
    // user can keep closing tabs in place.
    await expect(page.getByRole('button', { name: 'Close third.md' })).toBeFocused()

    // ---- Close third.md too (now the LAST tab) — focus should fall back to
    // the remaining neighbour, first.md's close button.
    await page.keyboard.press('Enter')
    await expect(page.getByRole('tab', { name: 'third.md' })).toHaveCount(0)
    expect(await activeElementDescription(page)).not.toBe('body')
    await expect(page.getByRole('button', { name: 'Close first.md' })).toBeFocused()

    // ---- Close the sole remaining tab — TabBar itself unmounts; the app
    // must not crash or hang (where focus lands next is the shell's concern,
    // not TabBar's — this only proves the keyboard close path is safe here).
    await page.keyboard.press('Enter')
    await expect(page.getByRole('tab', { name: 'first.md' })).toHaveCount(0)
    await expect(page.getByRole('button', { name: 'Load Sample Document' })).toBeVisible({ timeout: 10_000 })
  } finally {
    kill(app)
  }
})
