import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'

/**
 * A11Y pass 3 — a real keyboard-only journey through the app: open, edit,
 * save, switch tabs, find, change theme/language, and open/close a panel and
 * a dialog, none of it using `.click()`. `locator.focus()` moves focus the
 * way a screen-reader user's Tab-then-arrow browsing would land on a given
 * control (it dispatches no pointer events), and every actual INTERACTION
 * from there on is a real `keyboard.press`/`keyboard.type`. A leading real
 * Tab-walk (see the first test) additionally proves the toolbar itself has
 * no dead end for a user who only has Tab to get around.
 */

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

test('keyboard-only: Tab walks through the toolbar with no dead end', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-kbd-'))
  const first = path.join(dir, 'first.md')
  fs.writeFileSync(first, '# Keyboard-only\n\nHello.\n')

  const app = await electron.launch({
    args: ['.', first],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  try {
    const page = await app.firstWindow()
    await expect(page.getByRole('tab', { name: 'first.md' })).toBeVisible({ timeout: 20_000 })

    // Walk 25 Tab presses from a clean slate and make sure focus never gets
    // stuck repeating the same element (a real dead end) and never falls
    // back to <body> mid-walk (focus escaping the document entirely).
    await page.evaluate(() => (document.activeElement as HTMLElement | null)?.blur())
    const seen: string[] = []
    for (let i = 0; i < 25; i++) {
      await page.keyboard.press('Tab')
      seen.push(await activeElementDescription(page))
    }
    const stuckRun = seen.some((_, i) => i > 0 && seen[i] === seen[i - 1])
    expect(stuckRun, `focus repeated the same element back-to-back:\n${seen.join('\n')}`).toBe(false)
    // The walk starts back at <body> once (the very first Tab, from nothing
    // focused) but must not return there again afterwards.
    expect(seen.filter((s) => s === 'body').length, seen.join('\n')).toBeLessThanOrEqual(1)
  } finally {
    kill(app)
  }
})

test('keyboard-only: edit, save, switch tabs by arrow key, find, theme, language, and a dialog', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-kbd-'))
  const first = path.join(dir, 'first.md')
  const second = path.join(dir, 'second.md')
  fs.writeFileSync(first, '# First\n\nOriginal text.\n')
  fs.writeFileSync(second, '# Second\n\nFindable needle here.\n')

  const app = await electron.launch({
    args: ['.', first],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  try {
    const page = await app.firstWindow()
    await expect(page.getByRole('tab', { name: 'first.md' })).toBeVisible({ timeout: 20_000 })

    // ---- Open a second document (Ctrl+O; the OS dialog is stubbed) -------
    await app.evaluate(async ({ dialog }, filePath) => {
      dialog.showOpenDialog = (async () => ({ canceled: false, filePaths: [filePath] })) as typeof dialog.showOpenDialog
    }, second)
    await page.keyboard.press('Control+o')
    await expect(page.getByRole('tab', { name: 'second.md' })).toHaveAttribute('aria-selected', 'true', {
      timeout: 20_000,
    })

    // ---- Switch tabs with the ARIA tablist keyboard pattern (no click) ---
    // Regression coverage for TabBar's roving-tabindex fix: focus the
    // (already-selected) second.md tab, then arrow to first.md and activate
    // it with Enter — exactly the sequence a screen-reader user would use.
    await page.getByRole('tab', { name: 'second.md' }).focus()
    await page.keyboard.press('ArrowLeft')
    await expect(page.getByRole('tab', { name: 'first.md' })).toBeFocused()
    await page.keyboard.press('Enter')
    await expect(page.getByRole('tab', { name: 'first.md' })).toHaveAttribute('aria-selected', 'true')
    await expect(page.locator('.markdown-body')).toContainText('Original text')

    // ---- Edit + save (Editor mode, typed text, Ctrl+S) --------------------
    await page.getByRole('button', { name: 'Editor' }).focus()
    await page.keyboard.press('Enter')
    const textarea = page.locator('.editor-panel__textarea')
    await expect(textarea).toBeVisible()
    await textarea.focus()
    await page.keyboard.press('Control+End')
    await page.keyboard.type('\n\nAdded by keyboard.')
    await page.keyboard.press('Control+s')
    await expect
      .poll(() => fs.readFileSync(first, 'utf8'), { timeout: 10_000 })
      .toContain('Added by keyboard.')

    // ---- Find (Ctrl+F) on the OTHER tab, with the "N of M" live count -----
    await page.getByRole('tab', { name: 'second.md' }).focus()
    await page.keyboard.press('Enter')
    await expect(page.getByRole('tab', { name: 'second.md' })).toHaveAttribute('aria-selected', 'true')
    // Find scans the rendered preview, not the raw-editor textarea — the
    // view mode carried over from editing first.md above.
    await page.getByRole('button', { name: 'Preview' }).focus()
    await page.keyboard.press('Enter')
    await expect(page.locator('.markdown-body')).toContainText('Findable needle')
    await page.keyboard.press('Control+f')
    const findInput = page.getByPlaceholder('Search in document...')
    await expect(findInput).toBeFocused()
    await page.keyboard.type('Findable')
    await expect(page.locator('.search-overlay__count')).toHaveText('1 of 1', { timeout: 10_000 })
    await page.keyboard.press('Escape')
    await expect(findInput).toHaveCount(0)

    // ---- Theme menu: open, arrow to an item, activate, Escape restores focus
    const themeTrigger = page.getByRole('button', { name: 'Theme', exact: true })
    await themeTrigger.focus()
    await page.keyboard.press('Enter')
    const themeMenu = page.getByRole('list', { name: /theme/i }).or(page.locator('.dropdown__menu').first())
    await expect(themeMenu).toBeVisible()
    const initialTheme = await page.evaluate(() => document.documentElement.getAttribute('data-theme'))
    // Tab through the plain-button list (no arrow-key pattern implemented,
    // see ThemeMenu.tsx's own UX-14 note) until landing on an item that
    // ISN'T the current theme, whichever position that happens to be in.
    let landedOnDifferentTheme = false
    for (let i = 0; i < 5; i++) {
      await page.keyboard.press('Tab')
      const isCurrent = await page.evaluate(() => document.activeElement?.getAttribute('aria-current'))
      if (isCurrent !== 'true') {
        landedOnDifferentTheme = true
        break
      }
    }
    expect(landedOnDifferentTheme).toBe(true)
    await page.keyboard.press('Enter')
    await expect
      .poll(() => page.evaluate(() => document.documentElement.getAttribute('data-theme')))
      .not.toBe(initialTheme)
    // Selecting a theme closes the menu and returns focus to its trigger.
    await expect(themeTrigger).toBeFocused()

    // ---- Language menu: switch to French and back, verifying a real label
    const languageTrigger = page.getByRole('button', { name: 'Language', exact: true })
    // Its own aria-label is translated the moment French is selected, so grab
    // a handle now rather than re-querying by (now-stale) English name below.
    const languageTriggerHandle = await languageTrigger.elementHandle()
    await languageTrigger.focus()
    await page.keyboard.press('Enter')
    await expect(page.locator('.dropdown__menu').first()).toBeVisible()
    // First item is English, second is French (see LanguageMenu.tsx's OPTIONS).
    await page.keyboard.press('Tab')
    await page.keyboard.press('Tab')
    await page.keyboard.press('Enter')
    await expect
      .poll(() => page.evaluate((el) => document.activeElement === el, languageTriggerHandle))
      .toBe(true)
    await expect(page.getByRole('button', { name: 'Éditeur' })).toBeVisible({ timeout: 10_000 })

    // Switch back to English so the rest of this test's locators keep
    // working — focus is already on the trigger (just confirmed above), so
    // no need to re-locate it by its (now-French) accessible name.
    await page.keyboard.press('Enter')
    await page.keyboard.press('Tab')
    await page.keyboard.press('Enter')
    await expect(page.getByRole('button', { name: 'Editor' })).toBeVisible({ timeout: 10_000 })

    // ---- Shortcuts dialog: Ctrl+/, focus trapped inside, Escape restores --
    const shortcutsTrigger = page.getByRole('button', { name: 'Keyboard shortcuts', exact: true })
    await shortcutsTrigger.focus()
    await page.keyboard.press('Control+/')
    const dialog = page.getByRole('dialog')
    await expect(dialog).toBeVisible()
    await expect(dialog.locator(':focus')).toHaveCount(1) // focus moved inside the dialog
    await page.keyboard.press('Escape')
    await expect(dialog).toHaveCount(0)
    await expect(shortcutsTrigger).toBeFocused()

    // ---- Sidebar panel: Ctrl+B toggles it without losing focus ------------
    const sidebarToggle = page.getByRole('button', { name: 'Toggle sidebar', exact: true })
    await sidebarToggle.focus()
    await page.keyboard.press('Control+b')
    await page.keyboard.press('Control+b')
    await expect(sidebarToggle).toBeFocused()
  } finally {
    kill(app)
  }
})
