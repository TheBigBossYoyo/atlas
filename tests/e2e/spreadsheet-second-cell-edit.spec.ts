/**
 * F6c · regression test for a HIGH-severity silent data-loss bug, fixed 2026-09-27.
 *
 * The bug: edit a cell and commit it, then click a DIFFERENT cell and type at
 * once — the second edit was silently lost AND the Ctrl+S after it silently did
 * not save. The file's `mtime` was unchanged 20s later and nothing was shown to
 * the user. Reachable by hand, and measured at 5/6 with the renderer throttled 8x
 * and 1/6 unthrottled, so a slow machine was not a precondition.
 *
 * MECHANISM. Committing an edit leaves DOM focus on a `<td>` of glide-data-grid's
 * own accessibility table rather than on `canvas[data-testid="data-grid-canvas"]`.
 * `SpreadsheetDataEditor`'s `holdKeysOnMouseDown` then skipped arming `KeyHold`
 * whenever `grid.contains(document.activeElement)` — which that `<td>` satisfies —
 * so nothing was held, and the keys went straight to an element that does nothing
 * with them. The overlay opened (glide's `<td>` handler sees the key) but never
 * received the text, so Enter could not commit it; and because `KeyHold` holds
 * chords while an edit is in flight, the Ctrl+S queued behind an edit that could
 * now never commit and was never replayed.
 *
 * THE FIX, in two halves that only work together — ask whether the CANVAS has
 * focus, not whether the grid *contains* focus:
 *   1. `SpreadsheetDataEditor.holdKeysOnMouseDown` arms `KeyHold` unless the canvas
 *      itself (or an overlay in `#portal`) already has focus.
 *   2. `KeyHold.release` focuses the canvas unless focus is held outside the grid.
 *
 * Proven by a control experiment rather than assumed: both halves gives 6/6 at 8x;
 * reverting only half 1 and keeping half 2 gives 6/6 FAILING. Three earlier attempts
 * that changed `KeyHold` alone could not have worked, because half 1 meant
 * `KeyHold.start` was never called in the first place.
 *
 * Throttled at 8x deliberately: the bug reproduced only 1 in 6 at full speed, so an
 * unthrottled test would not reliably catch a regression. This is the same
 * technique `spreadsheet-keystroke-seed.spec.ts` uses for F6/F6b.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import * as XLSX from 'xlsx'

const projectRoot = process.cwd()
const HEADER_HEIGHT = 36
const ROW_HEIGHT = 32
const COLUMN_WIDTH = 120

async function launch(file: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({ args: ['.', file], cwd: projectRoot, env: { ...process.env, CI: '1', PLAYWRIGHT: '1' } })
  const page = await app.firstWindow()
  await page.waitForSelector('.spreadsheet-viewer__grid canvas', { timeout: 20_000 })
  await page.waitForTimeout(800)
  return { app, page }
}

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

/** Simulates a slow machine; the bug reproduces unthrottled too, just less often. */
async function throttleCpu(page: Page, rate: number): Promise<void> {
  const cdp = await page.context().newCDPSession(page)
  await cdp.send('Emulation.setCPUThrottlingRate', { rate })
}

function tmpFile(name: string): string {
  return path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-f6c-')), name)
}

async function clickCell(page: Page, col: number, row: number): Promise<void> {
  const box = (await page.locator('.spreadsheet-viewer__grid canvas').first().boundingBox())!
  await page.mouse.click(box.x + 48 + COLUMN_WIDTH * col + COLUMN_WIDTH / 2, box.y + HEADER_HEIGHT + ROW_HEIGHT * row + ROW_HEIGHT / 2)
}

/** What actually has focus, plus how many edit overlays exist — the diagnosis above. */
async function focusState(page: Page): Promise<string> {
  return page.evaluate(() => {
    const el = document.activeElement
    const overlays = document.querySelectorAll('#portal textarea').length
    return `${el === null ? 'null' : el.tagName} (overlays: ${overlays})`
  })
}

function sheetCells(file: string): Record<string, string> {
  const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer', cellNF: true })
  const sheet = wb.Sheets[wb.SheetNames[0]]
  const cells: Record<string, string> = {}
  for (const k of Object.keys(sheet)) {
    if (!k.startsWith('!') && sheet[k].v !== undefined && sheet[k].v !== '') cells[k] = String(sheet[k].v)
  }
  return cells
}

test('F6c: a second cell edit is kept, and Ctrl+S after it actually saves', async () => {
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['H1', 'H2']]), 'Data')
  const file = tmpFile('two-cells.xlsx')
  fs.writeFileSync(file, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer)
  const mtimeBefore = fs.statSync(file).mtimeMs

  const { app, page } = await launch(file)
  const diag: string[] = []
  try {
    await throttleCpu(page, 8)

    await clickCell(page, 0, 1)
    await page.keyboard.type('5')
    await expect(page.locator('#portal textarea')).toHaveCount(1, { timeout: 20_000 })
    diag.push(`cell1 first char: ${await focusState(page)}`)
    await page.keyboard.type('0%', { delay: 20 })
    await page.keyboard.press('Enter')
    await expect(page.locator('#portal textarea')).toHaveCount(0, { timeout: 20_000 })
    diag.push(`cell1 Enter: ${await focusState(page)}`)

    await clickCell(page, 1, 1)
    await page.keyboard.type('2')
    await expect(page.locator('#portal textarea')).toHaveCount(1, { timeout: 20_000 })
    diag.push(`cell2 first char: ${await focusState(page)}`)
    await page.keyboard.type('024-03-14', { delay: 20 })
    await page.keyboard.press('Enter')
    await page.waitForTimeout(1000)
    diag.push(`cell2 Enter: ${await focusState(page)}`)

    await page.keyboard.press('Control+s')
    await expect
      .poll(() => fs.statSync(file).mtimeMs, { message: `the file was never saved. ${diag.join(' | ')}`, timeout: 20_000 })
      .toBeGreaterThan(mtimeBefore)

    expect(sheetCells(file), `${diag.join(' | ')}`).toMatchObject({ A2: '0.5', B2: '45365' })
  } finally {
    kill(app)
  }
})
