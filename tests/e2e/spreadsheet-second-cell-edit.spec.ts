/**
 * F6c · HIGH · SILENT DATA LOSS · OPEN — the second cell edit of a session can be
 * lost, and the save that follows it silently does nothing.
 *
 * Found 2026-09-26/27 while investigating why `spreadsheet-grid-fixes.spec.ts`'s
 * SHEET-4/5 test failed once in a full e2e run and passed in isolation. It is not
 * a flake. The reproduction is below and it is a user-reachable sequence:
 *
 *   1. Click a cell and start typing at once (no pause) — this part works.
 *   2. Press Enter to commit.
 *   3. Click a DIFFERENT cell and start typing at once.
 *   4. Press Enter, then Ctrl+S.
 *
 * Step 3's text never reaches the cell, step 4's Enter does not commit, and the
 * Ctrl+S never writes the file at all — `mtime` is unchanged 20s later. Nothing
 * tells the user; the app looks like it accepted the edit.
 *
 * Measured frequency (this laptop, `Emulation.setCPUThrottlingRate`, n=6 per rate):
 * 5/6 at 8x, 1/6 unthrottled. It reproduces unthrottled, so a slow or loaded
 * machine is not a precondition — throttling only makes it reliable.
 *
 * MECHANISM, from `document.activeElement` sampled at each step:
 *
 *   after cell1 first char : TEXTAREA  (portal textareas: 1)   <- correct
 *   after cell1 Enter      : TD        (portal textareas: 0)   <- focus on glide's
 *                                                                a11y table cell,
 *                                                                not the canvas
 *   after cell2 first char : TD        (portal textareas: 1)   <- overlay OPEN but
 *                                                                unfocused
 *   after cell2 Enter      : TD        (portal textareas: 1)   <- never committed
 *
 * So: committing an edit leaves DOM focus on a `<td>` of glide-data-grid's own
 * accessibility table rather than on `canvas[data-testid="data-grid-canvas"]`.
 * `KeyHold` (src/viewers/shared/keyHold.ts) then replays the held keystrokes to
 * `document.activeElement`, which is that `<td>`, so they do nothing. The overlay
 * opens (glide's own `<td>` handler sees the key) but never receives the text, so
 * Enter cannot commit it — and `KeyHold` holds chords while an edit is in flight,
 * so the Ctrl+S queues behind an edit that can now never commit and is never
 * replayed. That is why the file is not written rather than written with stale
 * text.
 *
 * APPROACHES TRIED AND REJECTED (all in `KeyHold.release`/`drain`; each was
 * measured, none fixed it — recorded so the next attempt does not repeat them):
 *   1. Guard on "is the canvas focused" instead of "does the grid contain focus".
 *      Logically right and necessary, but not sufficient: 8x went 3/3 fail -> 1/3.
 *   2. Also treat "focus moved to another element inside the grid" as not-moved-away.
 *      No measurable improvement (8x 2/3 fail at n=3).
 *   3. Re-assert canvas focus immediately before every replay, not once at release.
 *      Still 5/6 fail at 8x with n=6. glide appears to re-take focus after the
 *      replay too, not only between release and the first replay.
 * All three were reverted. The fix likely belongs in how the edit overlay hands
 * focus back on commit (`SpreadsheetDataEditor.tsx`), not in `KeyHold` — KeyHold is
 * downstream of a focus target that is already wrong.
 *
 * `test.fixme` rather than deleted or left failing: the repro is worth keeping
 * executable and ready to un-fixme, without turning CI red on every run. The
 * bug itself is tracked in the phase-4 backlog as F6c and in
 * docs/KNOWN_LIMITATIONS.md.
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

test.fixme('F6c: a second cell edit is kept, and Ctrl+S after it actually saves', async () => {
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
