/**
 * Throttled edit/commit/save matrix for the spreadsheet grid.
 *
 * WHY THIS EXISTS. F6, F6b and F6c were three separate silent-data-loss bugs in the
 * same click -> type -> commit -> save path, and every one was invisible at full
 * speed:
 *   - F6  (3.7.0): the first character after a click was dropped.
 *   - F6b (3.7.x): the F6 fix held only on a fast machine.
 *   - F6c (3.8.0): the SECOND cell edit of a session was discarded and the Ctrl+S
 *     after it silently wrote nothing — 5/6 at 8x throttle, 1/6 unthrottled.
 *
 * Each was found by a user-visible loss or a one-off CI red, then reproduced by hand
 * with `Emulation.setCPUThrottlingRate`. The existing throttled coverage
 * (`spreadsheet-keystroke-seed.spec.ts`, `spreadsheet-second-cell-edit.spec.ts`) pins
 * the three specific scenarios already known to have broken. Everything else in the
 * path is tested at full speed only — where F6c showed up 1 time in 6.
 *
 * So this walks the dimensions those three bugs fell through, under throttling,
 * asserting the SAVED FILE BYTES. Screen assertions are deliberately avoided: as
 * `spreadsheet-keystroke-seed.spec.ts`'s header puts it, a screen assertion right
 * after typing "can catch the grid mid-race... and would have passed even on the
 * buggy build". Bytes cannot be fooled that way.
 *
 * DIMENSIONS
 *   - which edit in the session: the first, or a later one (F6c lived here)
 *   - how the edit is committed: Enter, Tab, or clicking a different cell
 *   - what is typed: plain text, a number, a percent, an ISO date, a formula
 *     (these take different paths through the passthrough writer)
 *   - CPU rate: 1x and 8x
 *
 * Kept to a bounded set rather than a full cross-product: every case launches a real
 * Electron app, so an exhaustive matrix would cost more wall-clock than it is worth.
 * Each case below covers a combination none of the existing specs reach.
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
  await page.waitForSelector('.spreadsheet-viewer__grid canvas', { timeout: 30_000 })
  await page.waitForTimeout(1000)
  return { app, page }
}

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

async function throttleCpu(page: Page, rate: number): Promise<void> {
  if (rate <= 1) return
  const cdp = await page.context().newCDPSession(page)
  await cdp.send('Emulation.setCPUThrottlingRate', { rate })
}

/** A two-column sheet with a header row, so row 1 (index 1) is always free to edit. */
function makeWorkbook(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-matrix-'))
  const file = path.join(dir, 'matrix.xlsx')
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['H1', 'H2', 'H3']]), 'Data')
  fs.writeFileSync(file, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer)
  return file
}

/**
 * Scrolls the grid back to the top-left before a click, by setting the scroll
 * position directly and waiting for it to settle.
 *
 * Load-bearing, and the mechanism matters. Committing with Enter moves the selection
 * down, and the sheet carries an Excel-like blank margin past the data (USR-17), so
 * the grid scrolls between edits — after which clicking the same pixel lands on a
 * different row. The first draft of this file did not reset at all, and values saved
 * correctly into A60/B60/C40 instead of row 2: a test-coordinate bug whose assertion
 * message looks exactly like data loss.
 *
 * The second draft reset with `Control+Home`, which was worse: that both scrolls AND
 * moves the selection, so a click 150ms later raced the scroll settling and the next
 * keystroke went to A1 — it overwrote the header row. That looked like a product bug
 * and is not one: a realistic sequence (click, ArrowDown, click elsewhere, type at
 * once) puts the text in the right cell 6 times out of 6 at both 1x and 8x, probed
 * separately.
 *
 * So this touches neither the keyboard nor the selection: it sets `scrollTop`/
 * `scrollLeft` on glide's scroller and polls until they have actually taken, leaving
 * the click -> type race (the whole point of these cases) as the only thing in play.
 */
async function resetViewport(page: Page): Promise<void> {
  await page.evaluate(() => {
    const scroller = document.querySelector('.dvn-scroller')
    if (scroller !== null) {
      scroller.scrollTop = 0
      scroller.scrollLeft = 0
    }
  })
  await page.waitForFunction(
    () => {
      const scroller = document.querySelector('.dvn-scroller')
      return scroller === null || (scroller.scrollTop === 0 && scroller.scrollLeft === 0)
    },
    null,
    { timeout: 10_000 },
  )
}

async function clickCell(page: Page, col: number, row: number): Promise<void> {
  const box = (await page.locator('.spreadsheet-viewer__grid canvas').first().boundingBox())!
  await page.mouse.click(box.x + 48 + COLUMN_WIDTH * col + COLUMN_WIDTH / 2, box.y + HEADER_HEIGHT + ROW_HEIGHT * row + ROW_HEIGHT / 2)
}

type Commit = 'Enter' | 'Tab' | 'click-away'

/**
 * What has focus, and how many edit overlays exist — the two facts that told F6c
 * apart from a slow machine. Mirrors `spreadsheet-second-cell-edit.spec.ts`'s
 * `focusState`, for the same reason: a bare timeout on the overlay wait says
 * nothing about WHY, and this path has produced three real data-loss bugs whose
 * only outward sign was a timeout.
 */
async function focusState(page: Page): Promise<string> {
  return page.evaluate(() => {
    const active = document.activeElement
    const name = active === null ? 'null' : active.tagName + (active instanceof HTMLElement && active.dataset.testid !== undefined ? `[${active.dataset.testid}]` : '')
    const text = document.querySelector('#portal textarea')
    return `focus=${name} overlays=${document.querySelectorAll('#portal textarea').length}${
      text instanceof HTMLTextAreaElement ? ` overlayValue=${JSON.stringify(text.value)}` : ''
    }`
  })
}

/**
 * Types `text` into (col, row) exactly as a user does — click, then type with no
 * pause — and commits it the requested way. No screen assertions beyond waiting for
 * the overlay to appear and then go, which is synchronisation, not verification.
 *
 * Both waits carry the focus/overlay state in their failure message. Added after
 * the second full-suite run saw the unthrottled percent/date/formula case time out
 * on the closing wait once (8/8 in isolation, 14/14 at file scope, so not
 * reproducible on demand) — without this, the next occurrence would be as
 * undiagnosable as that one was.
 */
async function editCell(page: Page, col: number, row: number, text: string, commit: Commit): Promise<void> {
  await resetViewport(page)
  await clickCell(page, col, row)
  await page.keyboard.type(text.slice(0, 1))
  try {
    await page.waitForSelector('#portal textarea', { timeout: 20_000 })
  } catch (cause) {
    throw new Error(`the edit overlay never opened for ${text} at (${col},${row}): ${await focusState(page)}`, { cause })
  }
  if (text.length > 1) await page.keyboard.type(text.slice(1), { delay: 20 })

  if (commit === 'click-away') {
    // Commit by moving to another cell with the mouse — a path neither Enter nor
    // Tab exercises, and the one a user takes most often without thinking.
    await clickCell(page, col === 0 ? 2 : 0, row + 1)
  } else {
    await page.keyboard.press(commit)
  }
  try {
    await page.waitForFunction(() => document.querySelectorAll('#portal textarea').length === 0, null, { timeout: 20_000 })
  } catch (cause) {
    throw new Error(
      `the edit overlay never closed after committing ${text} at (${col},${row}) with ${commit}: ${await focusState(page)}`,
      { cause },
    )
  }
}

/** Saves and waits for the write itself, so "never saved" and "saved wrong" fail differently. */
async function save(page: Page, file: string): Promise<void> {
  const before = fs.statSync(file).mtimeMs
  await page.keyboard.press('Control+s')
  await expect
    .poll(() => fs.statSync(file).mtimeMs, { message: 'the file was never saved', timeout: 25_000 })
    .toBeGreaterThan(before)
}

/** Every non-empty cell, so a failure shows where the text actually went. */
function cells(file: string): Record<string, string> {
  const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer', cellNF: true })
  const sheet = wb.Sheets.Data
  const out: Record<string, string> = {}
  for (const k of Object.keys(sheet)) {
    if (!k.startsWith('!') && sheet[k].v !== undefined && sheet[k].v !== '') out[k] = String(sheet[k].v)
  }
  return out
}

for (const rate of [1, 8]) {
  const label = rate > 1 ? ` (renderer throttled ${rate}x)` : ''

  // F6c was specifically about the edit AFTER a commit. These walk each commit
  // gesture through two consecutive edits, which is the shape that broke.
  for (const commit of ['Enter', 'Tab', 'click-away'] as Commit[]) {
    test(`two consecutive edits both survive, committed with ${commit}${label}`, async () => {
      const file = makeWorkbook()
      const { app, page } = await launch(file)
      try {
        await throttleCpu(page, rate)
        await editCell(page, 0, 1, 'first', commit)
        await editCell(page, 1, 1, 'second', commit)
        await save(page, file)
        const saved = cells(file)
        expect(saved, `saved: ${JSON.stringify(saved)}`).toMatchObject({ A2: 'first', B2: 'second' })
      } finally {
        kill(app)
      }
    })
  }

  // Three in a row: if the Nth edit depends on state left by the N-1th, two is not
  // always enough to show it.
  test(`three consecutive edits all survive${label}`, async () => {
    const file = makeWorkbook()
    const { app, page } = await launch(file)
    try {
      await throttleCpu(page, rate)
      await editCell(page, 0, 1, 'one', 'Enter')
      await editCell(page, 1, 1, 'two', 'Enter')
      await editCell(page, 2, 1, 'three', 'Enter')
      await save(page, file)
      const saved = cells(file)
      expect(saved, `saved: ${JSON.stringify(saved)}`).toMatchObject({ A2: 'one', B2: 'two', C2: 'three' })
    } finally {
      kill(app)
    }
  })

  // Value KINDS take different paths through the xlsx passthrough writer (percent and
  // date get a number format; a formula is stored as one). A focus/commit bug can
  // therefore show for one kind and not another.
  test(`a percent, an ISO date and a formula all survive consecutive edits${label}`, async () => {
    const file = makeWorkbook()
    const { app, page } = await launch(file)
    try {
      await throttleCpu(page, rate)
      await editCell(page, 0, 1, '50%', 'Enter')
      await editCell(page, 1, 1, '2024-03-14', 'Enter')
      await editCell(page, 2, 1, '=1+1', 'Enter')
      await save(page, file)
      const saved = cells(file)
      expect(saved.A2, `saved: ${JSON.stringify(saved)}`).toBe('0.5')
      expect(saved.B2, `saved: ${JSON.stringify(saved)}`).toBe('45365') // 2024-03-14's serial
      expect(saved.C2, `saved: ${JSON.stringify(saved)}`).toBe('2')
    } finally {
      kill(app)
    }
  })

  // Re-editing the SAME cell: the overlay opens over a cell that already has a value,
  // which is a different glide code path from editing an empty one.
  test(`re-editing the same cell keeps the newer value${label}`, async () => {
    const file = makeWorkbook()
    const { app, page } = await launch(file)
    try {
      await throttleCpu(page, rate)
      await editCell(page, 0, 1, 'before', 'Enter')
      await editCell(page, 0, 1, 'after', 'Enter')
      await save(page, file)
      const saved = cells(file)
      expect(saved, `saved: ${JSON.stringify(saved)}`).toMatchObject({ A2: 'after' })
    } finally {
      kill(app)
    }
  })

  // F6b's second half was Ctrl+S overtaking an edit. That is pinned for the FIRST
  // edit; this is the same race after a previous commit has already happened.
  test(`Ctrl+S immediately after a second edit saves that edit${label}`, async () => {
    const file = makeWorkbook()
    const { app, page } = await launch(file)
    try {
      await throttleCpu(page, rate)
      await editCell(page, 0, 1, 'first', 'Enter')
      // No settle wait: save the instant the second edit is committed.
      await resetViewport(page)
      await clickCell(page, 1, 1)
      await page.keyboard.type('s')
      await page.waitForSelector('#portal textarea', { timeout: 20_000 })
      await page.keyboard.type('econd', { delay: 20 })
      await page.keyboard.press('Enter')
      await save(page, file)
      const saved = cells(file)
      expect(saved, `saved: ${JSON.stringify(saved)}`).toMatchObject({ A2: 'first', B2: 'second' })
    } finally {
      kill(app)
    }
  })
}
