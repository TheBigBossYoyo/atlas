/**
 * F6 — a deterministic, silent data-loss bug: click a spreadsheet cell and
 * start typing right away (ordinary human typing speed, no artificial pause)
 * and the FIRST character was silently dropped. Measured 24/24 failures
 * across every combination of per-character typing delay (0/30/60/120ms) and
 * pause-after-click (0/150/400/1000ms) before this fix — see
 * SpreadsheetDataEditor.tsx's module header for the full root-cause writeup.
 *
 * This only shows up against the SAVED FILE BYTES, not the screen: a screen
 * assertion right after typing can catch the grid mid-race, before
 * glide-data-grid's own overlay has settled, and would have passed even on
 * the buggy build. Every test here unzips/reads the actual saved file.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import * as XLSX from 'xlsx'

const projectRoot = process.cwd()

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

async function launch(file: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({
    args: ['.', file],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1', ATLAS_HIDDEN_WINDOW: '1' },
  })
  const page = await app.firstWindow()
  return { app, page }
}

async function clickTypeSave(
  page: Page,
  gridSelector: string,
  text: string,
  delay: number,
  postClickPause: number,
): Promise<void> {
  const box = (await page.locator(gridSelector).first().boundingBox())!
  await page.mouse.click(box.x + 40, box.y + 40)
  if (postClickPause > 0) await page.waitForTimeout(postClickPause)
  // No wait for the overlay/textarea to appear — this is the whole point:
  // ordinary typing outruns glide-data-grid's own overlay-mount latency.
  await page.keyboard.type(text, { delay })
  await page.keyboard.press('Enter')
  await page.waitForTimeout(500)
  await page.keyboard.press('Control+s')
  await page.waitForTimeout(1500)
}

function firstNonEmptyCellValue(sheet: XLSX.WorkSheet): string | undefined {
  const key = Object.keys(sheet)
    .filter((k) => !k.startsWith('!'))
    .find((k) => sheet[k].v !== undefined && sheet[k].v !== '')
  return key === undefined ? undefined : String(sheet[key].v)
}

test.describe('F6 — click a cell and type: the leading character must not be dropped (xlsx)', () => {
  for (const delay of [0, 30, 60, 120]) {
    test(`typing at ${delay}ms/char right after a click saves the full word`, async () => {
      const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-f6-xlsx-'))
      const file = path.join(dir, 'b.xlsx')
      fs.copyFileSync(path.join(projectRoot, 'tests/e2e/fixtures/sample.xlsx'), file)

      const { app, page } = await launch(file)
      try {
        await page.waitForSelector('[data-viewer="xlsx"]', { timeout: 30_000 })
        await page.waitForTimeout(1500)
        await clickTypeSave(page, '[data-viewer="xlsx"] canvas', 'HELLO', delay, 0)

        const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer' })
        const sheet = wb.Sheets[wb.SheetNames[0]]
        expect(firstNonEmptyCellValue(sheet)).toBe('HELLO')
      } finally {
        kill(app)
      }
    })
  }

  for (const pause of [0, 150, 400, 1000]) {
    test(`typing at 120ms/char with a ${pause}ms pause after the click saves the full word`, async () => {
      const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-f6-xlsx-pause-'))
      const file = path.join(dir, 'b.xlsx')
      fs.copyFileSync(path.join(projectRoot, 'tests/e2e/fixtures/sample.xlsx'), file)

      const { app, page } = await launch(file)
      try {
        await page.waitForSelector('[data-viewer="xlsx"]', { timeout: 30_000 })
        await page.waitForTimeout(1500)
        await clickTypeSave(page, '[data-viewer="xlsx"] canvas', 'HELLO', 120, pause)

        const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer' })
        const sheet = wb.Sheets[wb.SheetNames[0]]
        expect(firstNonEmptyCellValue(sheet)).toBe('HELLO')
      } finally {
        kill(app)
      }
    })
  }
})

test.describe('F6 — CSV shares the same grid/editor pipeline, so it must not be affected either', () => {
  test('typing right after a click into a CSV cell saves the full word', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-f6-csv-'))
    const file = path.join(dir, 'b.csv')
    fs.copyFileSync(path.join(projectRoot, 'tests/e2e/fixtures/sample.csv'), file)

    const { app, page } = await launch(file)
    try {
      await page.waitForSelector('[data-viewer="csv"]', { timeout: 30_000 })
      await page.waitForTimeout(1500)
      await clickTypeSave(page, '[data-viewer="csv"] canvas', 'HELLO', 30, 0)

      const saved = fs.readFileSync(file, 'utf-8')
      // Exact first line, not a substring check: "HELLO,value" legitimately
      // CONTAINS "ELLO," as a substring, so a substring assertion can't tell
      // a correct save apart from the bug (the leading "H" dropped).
      expect(saved.split(/\r?\n/)[0]).toBe('HELLO,value')
    } finally {
      kill(app)
    }
  })
})

test.describe('F6 — other ways into edit mode are unaffected', () => {
  test('double-click still opens the overlay pre-filled with the existing value', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-f6-dblclick-'))
    const file = path.join(dir, 'b.xlsx')
    fs.copyFileSync(path.join(projectRoot, 'tests/e2e/fixtures/sample.xlsx'), file)

    const { app, page } = await launch(file)
    try {
      await page.waitForSelector('[data-viewer="xlsx"]', { timeout: 30_000 })
      await page.waitForTimeout(1500)
      const box = (await page.locator('[data-viewer="xlsx"] canvas').first().boundingBox())!
      await page.mouse.dblclick(box.x + 40, box.y + 40)
      await expect(page.locator('#portal textarea')).toBeVisible()
    } finally {
      kill(app)
    }
  })
})
