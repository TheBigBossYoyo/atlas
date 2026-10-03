/**
 * SHEETFN-5 — a formula that crosses sheets, in the real app.
 *
 * The unit tests cover the evaluator and the document-level dependency graph.
 * This covers the part only the real app can show: that a cross-sheet formula
 * resolves when the workbook is OPENED (the viewer builds its document through
 * `createDocument`, which had to change for this), that editing the referenced
 * sheet updates the referring one, and that the formula survives a save.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import * as XLSX from 'xlsx'

const projectRoot = process.cwd()

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

/**
 * A two-sheet workbook where Summary!A1 is `=Data!A1*2`.
 *
 * Written with SheetJS so the formula is stored as a real `<f>` with a cached
 * value, which is what a workbook from Excel looks like — and what makes this a
 * test of reading a cross-sheet formula rather than of typing one.
 */
function writeCrossSheetWorkbook(): string {
  const wb = XLSX.utils.book_new()

  const summary = XLSX.utils.aoa_to_sheet([['', 'label']])
  // The cached value is deliberately WRONG (0): if Atlas displayed the cache
  // instead of evaluating, this test would read 0 and fail.
  summary.A1 = { t: 'n', f: 'Data!A1*2', v: 0 }
  XLSX.utils.book_append_sheet(wb, summary, 'Summary')

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[21]]), 'Data')

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-crosssheet-'))
  const file = path.join(dir, 'book.xlsx')
  fs.writeFileSync(file, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer)
  return file
}

async function launch(file: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({
    args: ['.', file],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  const page = await app.firstWindow()
  await page.waitForSelector('.spreadsheet-viewer__grid canvas', { timeout: 30_000 })
  await page.waitForTimeout(1500)
  return { app, page }
}

const HEADER_HEIGHT = 36
const ROW_HEIGHT = 32
const COLUMN_WIDTH = 120

async function clickCell(page: Page, col: number, row: number): Promise<void> {
  const box = (await page.locator('.spreadsheet-viewer__grid canvas').first().boundingBox())!
  await page.mouse.click(
    box.x + 48 + col * COLUMN_WIDTH + COLUMN_WIDTH / 2,
    box.y + HEADER_HEIGHT + row * ROW_HEIGHT + ROW_HEIGHT / 2,
  )
}

/**
 * Switches sheets.
 *
 * The sheet switcher is a row of plain buttons (`.spreadsheet-viewer__tab`),
 * not an ARIA tablist — `getByRole('tab')` matches the FILE tab strip instead,
 * which is what the first version of this test hit.
 */
async function selectSheet(page: Page, name: string): Promise<void> {
  await page.locator('.spreadsheet-viewer__tab', { hasText: name }).first().click()
  await page.waitForTimeout(800)
}

/**
 * The value the grid reports for the cell the user just clicked.
 *
 * Read from the `aria-live` status region (`useGridCellAnnouncement`), which
 * announces the selected cell's reference and VALUE. That is the only channel a
 * test can read a computed value from: the grid is a canvas, and the
 * accessibility `<table>` beside it publishes the column header strip rather
 * than the cell contents — which is what the first version of this test tried
 * to read, and why it saw nothing.
 */
async function selectedCellAnnouncement(page: Page): Promise<string> {
  return page.evaluate(() => {
    const region = document.querySelector('.spreadsheet-viewer__cell-announcement, [role="status"]')
    return region?.textContent?.trim() ?? ''
  })
}

test('a cross-sheet formula resolves on open and follows an edit to the other sheet', async () => {
  const file = writeCrossSheetWorkbook()
  const { app, page } = await launch(file)
  try {
    // Summary!A1 is `=Data!A1*2` with Data!A1 = 21, so 42 — and the stored
    // cache says 0, so reading 42 proves Atlas EVALUATED it rather than
    // displaying the cached value.
    await clickCell(page, 0, 0)
    await expect
      .poll(async () => selectedCellAnnouncement(page), {
        message: 'the cross-sheet formula should resolve when the workbook opens',
        timeout: 15_000,
      })
      .toContain('42')

    // Now change what it points at, on the OTHER sheet.
    await selectSheet(page, 'Data')
    await clickCell(page, 0, 0)
    await page.keyboard.type('5')
    await expect(page.locator('#portal textarea')).toBeFocused()
    await page.keyboard.press('Enter')
    await page.waitForTimeout(800)

    // Back to Summary: the formula must have followed the edit.
    await selectSheet(page, 'Summary')
    await clickCell(page, 0, 0)
    await expect
      .poll(async () => selectedCellAnnouncement(page), {
        message: 'editing the referenced sheet should update the referring one',
        timeout: 15_000,
      })
      .toContain('10')

    // And the formula itself survives a save, rather than being flattened to
    // the value it happened to show.
    await page.keyboard.press('Control+s')
    await page.waitForTimeout(3000)

    const zip = await JSZip.loadAsync(fs.readFileSync(file))
    // Only the worksheet PARTS: `zip.files` also lists directory entries and
    // the `_rels` folder, and `zip.file(dir)` returns null for those.
    const names = Object.keys(zip.files).filter(
      (n) => n.startsWith('xl/worksheets/') && n.endsWith('.xml') && !n.includes('_rels'),
    )
    let found = false
    for (const name of names) {
      const xml = await zip.file(name)!.async('string')
      if (/<f[^>]*>Data!A1\*2<\/f>/.test(xml)) found = true
    }
    expect(found, `the cross-sheet formula should still be in the saved file (checked ${names.join(', ')})`).toBe(
      true,
    )
  } finally {
    kill(app)
  }
})
