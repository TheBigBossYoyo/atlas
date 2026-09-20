/**
 * USR-17 follow-up — the formula-reference-rewrite machinery
 * (`formulaRefs.ts`/`xlsxPassthrough.ts`) is thoroughly unit-tested against
 * each operation (rename, delete, row/column insert) in isolation, but never
 * driven end-to-end through the real editor UI with several structural
 * operations landing in the SAME save. This exercises exactly that: rename a
 * sheet a formula references, delete another sheet a different formula
 * references, and insert a row under the renamed sheet — all before one
 * save.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import * as XLSX from 'xlsx'

const projectRoot = process.cwd()
const HEADER_HEIGHT = 36
const ROW_HEIGHT = 32
const COLUMN_WIDTH = 120

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

async function launch(file: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({ args: ['.', file], cwd: projectRoot, env: { ...process.env, CI: '1', PLAYWRIGHT: '1' } })
  const page = await app.firstWindow()
  await page.waitForSelector('.spreadsheet-viewer__grid canvas', { timeout: 20_000 })
  await page.waitForTimeout(800)
  return { app, page }
}

async function goToCell(page: Page, col: number, row: number): Promise<void> {
  const box = (await page.locator('.spreadsheet-viewer__grid canvas').first().boundingBox())!
  await page.mouse.click(box.x + 48 + COLUMN_WIDTH / 2, box.y + HEADER_HEIGHT + ROW_HEIGHT / 2)
  for (let i = 0; i < col; i++) await page.keyboard.press('ArrowRight')
  for (let i = 0; i < row; i++) await page.keyboard.press('ArrowDown')
}

function buildWorkbook(): Buffer {
  const wb = XLSX.utils.book_new()

  const main = XLSX.utils.aoa_to_sheet([['Formula A', 'Formula B'], [0, 0]])
  main['A2'] = { t: 'n', f: 'Data!A1', v: 10 }
  main['B2'] = { t: 'n', f: 'Old!A1', v: 99 }
  XLSX.utils.book_append_sheet(wb, main, 'Main')

  const data = XLSX.utils.aoa_to_sheet([[10, 20]])
  XLSX.utils.book_append_sheet(wb, data, 'Data')

  const old = XLSX.utils.aoa_to_sheet([[99]])
  XLSX.utils.book_append_sheet(wb, old, 'Old')

  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer
}

test('rename a sheet referenced by a formula, delete another, insert a row under the renamed one, then save/reopen', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-formula-scenario-'))
  const file = path.join(dir, 'formulas.xlsx')
  fs.writeFileSync(file, buildWorkbook())

  const { app, page } = await launch(file)
  const pageErrors: string[] = []
  try {
    page.on('pageerror', (e) => pageErrors.push(String(e)))

    // Rename "Data" -> "Data2".
    await page.locator('.spreadsheet-viewer__tab', { hasText: 'Data' }).first().dblclick()
    const renameInput = page.locator('.spreadsheet-viewer__tab-rename')
    await renameInput.fill('Data2')
    await renameInput.press('Enter')
    await page.waitForTimeout(300)

    // Delete "Old".
    await page.getByRole('button', { name: 'Delete sheet Old' }).click()
    await page.waitForTimeout(300)

    // Switch to Data2, insert a row above row 1 (shifting its A1=10 down to A2).
    await page.locator('.spreadsheet-viewer__tab', { hasText: 'Data2' }).click()
    await page.waitForTimeout(300)
    await goToCell(page, 0, 0)
    await page.getByRole('button', { name: 'Insert row above' }).click()
    await page.waitForTimeout(300)

    const before = fs.statSync(file).mtimeMs
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect.poll(() => fs.statSync(file).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(before)
    await page.waitForTimeout(500)

    expect(pageErrors, pageErrors.join('\n')).toEqual([])

    const bytes = fs.readFileSync(file)
    const zip = await JSZip.loadAsync(bytes)
    const workbookXml = await zip.file('xl/workbook.xml')!.async('string')
    console.log('sheet names in workbook.xml:', workbookXml.match(/<sheet [^>]*name="[^"]*"/g))

    // Find which worksheet part is "Main" — sheet order/rIds can vary — then check its formulas.
    const wb = XLSX.read(bytes)
    console.log('sheet names (SheetJS):', wb.SheetNames)
    expect(wb.SheetNames).toEqual(['Main', 'Data2'])

    const mainSheet = wb.Sheets['Main']
    console.log('Main!A2 formula:', mainSheet['A2']?.f, 'value:', mainSheet['A2']?.v)
    console.log('Main!B2 formula:', mainSheet['B2']?.f, 'value:', mainSheet['B2']?.v)

    // The reference to the renamed sheet should follow the rename AND the row insert
    // (Data!A1, cloned byte-for-byte from the original file since the user never
    // edited this formula cell, gets both halves re-anchored — see formulaRefs.ts).
    expect(mainSheet['A2']?.f).toBe('Data2!A2')
    // The reference to the deleted sheet should become #REF!.
    expect(mainSheet['B2']?.f).toBe('#REF!')

    // Data2's own row insert actually happened: old A1 (10) is now at A2, row 1 is blank.
    const data2Sheet = wb.Sheets['Data2']
    console.log('Data2!A1:', data2Sheet['A1']?.v, 'Data2!A2:', data2Sheet['A2']?.v)
    expect(data2Sheet['A1']).toBeUndefined()
    expect(data2Sheet['A2']?.v).toBe(10)
  } finally {
    kill(app)
  }
})
