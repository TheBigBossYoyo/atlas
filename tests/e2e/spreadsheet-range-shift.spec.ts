import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'

import { buildStyledWorkbook } from '../../src/viewers/spreadsheet/__tests__/styledWorkbook'

/**
 * SHEET-3 — a whole-column (`A:A`) conditional format must survive a row
 * insert + save through the REAL Electron app, not just Atlas's own unit
 * tests (`spreadsheetRangeShift.test.ts`, `xlsxPassthrough.test.ts`).
 *
 * Before the fix, `decodeRange` (which only understands `A1`-style cell
 * corners) could not parse a whole-column/whole-row `sqref` piece like
 * `A:A` or `1:1`; `remapSqref` then treated it as malformed input and
 * dropped it, and `updateShiftedRanges` in `xlsxPassthrough.ts` removed the
 * owning `<conditionalFormatting>` element outright — on ANY row or column
 * insert/delete anywhere on the sheet, not just one that actually touched
 * column A. This test proves the rule survives a row insert + save, and
 * that the saved file still opens cleanly, by unzipping the actual bytes
 * written to disk — not by trusting what the UI shows.
 */

const projectRoot = process.cwd()
const HEADER_HEIGHT = 36
const ROW_HEIGHT = 32
const COLUMN_WIDTH = 120

async function createWorkbookWithWholeColumnCF(): Promise<string> {
  const zip = await JSZip.loadAsync(await buildStyledWorkbook())
  const sheetPath = 'xl/worksheets/sheet1.xml'
  const sheetXml = await zip.file(sheetPath)!.async('string')
  // `buildStyledWorkbook`'s own conditional format is scoped to `C2:C3`;
  // replace it with a whole-column one so this fixture reproduces SHEET-3.
  const patched = sheetXml.replace('sqref="C2:C3"', 'sqref="A:A"')
  if (patched === sheetXml) throw new Error('fixture sheet XML did not contain the expected sqref="C2:C3" to patch')
  zip.file(sheetPath, patched)
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-sheet-rangeshift-'))
  const file = path.join(dir, 'wholecol.xlsx')
  fs.writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer' }))
  return file
}

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

/** Clicks A1 (row-marker width varies, so aim at the middle of column A) and moves with the arrow keys. */
async function goToCell(page: Page, col: number, row: number): Promise<void> {
  const box = (await page.locator('.spreadsheet-viewer__grid canvas').first().boundingBox())!
  await page.mouse.click(box.x + 48 + COLUMN_WIDTH / 2, box.y + HEADER_HEIGHT + ROW_HEIGHT / 2)
  for (let i = 0; i < col; i++) await page.keyboard.press('ArrowRight')
  for (let i = 0; i < row; i++) await page.keyboard.press('ArrowDown')
}

let fixture = ''
test.beforeAll(async () => {
  fixture = await createWorkbookWithWholeColumnCF()
})

test('SHEET-3: a whole-column conditional format survives a row insert + save, and the saved file reopens cleanly', async () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-sheet-rangeshift-save-')), 'wholecol.xlsx')
  fs.copyFileSync(fixture, file)

  const before = await JSZip.loadAsync(fs.readFileSync(file))
  expect(await before.file('xl/worksheets/sheet1.xml')!.async('string')).toContain('<conditionalFormatting sqref="A:A">')

  const { app, page } = await launch(file)
  try {
    // Select A2 and insert a row above it — a structural edit that has
    // nothing to do with column A's own position, which is exactly the
    // SHEET-3 reproduction: ANY row/column edit used to wipe the rule.
    await goToCell(page, 0, 1)
    await page.getByRole('button', { name: 'Insert row above' }).click()
    await expect(page.locator('.spreadsheet-viewer__stats')).toHaveText('4 rows × 3 columns')

    const beforeMtime = fs.statSync(file).mtimeMs
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect.poll(() => fs.statSync(file).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(beforeMtime)
    await page.waitForTimeout(500)
  } finally {
    kill(app)
  }

  // Verify against the ACTUAL bytes written to disk, not the UI.
  const savedBytes = fs.readFileSync(file)
  const zip = await JSZip.loadAsync(savedBytes)
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  // A whole-column range has no row bound to shift, so it survives the row
  // insert unchanged — the bug was that it was DELETED, not mis-anchored.
  expect(sheetXml).toContain('<conditionalFormatting sqref="A:A">')
  expect(sheetXml).toContain('<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>20</formula></cfRule>')

  // The spec-level OPC/OOXML validator must accept the file we just saved.
  const validation = execSync(`node scripts/validate-office-file.mjs "${file}"`, { cwd: projectRoot, encoding: 'utf8' })
  expect(validation).not.toContain('error')

  // "Reopen": relaunch the real app against the saved file and confirm it
  // opens cleanly (no crash, no fallback-to-raw-text) with the row insert
  // still applied.
  const reopened = await launch(file)
  try {
    await expect(reopened.page.locator('.spreadsheet-viewer__stats')).toHaveText('4 rows × 3 columns')
    await expect(reopened.page.locator('.spreadsheet-viewer__grid canvas').first()).toBeVisible()
  } finally {
    kill(reopened.app)
  }
})
