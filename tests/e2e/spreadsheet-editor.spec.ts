import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import * as XLSX from 'xlsx'

import { buildStyledWorkbook } from '../../src/viewers/spreadsheet/__tests__/styledWorkbook'

/**
 * USR-17 — spreadsheet editing in the real Electron app. Unit tests mock the
 * canvas grid, which is how 3.1.0 shipped with every cell editor failing to
 * open (glide-data-grid needs a `#portal` element in index.html).
 */

const projectRoot = process.cwd()
const HEADER_HEIGHT = 36
const ROW_HEIGHT = 32
const COLUMN_WIDTH = 120

const TABLE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Fruit" displayName="Fruit" ref="A1:C3" totalsRowShown="0"><autoFilter ref="A1:C3"/><tableColumns count="3"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Qty"/><tableColumn id="3" name="Price"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>`

async function createWorkbookWithTable(): Promise<string> {
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ['Name', 'Qty', 'Price'],
      ['Apple', 3, 1.5],
      ['Pear', 5, 2],
    ]),
    'Data',
  )
  const zip = await JSZip.loadAsync(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer)
  zip.file('xl/tables/table1.xml', TABLE_XML)
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
  )
  const sheetPath = 'xl/worksheets/sheet1.xml'
  const sheetXml = await zip.file(sheetPath)!.async('string')
  zip.file(sheetPath, sheetXml.replace('</worksheet>', '<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>'))
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-sheet-editor-'))
  const file = path.join(dir, 'table.xlsx')
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

async function typeIntoCell(page: Page, col: number, row: number, text: string): Promise<void> {
  await goToCell(page, col, row)
  // The first key opens the editor overlay (seeded with that key); the rest is
  // typed once it has focus and Enter follows immediately — the fast-typing
  // case that used to drop the edit.
  await page.keyboard.type(text.slice(0, 1))
  await expect(page.locator('#portal textarea')).toBeFocused()
  await page.keyboard.type(text.slice(1), { delay: 20 })
  await page.keyboard.press('Enter')
  await expect(page.locator('#portal textarea')).toHaveCount(0)
}

let fixture = ''
test.beforeAll(async () => {
  fixture = await createWorkbookWithTable()
})

test('USR-17: double-click opens the cell editor and commits the edit', async () => {
  const { app, page } = await launch(fixture)
  try {
    await goToCell(page, 0, 1)
    const box = (await page.locator('.spreadsheet-viewer__grid canvas').first().boundingBox())!
    await page.mouse.dblclick(box.x + 48 + COLUMN_WIDTH / 2, box.y + HEADER_HEIGHT + ROW_HEIGHT * 1.5)
    const editor = page.locator('#portal textarea')
    await expect(editor).toBeVisible()
    await expect(editor).toHaveValue('Apple')
    await editor.fill('Plum')
    await page.keyboard.press('Enter')
    await expect(editor).toHaveCount(0)
    await expect(page.getByRole('button', { name: 'Undo' })).toBeEnabled()
  } finally {
    kill(app)
  }
})

test('USR-17: typing edits cells inside a table and in the blank area, and Save keeps the Excel table', async () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-sheet-save-')), 'table.xlsx')
  fs.copyFileSync(fixture, file)
  const { app, page } = await launch(file)
  try {
    await typeIntoCell(page, 1, 2, '42')
    await expect(page.getByRole('button', { name: 'Undo' })).toBeEnabled()
    await typeIntoCell(page, 4, 5, 'outside')
    await expect(page.locator('.spreadsheet-viewer__stats')).toHaveText('6 rows × 5 columns')

    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect.poll(() => fs.statSync(file).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(fs.statSync(fixture).mtimeMs)
    await page.waitForTimeout(500)

    const bytes = fs.readFileSync(file)
    const sheet = XLSX.read(bytes).Sheets.Data
    expect(sheet.B3.v).toBe(42)
    expect(sheet.E6.v).toBe('outside')
    const zip = await JSZip.loadAsync(bytes)
    const tableXml = await zip.file('xl/tables/table1.xml')?.async('string')
    expect(tableXml).toContain('displayName="Fruit"')
    expect(tableXml).toContain('ref="A1:C3"')
    expect(await zip.file('xl/worksheets/sheet1.xml')!.async('string')).toContain('<tableParts count="1">')
  } finally {
    kill(app)
  }
})

test('USR-17: saving a styled workbook keeps its formatting and untouched cells', async () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-sheet-styled-')), 'budget.xlsx')
  fs.writeFileSync(file, Buffer.from(await buildStyledWorkbook()))
  const originalStyles = await (await JSZip.loadAsync(fs.readFileSync(file))).file('xl/styles.xml')!.async('string')

  const { app, page } = await launch(file)
  try {
    await typeIntoCell(page, 0, 1, 'Recycled paper')

    const before = fs.statSync(file).mtimeMs
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect.poll(() => fs.statSync(file).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(before)
    await page.waitForTimeout(400)

    const zip = await JSZip.loadAsync(fs.readFileSync(file))
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toContain('Recycled paper')
    // The date and currency cells the user never touched keep their style and value.
    expect(sheet).toContain('<c r="B2" s="3"><v>45000</v></c>')
    expect(sheet).toContain('<c r="C3" s="4"><f>C2*2</f><v>25</v></c>')
    expect(sheet).toContain('<conditionalFormatting sqref="C2:C3">')
    expect(await zip.file('xl/styles.xml')!.async('string')).toBe(originalStyles)
  } finally {
    kill(app)
  }
})
