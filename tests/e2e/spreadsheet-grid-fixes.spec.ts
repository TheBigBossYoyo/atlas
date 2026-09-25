import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import * as XLSX from 'xlsx'
import { expectDomFocused } from './helpers/domFocus'

/**
 * Night sweep findings SHEET-1 (Ctrl+End overshoots the used range), SHEET-2
 * (frozen-pane click selects the wrong row — confirmed NOT reproducible
 * against current code, see the "does not" test below, which pins the
 * correct behavior instead) and SHEET-4/SHEET-5 (percent/date auto-typing
 * on the xlsx passthrough save path). Driven through the real Electron app
 * (glide-data-grid needs a live canvas + `#portal`, same reasoning as
 * `spreadsheet-editor.spec.ts`), asserting on the bytes actually saved to
 * disk.
 */

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

async function gridBox(page: Page) {
  return (await page.locator('.spreadsheet-viewer__grid canvas').first().boundingBox())!
}

async function clickCell(page: Page, col: number, row: number): Promise<void> {
  const box = await gridBox(page)
  await page.mouse.click(box.x + 48 + COLUMN_WIDTH * col + COLUMN_WIDTH / 2, box.y + HEADER_HEIGHT + ROW_HEIGHT * row + ROW_HEIGHT / 2)
}

async function typeAt(page: Page, col: number, row: number, text: string): Promise<void> {
  await clickCell(page, col, row)
  await page.keyboard.type(text.slice(0, 1))
  await expectDomFocused(page, '#portal textarea')
  if (text.length > 1) await page.keyboard.type(text.slice(1), { delay: 20 })
  await page.keyboard.press('Enter')
  await expect(page.locator('#portal textarea')).toHaveCount(0)
}

async function saveAndWait(page: Page, file: string): Promise<void> {
  const before = fs.statSync(file).mtimeMs
  await page.getByRole('button', { name: 'Save', exact: true }).click()
  await expect.poll(() => fs.statSync(file).mtimeMs, { timeout: 10_000 }).toBeGreaterThan(before)
  await page.waitForTimeout(400)
}

function tmpFile(name: string): string {
  return path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-sheet-grid-')), name)
}

// ---------------------------------------------------------------------------
// SHEET-1 — Ctrl+End must land on the real last used cell, not the blank
// editing margin `useGridBlankMargin` draws past the data. A 150x5 sheet is
// enough to trigger the exact same overscan `useGridBlankMargin` always
// applies (50 blank rows / up to 26 columns, uncapped by data size) that a
// manually-verified 20,001-row repro hit as `Z20051` — see the fix's own
// investigation notes; kept small here so the spec stays fast under the
// "one Playwright spec at a time" CPU rule.
// ---------------------------------------------------------------------------

function buildLargeFixture(): string {
  const rows = Array.from({ length: 150 }, (_, r) => [`r${r}c1`, `r${r}c2`, `r${r}c3`, `r${r}c4`, `r${r}c5`])
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Data')
  const file = tmpFile('large.xlsx')
  fs.writeFileSync(file, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer)
  return file
}

test('SHEET-1: Ctrl+End lands on the real last used cell, not the blank margin', async () => {
  const file = buildLargeFixture()
  const { app, page } = await launch(file)
  try {
    await clickCell(page, 0, 0)
    await page.keyboard.press('Control+End')
    await page.waitForTimeout(300)

    // Typing here must land on/overwrite the real last cell (E150), not
    // create a phantom cell far past the data (the bug landed on Z200 for
    // this fixture's shape, mirroring the manually-verified Z20051 on the
    // 20,001-row repro).
    await page.keyboard.type('9')
    await expectDomFocused(page, '#portal textarea')
    await page.keyboard.type('99999', { delay: 20 })
    await page.keyboard.press('Enter')
    await expect(page.locator('#portal textarea')).toHaveCount(0)

    await saveAndWait(page, file)

    const wb = XLSX.read(fs.readFileSync(file))
    const ws = wb.Sheets.Data
    expect(ws['!ref']).toBe('A1:E150') // dimensions must NOT have grown past the real data
    expect(ws.E150.v).toBe(999999) // the overwritten real last cell
    expect(ws.Z200).toBeUndefined() // no phantom cell out in the old blank-margin overscan
  } finally {
    kill(app)
  }
})

// ---------------------------------------------------------------------------
// SHEET-2 — investigated as a reported bug (a frozen-pane click landing one
// row off); driving the real app with a genuine mouse click (not the
// click-then-arrow-key sequence the original report's harness used, which
// double-counted the frozen row already excluded from the scrollable grid's
// own row numbering) shows the click -> sheet-row mapping is already
// correct. This test locks that in as a regression guard: a plain click on
// the first scrollable row below a frozen header must land exactly there.
// ---------------------------------------------------------------------------

function buildFrozenPaneWorkbookBuffer(): Buffer {
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ['H1', 'H2'],
      ['R2C1', 'R2C2'],
      ['R3C1', 'R3C2'],
    ]),
    'Data',
  )
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer
}

async function writeFrozenPaneFixture(): Promise<string> {
  const zip = await JSZip.loadAsync(buildFrozenPaneWorkbookBuffer())
  let sheet1 = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  sheet1 = sheet1.replace(
    /<sheetView([^>]*)\/>/,
    '<sheetView$1><pane xSplit="0" ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView>',
  )
  zip.file('xl/worksheets/sheet1.xml', sheet1)
  const file = tmpFile('frozen.xlsx')
  fs.writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer' }))
  return file
}

test('SHEET-2: a plain click on the first scrollable row under a frozen header hits that row, not one below', async () => {
  const file = await writeFrozenPaneFixture()
  const { app, page } = await launch(file)
  try {
    // gridRow 0 is the first row BELOW the frozen strip — already sheet row
    // index 1 (A2), since the frozen row is excluded from the scrollable
    // grid's own numbering (`useGridBlankMargin`'s `rowOffset`). A single
    // click there (no arrow-key navigation) must select A2, not A3.
    await typeAt(page, 0, 0, 'MARKER')
    await saveAndWait(page, file)

    const wb = XLSX.read(fs.readFileSync(file))
    const ws = wb.Sheets.Data
    expect(ws.A2?.v).toBe('MARKER')
    expect(ws.A3?.v).toBe('R3C1') // untouched — the bug would have overwritten this instead
  } finally {
    kill(app)
  }
})

test('SHEET-2: "Insert row above" on a frozen-pane sheet inserts at the clicked row', async () => {
  const file = await writeFrozenPaneFixture()
  const { app, page } = await launch(file)
  try {
    await clickCell(page, 0, 1) // gridRow 1 -> sheet row index 2 (A3), the second scrollable row
    await page.getByRole('button', { name: 'Insert row above' }).click()
    await saveAndWait(page, file)

    const wb = XLSX.read(fs.readFileSync(file))
    const ws = wb.Sheets.Data
    expect(ws.A2?.v).toBe('R2C1') // row 2 (first scrollable row) is untouched
    expect(ws.A3).toBeUndefined() // the blank row landed exactly where clicked
    expect(ws.A4?.v).toBe('R3C1') // the old row 3 shifted down by one
  } finally {
    kill(app)
  }
})

// ---------------------------------------------------------------------------
// SHEET-4/SHEET-5 — typing `50%`/`12.5%` and an ISO date save as a real
// percent/date value with a matching number format, not literal text (the
// xlsx passthrough save path only — `xlsxPassthrough.test.ts` covers the
// unit-level styles.xml details this only spot-checks end to end).
// ---------------------------------------------------------------------------

test('SHEET-4/5: typing 50% and an ISO date save as a real percent/date value, not literal text', async () => {
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['H1', 'H2']]), 'Data')
  const file = tmpFile('percent-date.xlsx')
  fs.writeFileSync(file, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer)

  const { app, page } = await launch(file)
  try {
    await typeAt(page, 0, 1, '50%')
    await typeAt(page, 1, 1, '2024-03-14')
    await saveAndWait(page, file)

    const bytes = fs.readFileSync(file)
    const wbOut = XLSX.read(bytes, { cellNF: true })
    const ws = wbOut.Sheets.Data
    expect(ws.A2).toMatchObject({ t: 'n', v: 0.5 })
    expect(ws.A2!.z).toContain('%')
    expect(ws.B2).toMatchObject({ t: 'n', v: 45365 }) // 2024-03-14's real Excel serial
    expect(ws.B2!.z).toBe('yyyy-mm-dd')

    const zip = await JSZip.loadAsync(bytes)
    const stylesXml = await zip.file('xl/styles.xml')!.async('string')
    expect(stylesXml).toContain('formatCode="yyyy-mm-dd"')

    const out = execSync(`node scripts/validate-office-file.mjs "${file}"`, { cwd: projectRoot, encoding: 'utf8' })
    expect(out).toContain('OK')
  } finally {
    kill(app)
  }
})
