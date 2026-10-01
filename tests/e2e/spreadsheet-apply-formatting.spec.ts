/**
 * SHEETFMT-2 — applying cell formatting, in the real app, all the way to disk.
 *
 * The whole loop, through the same code path a person uses: open a workbook,
 * select a cell, press the toolbar's Bold and pick a fill, save, and then read
 * the SAVED FILE's own XML to check Excel would see the same thing.
 *
 * That last step is the point. The unit tests prove each half separately, and a
 * test that only reopened the file in Atlas would pass on a file no other
 * spreadsheet program understands — which, for a format whose reason to exist
 * is interoperability, is the failure that matters.
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

/** A plain, unstyled workbook — so any formatting found afterwards can only be what the test applied. */
function writePlainWorkbook(): string {
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ['Item', 'Qty'],
      ['Paper', 3],
      ['Ink', 7],
    ]),
    'Sheet1',
  )
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-applyfmt-'))
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

// Grid geometry, as `spreadsheet-editor.spec.ts` uses it — the canvas has no
// DOM cells to address.
const HEADER_HEIGHT = 36
const ROW_HEIGHT = 32
const COLUMN_WIDTH = 120

/** Clicks cell (col,row) in the canvas grid. */
async function clickCell(page: Page, col: number, row: number): Promise<void> {
  const box = (await page.locator('.spreadsheet-viewer__grid canvas').first().boundingBox())!
  await page.mouse.click(
    box.x + 48 + col * COLUMN_WIDTH + COLUMN_WIDTH / 2,
    box.y + HEADER_HEIGHT + row * ROW_HEIGHT + ROW_HEIGHT / 2,
  )
}

/** The `<cellXfs>` entries of a saved styles part, in order — index n is a cell's `s="n"`. */
function cellXfs(stylesXml: string): string[] {
  const block = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml)
  if (!block) return []
  return block[1].match(/<xf\b[\s\S]*?(?:\/>|<\/xf>)/g) ?? []
}

function fontsOf(stylesXml: string): string[] {
  const block = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(stylesXml)
  if (!block) return []
  return block[1].match(/<font>[\s\S]*?<\/font>|<font\s*\/>/g) ?? []
}

test('bold and a fill applied in the app land in the saved file as real OOXML styles', async () => {
  const file = writePlainWorkbook()
  const before = fs.readFileSync(file)
  const { app, page } = await launch(file)
  try {
    // Select "Paper" (A2) and format it.
    await clickCell(page, 0, 1)
    await page.getByRole('button', { name: 'Bold', exact: true }).click()
    await page.waitForTimeout(300)

    await page.getByRole('button', { name: /Fill color/i }).click()
    await page.getByRole('button', { name: '#FF0000' }).click()
    await page.waitForTimeout(300)

    // Save through the real Ctrl+S path.
    await page.keyboard.press('Control+s')
    await page.waitForTimeout(3000)

    const after = fs.readFileSync(file)
    expect(after.equals(before), 'the file should have been written').toBe(false)

    const zip = await JSZip.loadAsync(after)
    const stylesFile = zip.file('xl/styles.xml')
    expect(stylesFile, 'the saved workbook should have a styles part').not.toBeNull()
    const styles = await stylesFile!.async('string')
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')

    // A2 must now point at an xf whose font is bold and whose fill is the red
    // that was picked.
    const cellTag = /<c[^>]*\br="A2"[^>]*>/.exec(sheet)
    expect(cellTag, 'A2 should be present in the saved sheet').not.toBeNull()
    const styleIndex = /\bs="(\d+)"/.exec(cellTag![0])
    expect(styleIndex, `A2 should carry a style index: ${cellTag![0]}`).not.toBeNull()

    const xf = cellXfs(styles)[Number(styleIndex![1])]
    expect(xf, 'the style index should resolve to a cellXfs entry').toBeTruthy()

    const fontId = Number(/\bfontId="(\d+)"/.exec(xf)![1])
    expect(fontsOf(styles)[fontId], `font ${fontId} should be bold`).toContain('<b')

    const fillId = Number(/\bfillId="(\d+)"/.exec(xf)![1])
    const fills = (/<fills\b[^>]*>([\s\S]*?)<\/fills>/.exec(styles)?.[1] ?? '').match(
      /<fill>[\s\S]*?<\/fill>|<fill\s*\/>/g,
    )!
    expect(fills[fillId], `fill ${fillId} should be solid red`).toContain('FFFF0000')

    // And the cell's VALUE survived being formatted.
    const cellBlock = /<c[^>]*\br="A2"[^>]*>[\s\S]*?<\/c>/.exec(sheet)?.[0] ?? cellTag![0]
    const sharedStrings = zip.file('xl/sharedStrings.xml')
      ? await zip.file('xl/sharedStrings.xml')!.async('string')
      : ''
    expect(
      cellBlock.includes('Paper') || sharedStrings.includes('Paper'),
      `A2 should still hold "Paper": ${cellBlock}`,
    ).toBe(true)
  } finally {
    kill(app)
  }
})

test('a formatted workbook reopens showing the formatting it was saved with', async () => {
  // The round trip a user actually notices: format, save, close, reopen.
  const file = writePlainWorkbook()
  const first = await launch(file)
  try {
    await clickCell(first.page, 1, 1)
    await first.page.getByRole('button', { name: 'Italic', exact: true }).click()
    await first.page.waitForTimeout(300)
    await first.page.keyboard.press('Control+s')
    await first.page.waitForTimeout(3000)
  } finally {
    kill(first.app)
  }

  const second = await launch(file)
  try {
    // Select the same cell and read the toolbar's own toggle: it reflects the
    // selected cell's format, so a pressed Italic means the file came back
    // formatted.
    await clickCell(second.page, 1, 1)
    await second.page.waitForTimeout(500)
    await expect(second.page.getByRole('button', { name: 'Italic', exact: true })).toHaveAttribute(
      'aria-pressed',
      'true',
    )
    // A neighbour that was never formatted must not be.
    await clickCell(second.page, 0, 1)
    await second.page.waitForTimeout(500)
    await expect(second.page.getByRole('button', { name: 'Italic', exact: true })).toHaveAttribute(
      'aria-pressed',
      'false',
    )
  } finally {
    kill(second.app)
  }
})
