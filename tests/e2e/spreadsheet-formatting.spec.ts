/**
 * SHEETFMT-1 — a workbook's own cell formatting, in the real app.
 *
 * The unit tests prove the styles are read and reach the cell the grid asks
 * for. This proves the last step nothing else can: that the canvas actually
 * draws them. The grid is a `<canvas>`, so there is no DOM node to assert on —
 * the test reads PIXELS back out of it.
 *
 * That makes it the only test in the suite that can tell "the format was
 * computed" apart from "the format was drawn", which is exactly the gap the
 * accessibility and editing specs leave open for this canvas grid.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'

const projectRoot = process.cwd()
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const rels = (items: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

/**
 * A workbook with one unmistakable piece of formatting: A1:C1 filled solid
 * magenta. A colour no theme of Atlas's uses, so finding it in the canvas can
 * only mean the file's own fill was drawn.
 */
async function writeFormattedWorkbook(): Promise<string> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
  )
  zip.file('_rels/.rels', rels(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`))
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${NS}" xmlns:r="${REL}"><sheets><sheet name="Report" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    rels(
      `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/>`,
    ),
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${NS}">` +
      `<dimension ref="A1:C2"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData>` +
      `<row r="1">` +
      `<c r="A1" s="1" t="inlineStr"><is><t>Item</t></is></c>` +
      `<c r="B1" s="1" t="inlineStr"><is><t>Qty</t></is></c>` +
      `<c r="C1" s="1" t="inlineStr"><is><t>Note</t></is></c>` +
      `</row>` +
      `<row r="2">` +
      `<c r="A2" t="inlineStr"><is><t>Paper</t></is></c>` +
      `<c r="B2"><v>3</v></c>` +
      `<c r="C2" t="inlineStr"><is><t>plain</t></is></c>` +
      `</row></sheetData></worksheet>`,
  )
  zip.file(
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="${NS}">` +
      `<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>` +
      `<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>` +
      `<fills count="3"><fill><patternFill patternType="none"/></fill>` +
      `<fill><patternFill patternType="gray125"/></fill>` +
      `<fill><patternFill patternType="solid"><fgColor rgb="FFFF00FF"/><bgColor indexed="64"/></patternFill></fill></fills>` +
      `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="2">` +
      `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
      `<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>` +
      `</cellXfs></styleSheet>`,
  )

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-sheetfmt-'))
  const file = path.join(dir, 'report.xlsx')
  fs.writeFileSync(file, Buffer.from(await zip.generateAsync({ type: 'arraybuffer' })))
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
  // The grid's first paint comes from the value parse; the formatting arrives
  // with the zip read behind it, so give the repaint a moment.
  await page.waitForTimeout(2500)
  return { app, page }
}

/**
 * Counts how many pixels in the grid canvas match `hex`, within `tolerance` per
 * channel.
 *
 * Reads the canvas the grid actually drew into, via its 2D context. The
 * tolerance exists because the grid composites and may antialias text over the
 * fill; an exact match would be brittle for no gain, while a wide tolerance
 * would stop distinguishing magenta from anything else.
 */
async function countPixels(page: Page, hex: string, tolerance = 12): Promise<number> {
  return page.evaluate(
    ({ hex, tolerance }) => {
      const target = {
        r: parseInt(hex.slice(1, 3), 16),
        g: parseInt(hex.slice(3, 5), 16),
        b: parseInt(hex.slice(5, 7), 16),
      }
      // The grid renders into one or more canvases; the biggest is the cell area.
      const canvases = Array.from(document.querySelectorAll('.spreadsheet-viewer__grid canvas'))
      let best: HTMLCanvasElement | null = null
      for (const c of canvases) {
        const canvas = c as HTMLCanvasElement
        if (!best || canvas.width * canvas.height > best.width * best.height) best = canvas
      }
      if (!best) return -1
      const ctx = best.getContext('2d', { willReadFrequently: true })
      if (!ctx) return -2
      const { data } = ctx.getImageData(0, 0, best.width, best.height)
      let count = 0
      for (let i = 0; i < data.length; i += 4) {
        if (
          Math.abs(data[i] - target.r) <= tolerance &&
          Math.abs(data[i + 1] - target.g) <= tolerance &&
          Math.abs(data[i + 2] - target.b) <= tolerance
        ) {
          count += 1
        }
      }
      return count
    },
    { hex, tolerance },
  )
}

test("a workbook's own cell fill is actually drawn in the grid", async () => {
  const file = await writeFormattedWorkbook()
  const { app, page } = await launch(file)
  try {
    const magenta = await countPixels(page, '#FF00FF')
    expect(magenta, 'the canvas could not be read').toBeGreaterThanOrEqual(0)

    // Three header cells at the grid's default 120x32 would be ~11k pixels.
    // A few thousand is a filled header row; a handful would be noise.
    expect(
      magenta,
      `expected the magenta header fill to be drawn, found ${magenta} matching pixels`,
    ).toBeGreaterThan(2000)

    // The control: a colour the file never uses must not appear. Without this,
    // a bug that flooded the canvas with one colour would satisfy the check
    // above.
    const unused = await countPixels(page, '#00FF00')
    expect(unused, `a colour the workbook never uses was drawn (${unused} pixels)`).toBeLessThan(100)
  } finally {
    kill(app)
  }
})
