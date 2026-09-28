/**
 * Closing a large spreadsheet tab releases what it used.
 *
 * WHY THIS EXISTS. `docs/KNOWN_LIMITATIONS.md` carried a "closed spreadsheet tab
 * retains ~24 MB" entry, root-caused in prose to glide-data-grid's image loader
 * capturing the viewer's scope. Re-measured on 2026-09-28 it does not reproduce:
 * a 100 000-row / 19.5 MB `.xlsx` peaks at about +20 MB of renderer heap and
 * gives back 92–94% of it when its tab closes, leaving roughly 0.3–0.4 MB per
 * open/close cycle. The entry has been rewritten around this measurement, and
 * this test is what stops it drifting back to prose.
 *
 * TWO MEASUREMENT TRAPS, both hit while writing this:
 *   - `performance.memory.usedJSHeapSize` is quantized and cached for about 20
 *     minutes in Chromium. It reported a flat 9.5 MB across a 19.5 MB workbook
 *     being opened AND closed — which reads exactly like "no leak" and is really
 *     "no measurement". CDP's `Runtime.getHeapUsage` gives the true figure.
 *   - A tab appearing is not a workbook being rendered. Without waiting for the
 *     grid canvas, a peak that never happened makes the release ratio meaningless.
 *
 * The assertions are deliberately loose. The point is to catch a return to
 * "almost nothing is released", not to pin a number that depends on the machine,
 * V8's heap growth policy and how much of the app happens to be warm.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication } from '@playwright/test'
import * as XLSX from 'xlsx'

const projectRoot = process.cwd()

/** Two cycles: one cannot show accumulation, and each costs ~20s of wall clock. */
const CYCLES = 2
const MIN_RELEASED_FRACTION = 0.6
const MAX_RETAINED_BYTES = 8 * 1024 * 1024

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

function makeSheet(name: string, rows: number): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-mem-'))
  const file = path.join(dir, name)
  const aoa: (string | number)[][] = [['A', 'B', 'C', 'D', 'E']]
  for (let r = 0; r < rows; r += 1) {
    aoa.push([`row-${r}-text-value`, r, r * 1.5, `label ${r}`, r % 7])
  }
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'Data')
  fs.writeFileSync(file, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer)
  return file
}

test('closing a large spreadsheet tab gives back most of the heap it used', async () => {
  test.setTimeout(600_000)
  // The small sheet stays open throughout, so the baseline already includes the
  // spreadsheet viewer's own lazy chunk, the grid library and one live workbook —
  // what is measured is the big workbook alone, not first-open cost.
  const small = makeSheet('small.xlsx', 5)
  const big = makeSheet('big.xlsx', 100_000)

  const app = await electron.launch({
    args: ['.', small],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  const notes: string[] = [`big.xlsx on disk: ${(fs.statSync(big).size / 1024 / 1024).toFixed(1)} MB`]
  try {
    const page = await app.firstWindow()
    await page.waitForSelector('.spreadsheet-viewer__grid canvas', { timeout: 60_000 })
    await page.waitForTimeout(1500)

    const cdp = await page.context().newCDPSession(page)
    const heap = async (label: string): Promise<number> => {
      // Repeated, because one collection leaves plenty behind: V8 needs a few
      // passes before a detached React tree and its buffers are actually gone,
      // and a single pass reads as retention that is really collection lag.
      for (let pass = 0; pass < 5; pass += 1) {
        await cdp.send('HeapProfiler.collectGarbage')
        await page.waitForTimeout(250)
      }
      const { usedSize } = (await cdp.send('Runtime.getHeapUsage')) as { usedSize: number }
      notes.push(`${label}: ${(usedSize / 1024 / 1024).toFixed(1)} MB`)
      return usedSize
    }

    const baseline = await heap('baseline (small sheet only)')

    await app.evaluate(async ({ dialog }, filePath) => {
      dialog.showOpenDialog = (async () => ({ canceled: false, filePaths: [filePath] })) as typeof dialog.showOpenDialog
    }, big)

    for (let cycle = 1; cycle <= CYCLES; cycle += 1) {
      await page.keyboard.press('Control+o')
      await expect(page.getByRole('tab', { name: 'big.xlsx' })).toBeVisible({ timeout: 300_000 })
      await expect
        .poll(async () => page.locator('.spreadsheet-viewer__grid canvas').count(), { timeout: 300_000 })
        .toBeGreaterThan(0)
      await page.waitForTimeout(4000)
      const open = await heap(`cycle ${cycle}: open`)

      await page.keyboard.press('Control+w')
      await expect(page.getByRole('tab', { name: 'big.xlsx' })).toHaveCount(0, { timeout: 60_000 })
      await page.waitForTimeout(3000)
      const closed = await heap(`cycle ${cycle}: closed`)

      const peak = open - baseline
      const retained = closed - baseline
      notes.push(
        `cycle ${cycle}: peak +${(peak / 1024 / 1024).toFixed(1)} MB, ` +
          `retained +${(retained / 1024 / 1024).toFixed(1)} MB, ` +
          `released ${((1 - retained / (peak || 1)) * 100).toFixed(0)}%`,
      )

      // A peak this small means the workbook never really loaded, so the release
      // ratio below would be measuring nothing. Fail loudly rather than pass.
      expect(peak, notes.join(' | ')).toBeGreaterThan(8 * 1024 * 1024)
      expect(1 - retained / peak, notes.join(' | ')).toBeGreaterThan(MIN_RELEASED_FRACTION)
      expect(retained, notes.join(' | ')).toBeLessThan(MAX_RETAINED_BYTES)
    }
  } finally {
    kill(app)
  }
})
