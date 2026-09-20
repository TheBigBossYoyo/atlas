/**
 * MEM-01 — closing a document tab must not keep its bytes alive forever.
 *
 * A performance pass found that after opening a large .xlsx, closing its tab
 * and forcing garbage collection through the DevTools protocol, the renderer
 * released only a small fraction of what it had used at peak. A heap
 * snapshot traced the biggest offender to `documentSessions.ts`'s
 * `recentlyClosed` stack (the one Ctrl+Shift+T reopens from): it kept the
 * FULL bytes of up to five closed documents alive, even though reopening one
 * always re-reads it from disk first (`reloadSessionFile`/`showSessionFile`)
 * and never actually uses those kept bytes on the happy path. See
 * `documentSessions.ts`'s `shedBytes` for the fix.
 *
 * This test opens more than `MAX_RECENTLY_CLOSED` (5) distinct multi-MB
 * binary documents one at a time, closing each before opening the next (so
 * every one of them passes through `recentlyClosed`), forces GC, and asserts
 * the renderer's process memory hasn't grown by anywhere near "several
 * closed documents' worth of bytes". The bound is deliberately generous
 * (tens of MB of slack) so ordinary JS engine/allocator noise across a CI
 * run never makes this flaky — it only needs to catch the shape of the old
 * bug (unbounded-looking growth that scales with how many documents were
 * opened and closed), not pin down an exact number.
 */
import { promises as fs } from 'node:fs'
import path from 'node:path'
import { execFileSync } from 'node:child_process'

import { _electron as electron, expect, test } from '@playwright/test'
import * as XLSX from 'xlsx'

const projectRoot = process.cwd()
const fixtureDir = path.join(projectRoot, 'tests', 'e2e', 'fixtures')

// ~30k distinct rows compresses to several MB per .xlsx (matching the order
// of magnitude of the original bug report) while staying fast enough to
// parse in a regression test — see perf.spec.ts's 100k-row fixture for the
// same "distinct data defeats zip compression" reasoning at 3x this size.
const ROWS_PER_FILE = 30_000
const FILE_COUNT = 8 // > MAX_RECENTLY_CLOSED (5), so eviction is also exercised.
// Retaining even half of one closed document's worth of bytes per slot
// across all 5 capped `recentlyClosed` entries would clear this bar easily;
// the fix should land far under it.
const GROWTH_BUDGET_MB = 40

function kill(app: import('@playwright/test').ElectronApplication): void {
  try {
    execFileSync('taskkill', ['/PID', String(app.process().pid), '/T', '/F'], { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

async function writeXlsxFixture(filePath: string, seed: number): Promise<void> {
  const rows: (string | number)[][] = [['id', 'name', 'note']]
  for (let i = 0; i < ROWS_PER_FILE; i++) {
    // `seed` + row index keeps every file's bytes distinct (so nothing
    // downstream can accidentally dedupe/share them) without needing truly
    // random data.
    rows.push([i, `Row ${seed}-${i}`, `padding-${seed}-${i}-${(i * seed) % 97}`])
  }
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), 'Sheet1')
  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }) as Buffer
  await fs.writeFile(filePath, buffer)
}

test('closing documents releases their bytes instead of piling up in the closed-tabs stack', async () => {
  test.setTimeout(180_000)

  const fixturePaths = Array.from({ length: FILE_COUNT }, (_, i) => path.join(fixtureDir, `mem-retention-${i}.xlsx`))
  await Promise.all(fixturePaths.map((p, i) => writeXlsxFixture(p, i + 1)))

  const app = await electron.launch({
    args: ['.', fixturePaths[0]],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })

  try {
    const page = await app.firstWindow()
    await page.waitForSelector('#root', { timeout: 30_000 })
    const client = await page.context().newCDPSession(page)
    await client.send('HeapProfiler.enable')

    async function forceGC(): Promise<void> {
      for (let i = 0; i < 4; i++) {
        await client.send('HeapProfiler.collectGarbage')
        await page.waitForTimeout(100)
      }
    }

    async function privateMB(): Promise<number> {
      const metrics = await app.evaluate(({ app: electronApp }) => electronApp.getAppMetrics())
      // Electron 44 labels the web-contents process "Tab" (its ProcessMetric
      // type doesn't even list the older "Renderer" name any more).
      const renderers = metrics.filter((m) => m.type === 'Tab')
      const totalKb = renderers.reduce((sum, m) => sum + (m.memory.privateBytes ?? m.memory.workingSetSize ?? 0), 0)
      return totalKb / 1024
    }

    async function waitForOpen(fileName: string): Promise<void> {
      await expect(page.locator('.statusbar__file')).toHaveText(fileName, { timeout: 30_000 })
      await page.waitForSelector('[data-viewer="xlsx"] canvas', { timeout: 30_000 })
    }

    async function closeActiveTab(fileName: string): Promise<void> {
      await page.getByRole('button', { name: `Close ${fileName}` }).click({ timeout: 15_000 })
    }

    // Open the first file (already loaded via the CLI arg) and close it —
    // this is the baseline measurement point: by here the app, its
    // dependency chunks and one document's worth of one-time setup cost
    // have all already been paid for, so later growth is attributable to
    // the open/close cycles themselves, not app startup.
    await waitForOpen(path.basename(fixturePaths[0]))
    await closeActiveTab(path.basename(fixturePaths[0]))
    await forceGC()
    await page.waitForTimeout(300)
    const baselineMB = await privateMB()

    // Open and close every remaining fixture, one at a time, through the
    // real "Open" dialog flow (not another CLI launch) — each one lands in
    // `recentlyClosed` on close, exactly like a user clicking through
    // several documents and closing each when done.
    for (let i = 1; i < FILE_COUNT; i++) {
      const filePath = fixturePaths[i]
      const fileName = path.basename(filePath)
      await app.evaluate(async ({ dialog }, targetPath) => {
        dialog.showOpenDialog = (async () => ({ canceled: false, filePaths: [targetPath] })) as typeof dialog.showOpenDialog
      }, filePath)
      await page.getByRole('button', { name: 'Open', exact: true }).click()
      await waitForOpen(fileName)
      await closeActiveTab(fileName)
    }

    await forceGC()
    await page.waitForTimeout(300)
    const afterMB = await privateMB()

    const growthMB = afterMB - baselineMB
    expect(
      growthMB,
      `renderer private memory grew ${growthMB.toFixed(1)}MB (baseline ${baselineMB.toFixed(1)}MB -> ${afterMB.toFixed(1)}MB) after opening and closing ${FILE_COUNT - 1} more ${ROWS_PER_FILE}-row .xlsx documents — the closed-tabs stack (recentlyClosed) should not be keeping their bytes alive`,
    ).toBeLessThan(GROWTH_BUDGET_MB)
  } finally {
    kill(app)
    await Promise.all(fixturePaths.map((p) => fs.rm(p, { force: true })))
  }
})
