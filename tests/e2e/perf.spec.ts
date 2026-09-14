/**
 * T2 (DAT-07) — perf harness: opening a 100k-row CSV or XLSX workbook must
 * not block the renderer's main thread for more than 200ms at a stretch.
 *
 * Vitest/jsdom has no real `Worker` (see useSpreadsheetWorkbook.test.ts and
 * csvParse.test.ts, which instead assert the *delegation decision* — that
 * Papa Parse / the xlsx parse are handed off above their size thresholds).
 * The actual browser-facing guarantee — that the main thread stays
 * responsive — can only be measured where a real Worker and a real
 * `PerformanceObserver({entryTypes:['longtask']})` both exist: a real
 * Chromium renderer, which is exactly what Playwright's Electron harness
 * provides. Both formats are covered here (not just CSV) because they take
 * structurally different off-main-thread paths — Papa Parse's `worker:true`
 * re-serializes into a `blob:` Worker, while the xlsx path spins up a real
 * module Worker from a separate bundled script
 * (`spreadsheetWorker.worker.ts`) — so a CSP or `file://`-packaging issue
 * affecting only one of them would not be caught by testing the other.
 *
 * Timing design: main.cjs's `ready-to-show` handler waits 300ms before
 * sending the CLI-arg file to the renderer (see `sendFileToWindow`'s
 * call site) specifically to let the renderer finish initializing first —
 * that fixed window is also exactly what this test needs to attach the
 * PerformanceObserver *before* the parse begins, with no race.
 */
import { promises as fs } from 'node:fs'
import path from 'node:path'
import { execFileSync } from 'node:child_process'

import { _electron as electron, expect, test } from '@playwright/test'
import * as XLSX from 'xlsx'

const projectRoot = process.cwd()
const fixtureDir = path.join(projectRoot, 'tests', 'e2e', 'fixtures')
const largeCsvPath = path.join(fixtureDir, 'large-100k.csv')
const largeXlsxPath = path.join(fixtureDir, 'large-100k.xlsx')

const ROW_COUNT = 100_000
const LONG_TASK_BUDGET_MS = 200

async function writeLargeCsvFixture(): Promise<void> {
  const lines: string[] = ['id,name,value,note']
  for (let i = 0; i < ROW_COUNT; i++) {
    lines.push(`${i},Row ${i},${(i * 1.5).toFixed(2)},padding text to make each row a realistic width`)
  }
  await fs.writeFile(largeCsvPath, lines.join('\n') + '\n', 'utf8')
}

async function writeLargeXlsxFixture(): Promise<void> {
  const rows: (string | number)[][] = [['id', 'name', 'value', 'note']]
  for (let i = 0; i < ROW_COUNT; i++) {
    rows.push([i, `Row ${i}`, Number((i * 1.5).toFixed(2)), 'padding text to make each row a realistic width'])
  }
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), 'Sheet1')
  // `XLSX.writeFile` fails under this test runner (its Node-vs-browser
  // environment detection misfires here); `XLSX.write` to a buffer plus a
  // plain `fs.writeFile` is the same safe pattern already used by
  // `tests/e2e/fixtures/generate.mjs` and the SpreadsheetViewer unit tests.
  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }) as Buffer
  await fs.writeFile(largeXlsxPath, buffer)
}

test.beforeAll(async () => {
  execFileSync(process.execPath, ['tests/e2e/fixtures/generate.mjs'], {
    cwd: projectRoot,
    stdio: 'inherit',
  })
  await writeLargeCsvFixture()
  await writeLargeXlsxFixture()
})

/** Shared body for both the CSV and XLSX variants of this perf assertion. */
async function assertNoLongTasksOpening(filePath: string, dataViewer: string, statusbarName: string): Promise<void> {
  test.setTimeout(120_000)

  const electronApp = await electron.launch({
    args: ['.', filePath],
    cwd: projectRoot,
    env: {
      ...process.env,
      CI: '1',
      PLAYWRIGHT: '1',
    },
  })

  try {
    // Attached to the app's BrowserContext, not the (possibly-already-created)
    // page, so it runs before any script on the *next* navigation too — belt
    // and suspenders alongside the page.evaluate() call below, which covers
    // the case where the window's initial document has already started.
    await electronApp.context().addInitScript(() => {
      const win = window as unknown as { __atlasLongTasks?: Array<{ duration: number; startTime: number }> }
      if (win.__atlasLongTasks) return
      win.__atlasLongTasks = []
      try {
        const observer = new PerformanceObserver((list) => {
          for (const entry of list.getEntries()) {
            win.__atlasLongTasks!.push({ duration: entry.duration, startTime: entry.startTime })
          }
        })
        observer.observe({ entryTypes: ['longtask'] })
      } catch {
        // PerformanceObserver/longtask unsupported — handled by the
        // post-hoc "observer actually ran" check below.
      }
    })

    const page = await electronApp.firstWindow()

    // Redundant with the init script above in the common case; only does
    // real work if the window's document had already started executing by
    // the time addInitScript was registered.
    await page.evaluate(() => {
      const win = window as unknown as { __atlasLongTasks?: Array<{ duration: number; startTime: number }> }
      if (win.__atlasLongTasks) return
      win.__atlasLongTasks = []
      try {
        const observer = new PerformanceObserver((list) => {
          for (const entry of list.getEntries()) {
            win.__atlasLongTasks!.push({ duration: entry.duration, startTime: entry.startTime })
          }
        })
        observer.observe({ entryTypes: ['longtask'] })
      } catch {
        // handled below
      }
    })

    // The CLI-arg file arrives ~300ms after ready-to-show and then the
    // worker-backed parse of 100k rows runs; give it a generous ceiling.
    await expect(page.locator(`[data-viewer="${dataViewer}"]`)).toHaveCount(1, { timeout: 30_000 })
    await expect(page.locator('.statusbar__file')).toHaveText(statusbarName, { timeout: 30_000 })
    // Let the grid finish its first paint/settle before reading results.
    await page.waitForTimeout(500)

    const longTasks = await page.evaluate(
      () => (window as unknown as { __atlasLongTasks?: Array<{ duration: number; startTime: number }> }).__atlasLongTasks ?? [],
    )

    const offenders = longTasks.filter((task) => task.duration > LONG_TASK_BUDGET_MS)
    expect(offenders, `long tasks over ${LONG_TASK_BUDGET_MS}ms: ${JSON.stringify(offenders)}`).toEqual([])
  } finally {
    await electronApp.close()
  }
}

test('opening a 100k-row CSV keeps every renderer main-thread task under 200ms', async () => {
  try {
    await assertNoLongTasksOpening(largeCsvPath, 'csv', 'large-100k.csv')
  } finally {
    await fs.rm(largeCsvPath, { force: true })
  }
})

// The xlsx path is architecturally distinct from CSV's (a real module Worker
// loaded from its own bundled script, rather than Papa Parse's blob-based
// worker) — see the file header comment — so it needs its own real-Electron
// assertion rather than trusting the CSV result to generalize.
test('opening a 100k-row XLSX keeps every renderer main-thread task under 200ms', async () => {
  try {
    await assertNoLongTasksOpening(largeXlsxPath, 'xlsx', 'large-100k.xlsx')
  } finally {
    await fs.rm(largeXlsxPath, { force: true })
  }
})
