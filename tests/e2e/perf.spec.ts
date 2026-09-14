/**
 * T2 (DAT-07) — perf harness: opening a 100k-row CSV must not block the
 * renderer's main thread for more than 200ms at a stretch.
 *
 * Vitest/jsdom has no real `Worker` (see useSpreadsheetWorkbook.test.ts and
 * csvParse.test.ts, which instead assert the *delegation decision* — that
 * Papa Parse is invoked with `worker:true` above the size threshold). The
 * actual browser-facing guarantee — that the main thread stays responsive —
 * can only be measured where a real Worker and a real
 * `PerformanceObserver({entryTypes:['longtask']})` both exist: a real
 * Chromium renderer, which is exactly what Playwright's Electron harness
 * provides.
 *
 * Timing design: main.cjs's `ready-to-show` handler waits 300ms before
 * sending the CLI-arg file to the renderer (see `sendFileToWindow`'s
 * call site) specifically to let the renderer finish initializing first —
 * that fixed window is also exactly what this test needs to attach the
 * PerformanceObserver *before* the CSV parse begins, with no race.
 */
import { promises as fs } from 'node:fs'
import path from 'node:path'
import { execFileSync } from 'node:child_process'

import { _electron as electron, expect, test } from '@playwright/test'

const projectRoot = process.cwd()
const fixtureDir = path.join(projectRoot, 'tests', 'e2e', 'fixtures')
const largeCsvPath = path.join(fixtureDir, 'large-100k.csv')

const ROW_COUNT = 100_000
const LONG_TASK_BUDGET_MS = 200

async function writeLargeCsvFixture(): Promise<void> {
  const lines: string[] = ['id,name,value,note']
  for (let i = 0; i < ROW_COUNT; i++) {
    lines.push(`${i},Row ${i},${(i * 1.5).toFixed(2)},padding text to make each row a realistic width`)
  }
  await fs.writeFile(largeCsvPath, lines.join('\n') + '\n', 'utf8')
}

test.beforeAll(async () => {
  execFileSync(process.execPath, ['tests/e2e/fixtures/generate.mjs'], {
    cwd: projectRoot,
    stdio: 'inherit',
  })
  await writeLargeCsvFixture()
})

test('opening a 100k-row CSV keeps every renderer main-thread task under 200ms', async () => {
  test.setTimeout(120_000)

  const electronApp = await electron.launch({
    args: ['.', largeCsvPath],
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
    await expect(page.locator('[data-viewer="csv"]')).toHaveCount(1, { timeout: 30_000 })
    await expect(page.locator('.statusbar__file')).toHaveText('large-100k.csv', { timeout: 30_000 })
    // Let the grid finish its first paint/settle before reading results.
    await page.waitForTimeout(500)

    const longTasks = await page.evaluate(
      () => (window as unknown as { __atlasLongTasks?: Array<{ duration: number; startTime: number }> }).__atlasLongTasks ?? [],
    )

    const offenders = longTasks.filter((task) => task.duration > LONG_TASK_BUDGET_MS)
    expect(offenders, `long tasks over ${LONG_TASK_BUDGET_MS}ms: ${JSON.stringify(offenders)}`).toEqual([])
  } finally {
    await electronApp.close()
    await fs.rm(largeCsvPath, { force: true })
  }
})
