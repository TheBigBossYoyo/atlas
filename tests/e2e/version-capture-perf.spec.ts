/**
 * VERSIONS-2 — capturing a large document must not block the thread the user is
 * typing on.
 *
 * The periodic capture re-serialises the whole document every two minutes while
 * it is dirty. For markdown that was free (the text was already a string); for a
 * `.docx` it means rebuilding `document.xml` and re-zipping the package, and
 * this repo holds renderer main-thread work to a 200 ms long-task budget
 * (`perf.spec.ts`). A capture that blew through that would make the feature feel
 * like a stutter every two minutes — the kind of cost that is invisible in a
 * correctness test and obvious to the person using it.
 *
 * So this measures it, on a genuinely large document (36 pages with images, the
 * same fixture the page-virtualization work used), and fails if the capture
 * costs a long task over the same budget `perf.spec.ts` uses.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execFileSync, execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication, type Page } from '@playwright/test'

const projectRoot = process.cwd()

// The same budget and the same CI slack as `perf.spec.ts`, for the same reasons
// its own comment gives: a shared runner is several times slower at this, and a
// real regression here is an order of magnitude rather than a factor of two.
const LONG_TASK_BUDGET_MS = 200
const CI_LONG_TASK_SLACK = 2.5
const longTaskBudgetMs = (): number =>
  process.env.CI === 'true' || process.env.GITHUB_ACTIONS === 'true'
    ? LONG_TASK_BUDGET_MS * CI_LONG_TASK_SLACK
    : LONG_TASK_BUDGET_MS

type LongTask = { duration: number; startTime: number }

function kill(app: ElectronApplication): void {
  try {
    execSync(`taskkill /PID ${app.process().pid} /T /F`, { stdio: 'ignore' })
  } catch {
    app.process().kill()
  }
}

function installObserver(): void {
  const win = window as unknown as { __atlasLongTasks?: LongTask[] }
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
    // Unsupported — caught by the "the observer actually ran" check below.
  }
}

async function longTasks(page: Page): Promise<LongTask[]> {
  return page.evaluate(() => (window as unknown as { __atlasLongTasks?: LongTask[] }).__atlasLongTasks ?? [])
}

/**
 * Proves the observer is alive by deliberately blocking the main thread.
 *
 * Without this, "no long task was recorded" is ambiguous — it reads the same
 * whether the capture was genuinely cheap or the PerformanceObserver never
 * installed, and the second reading turns this whole test into one that cannot
 * fail. `perf.spec.ts` gets its liveness for free because opening a 100k-row
 * file blocks the thread on its own; a capture might not, which is the result
 * we are hoping for.
 */
async function assertObserverIsAlive(page: Page): Promise<void> {
  // The block has to happen inside a TIMER task, not directly in this
  // CDP-driven evaluate: work the debugger protocol drives is not attributed to
  // the page's event loop, so the long-tasks API does not see it at all. That
  // cost a wrong conclusion once already — a deliberate 150ms block reported
  // nothing and looked like a broken observer. Verified against a real renderer:
  // the same block inside setTimeout is reported, at 150ms.
  await page.evaluate(
    () =>
      new Promise<void>((resolve) => {
        setTimeout(() => {
          const started = performance.now()
          while (performance.now() - started < 150) {
            // Deliberately blocking.
          }
          resolve()
        }, 0)
      }),
  )
  await page.waitForTimeout(300)
  const seen = await longTasks(page)
  const blocking = seen.filter((task) => task.duration > 100)
  expect(
    blocking.length,
    `the long-task observer did not report a deliberate 150ms block, so it is not measuring anything: ${JSON.stringify(seen)}`,
  ).toBeGreaterThan(0)
}

let fixture = ''

test.beforeAll(() => {
  // Run as a subprocess, the same way `docx-scale.spec.ts` and
  // `close-race.spec.ts` build this fixture — the generator is a plain .mjs with
  // no type declarations, so importing it would mean an implicit `any`.
  fixture = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-capture-perf-')), 'large.docx')
  execFileSync(
    process.execPath,
    [path.join(projectRoot, 'tests', 'e2e', 'fixtures', 'generateLargeDocx.mjs'), fixture, '36'],
    { cwd: projectRoot, stdio: 'inherit' },
  )
})

test('capturing a 36-page DOCX stays within the renderer long-task budget', async () => {
  test.setTimeout(180_000)

  const app = await electron.launch({
    args: ['.', fixture],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  try {
    await app.context().addInitScript(installObserver)
    const page = await app.firstWindow()
    await page.waitForSelector('[data-paragraph-path]', { timeout: 60_000 })
    await page.evaluate(installObserver)
    await page.waitForTimeout(1500)

    // One edit, so there is something to capture and the document is dirty.
    await page.locator('[data-paragraph-path="0"] .docx-run').first().click()
    await page.keyboard.type('Measured. ', { delay: 20 })
    await page.waitForTimeout(800)

    await assertObserverIsAlive(page)

    // Everything up to here — start-up, the 36-page parse, the first paint, the
    // edit itself, and the deliberate block above — has its own budgets
    // elsewhere. Only the capture is on trial.
    const beforeCapture = await page.evaluate(() => performance.now())

    await page.getByRole('button', { name: /Version history/i }).click()
    await expect(page.locator('.version-history')).toBeVisible({ timeout: 10_000 })
    const captureButton = page.locator('.version-history').getByRole('button', { name: /Save a version now/i })
    await expect(captureButton).toBeVisible()

    const startedAt = Date.now()
    await captureButton.click()
    // The version appearing is how we know the serialisation finished, rather
    // than measuring a window that the work had not yet entered.
    await expect
      .poll(async () => page.locator('.version-history__item').count(), { timeout: 60_000 })
      .toBeGreaterThanOrEqual(1)
    const wallClockMs = Date.now() - startedAt

    const observed = await longTasks(page)
    const duringCapture = observed.filter((task) => task.startTime >= beforeCapture)
    const budget = longTaskBudgetMs()
    const offenders = duringCapture.filter((task) => task.duration > budget)

    const longest = duringCapture.reduce((max, task) => Math.max(max, task.duration), 0)
    // On the record either way, so the cost is a measurement rather than
    // something inferred from a passing assertion.
    const measurement =
      `36-page DOCX capture: ${wallClockMs}ms wall clock, ` +
      `${duringCapture.length} long task(s), longest ${Math.round(longest)}ms (budget ${budget}ms)`
    test.info().annotations.push({ type: 'measurement', description: measurement })
    console.log(`[VERSIONS-2] ${measurement}`)

    expect(offenders, `${measurement}; offenders: ${JSON.stringify(offenders)}`).toEqual([])
  } finally {
    kill(app)
  }
})
