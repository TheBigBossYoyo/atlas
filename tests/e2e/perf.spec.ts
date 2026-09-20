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

/**
 * The budget above is the product bar, measured on a developer machine. The
 * observer also catches the app's own start-up work (module evaluation, JIT,
 * first React mount), and a shared CI runner is several times slower at that:
 * run 35469456280 reported 237 ms and 418 ms tasks ~0.5 s after start on a
 * commit that only touched documentation. A real regression here is an order
 * of magnitude, not a factor of two, so CI keeps a scaled bar rather than a
 * budget nobody can trust.
 */
const CI_LONG_TASK_SLACK = 2.5
const longTaskBudgetMs = (): number =>
  process.env.CI === 'true' || process.env.GITHUB_ACTIONS === 'true' ? LONG_TASK_BUDGET_MS * CI_LONG_TASK_SLACK : LONG_TASK_BUDGET_MS

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

    // The first window can still be on its initial blank document, which is
    // about to be replaced by the app's own: evaluating into it would race
    // that navigation (and, since Electron 44, reliably lose it). Waiting for
    // the app root means the document below is the final one — the init
    // script above has already installed the observer in it.
    await page.waitForSelector('#root', { timeout: 30_000 })
    // Everything before this point is start-up, not document work.
    const mountedAt = await page.evaluate(() => performance.now())

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

    const budget = longTaskBudgetMs()
    // Only what happened once the app was up: this assertion is about the
    // *document* not blocking the main thread. Start-up's own cost (module
    // evaluation, JIT, the first React mount) is a different question, with
    // its own test below — and on a shared CI runner it can alone exceed any
    // budget that is still meaningful for the parse.
    const offenders = longTasks.filter((task) => task.duration > budget && task.startTime >= mountedAt)
    expect(offenders, `long tasks over ${budget}ms: ${JSON.stringify(offenders)}`).toEqual([])
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

/**
 * PERF-01 — cold start (no file argv, just the welcome screen) must not
 * spend a long time evaluating JS before React's first commit.
 *
 * Before this task, `App.tsx` had a plain top-level
 * `import { exportMarkdown, ... } from './utils/export'` (pulling in `docx`,
 * `xlsx` via the shared spreadsheet parser, `html2canvas-pro`, and
 * `react-dom/server` — none of which the Export menu needs until a user
 * actually exports) and a plain top-level `import { MarkdownRenderer } from
 * './components/MarkdownRenderer'` (pulling in `katex`, `highlight.js`,
 * `dompurify`, `parse5`, and the whole react-markdown/rehype/remark chain —
 * none of which the welcome screen needs at all). Both got bundled into the
 * SAME eagerly-loaded entry chunk `main.tsx` itself sits in, every cold
 * start, for every format, whether or not that session ever opens a
 * markdown file or exports anything. Every other per-format viewer was
 * already code-split through `formats/registry.ts`'s `lazy()` loaders —
 * this made the export utilities and the markdown editor/preview split get
 * the same treatment, via `React.lazy()` (Suspense-wrapped, falling back to
 * the same `ViewerLoading` spinner every other viewer already shows) and a
 * `loadExportUtils = () => import('./utils/export')` called from inside the
 * two Export-menu handlers.
 *
 * Measured effect on `npx vite build`'s output: the main entry chunk
 * (`dist/assets/index-*.js`) dropped from ~2635 KB to ~295 KB (-89%), and
 * (this test's own metric) the single longest main-thread task recorded
 * between process launch and the welcome screen's first paint dropped from
 * ~155-190ms to ~50-56ms, measured back-to-back on the same machine under
 * the same concurrent-agent load, 5 runs each side.
 *
 * A `longtask` budget is used instead of raw wall-clock (unlike the simpler
 * open-to-visible timings used to develop this fix) because wall-clock also
 * bundles in OS process-spawn and disk-I/O variance that swamps the actual
 * JS-evaluation signal this regression cares about on a shared/loaded
 * machine — see the CSV/XLSX tests above for the same reasoning applied to
 * parse cost. The budget below sits well above the ~56ms this fix achieves
 * locally and well below the ~155-190ms a reverted regression produces
 * locally, then applies the same CI slack this file already uses for the
 * CSV/XLSX budget above.
 */
const STARTUP_LONG_TASK_BUDGET_MS = 120
const startupLongTaskBudgetMs = (): number =>
  process.env.CI === 'true' || process.env.GITHUB_ACTIONS === 'true'
    ? STARTUP_LONG_TASK_BUDGET_MS * CI_LONG_TASK_SLACK
    : STARTUP_LONG_TASK_BUDGET_MS

test('cold start to the welcome screen has no long main-thread task', async () => {
  test.setTimeout(60_000)

  const electronApp = await electron.launch({
    args: ['.'],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })

  try {
    // Same belt-and-suspenders init-script + page.evaluate pairing as
    // `assertNoLongTasksOpening` above — see its comment for why both are
    // needed to avoid a race with the window's own navigation.
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
        // handled by the offenders check below finding nothing either way
      }
    })

    const page = await electronApp.firstWindow()
    await page.waitForSelector('#root', { timeout: 30_000 })
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

    await expect(page.locator('.welcome')).toHaveCount(1, { timeout: 30_000 })
    await page.waitForTimeout(300)

    const longTasks = await page.evaluate(
      () => (window as unknown as { __atlasLongTasks?: Array<{ duration: number; startTime: number }> }).__atlasLongTasks ?? [],
    )

    const budget = startupLongTaskBudgetMs()
    const offenders = longTasks.filter((task) => task.duration > budget)
    expect(offenders, `long tasks over ${budget}ms before the welcome screen appeared: ${JSON.stringify(offenders)}`).toEqual([])
  } finally {
    await electronApp.close()
  }
})

/**
 * D23 — typing in a long DOCX must not stall. Every keystroke re-paginates
 * the document; before the per-paragraph line cache and the time-sliced
 * yielding, a single character on a ~35-page document took about three
 * seconds (most of it asleep, waiting on requestAnimationFrame, which
 * Chromium throttles to ~1 Hz whenever the window is not focused — exactly
 * the situation a Playwright-driven window is in). Wave 4 (a28f3ea) brought
 * that down to ~500ms.
 *
 * D23-PERF: the ~500ms left over was mostly wasted React work, not layout —
 * `documentModel` (which changes every keystroke) was passed straight
 * through to `PageStack`'s `document` prop, so every one of this test's ~24
 * pages re-rendered its whole-document run/bookmark-metadata walk TWICE per
 * keystroke: once the instant the command applied (still showing the OLD
 * `pages`), and again once pagination actually produced new `pages`. Pinning
 * `PageStack`'s `document` prop to the exact document a completed `pages`
 * array was laid out against (`pagesDocument` in DocxViewer.tsx), hoisting
 * the per-document metadata walk out of `PageView` into `PageStack` (once
 * per document instead of once per PAGE), and memoizing both components cut
 * this to one render pass per keystroke — confirmed by an instrumented
 * build showing 96 `PageView` renders/keystroke (4 × 24 pages) before this
 * fix and 24 (1 × 24, the unavoidable minimum once `pages` itself changes)
 * after. Measured wall-clock median on a dev machine dropped from
 * ~250-400ms to ~160-200ms (see the D23-PERF commit for the full
 * before/after methodology).
 *
 * D23-PERF-2: measuring the same keystroke split into command-apply/
 * pagination/React-commit/browser-layout-paint (performance.mark/measure
 * instrumentation, removed before this commit) on this machine, under the
 * SAME concurrent-agent CPU load that made the numbers above noisier than
 * a quiet machine, found commit (~40-50ms) and paint (~50-60ms) dominating
 * — NOT pagination (~15-25ms, already cheap thanks to the D23 line cache).
 * The cause: `paginate()` returns a brand-new `Page` object for every page
 * on every call, so even with `PageStack`/`PageView` memoized (D23-PERF),
 * every one of this test's ~24 pages still fails its memo check and
 * remounts a full DOM subtree on every keystroke, regardless of whether
 * that page is even on screen.
 *
 * Fix: `PageStack` now virtualizes — only pages within the scroll
 * container's viewport (plus a buffer, plus whichever page holds the
 * caret/selection) actually mount `PageView`; the rest render as a
 * lightweight placeholder that reserves the same box (see
 * `src/docx/render/pageVirtualization.ts` and `PageStack.tsx`). Measured
 * back-to-back on this machine, same method, same load: commit dropped to
 * ~4-10ms and paint to ~10-20ms; wall-clock median across 5 keystrokes went
 * from ~280-300ms to ~141-155ms.
 *
 * The budget below keeps generous margin over that ~155ms median — CI
 * hardware is slower than a developer machine, and typing latency here is
 * noisy under any concurrent CPU load (it has flaked at roughly 2x the
 * local numbers before) — but is tightened from 1_200ms now that
 * virtualization moved the achieved number that much further below it.
 */
const TYPING_BUDGET_MS = 600

test('typing in a 35-page DOCX stays responsive', async () => {
  const { Document, Packer, Paragraph, TextRun } = await import('docx')
  const children = Array.from(
    { length: 300 },
    (_, index) =>
      new Paragraph({
        children: [
          new TextRun(
            `Paragraph ${index + 1}. ${'The quick brown fox jumps over the lazy dog while the editor keeps every keystroke responsive. '.repeat(3)}`,
          ),
        ],
      }),
  )
  const fixture = path.join(fixtureDir, 'long-typing.docx')
  await fs.writeFile(fixture, await Packer.toBuffer(new Document({ sections: [{ children }] })))

  const app = await electron.launch({
    args: ['.', fixture],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  try {
    const page = await app.firstWindow()
    await page.waitForSelector('[data-paragraph-path]', { timeout: 60_000 })
    await page.waitForFunction(() => document.querySelectorAll('.docx-page').length > 5, null, { timeout: 60_000 })
    await page.waitForTimeout(1_500)
    // `.docx-page` counts BOTH real pages and virtualized placeholders (see
    // D23-PERF-2 / PageStack.tsx — a placeholder reserves the exact same box
    // a real page would, so this total still reflects the whole document).
    const totalPages = await page.locator('.docx-page').count()
    expect(totalPages).toBeGreaterThan(20)

    // D23-PERF-2 — the actual point of virtualization: with the window this
    // test runs at, only a handful of the 20+ pages should be REAL PageViews
    // at any one time, not all of them. A regression here (e.g. a future
    // change accidentally disabling virtualization) would silently give back
    // the exact DOM-mount cost this task set out to cut.
    const realPages = await page.locator('.docx-page:not([data-virtualized="placeholder"])').count()
    expect(realPages, `expected fewer than ${totalPages} real pages mounted, got ${realPages}`).toBeLessThan(totalPages)

    await page.locator('[data-paragraph-path="1"]').first().click()
    await page.waitForTimeout(200)

    const samples: number[] = []
    for (let i = 0; i < 5; i++) {
      const started = Date.now()
      await page.keyboard.type('z')
      await page.waitForFunction(
        (count) => (document.querySelector('[data-paragraph-path="1"]')?.textContent ?? '').includes('z'.repeat(count)),
        i + 1,
        { timeout: 30_000, polling: 10 },
      )
      samples.push(Date.now() - started)
    }

    const median = [...samples].sort((a, b) => a - b)[Math.floor(samples.length / 2)]
    expect(median, `keystroke latencies: ${samples.join(', ')} ms`).toBeLessThan(TYPING_BUDGET_MS)
  } finally {
    try {
      execFileSync('taskkill', ['/PID', String(app.process().pid), '/T', '/F'], { stdio: 'ignore' })
    } catch {
      app.process().kill()
    }
    await fs.rm(path.join(fixtureDir, 'long-typing.docx'), { force: true })
  }
})
