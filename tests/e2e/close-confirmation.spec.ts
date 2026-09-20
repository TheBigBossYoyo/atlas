import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execFileSync } from 'node:child_process'

import { _electron as electron, expect, test } from '@playwright/test'
import type { ElectronApplication, Page } from '@playwright/test'

/**
 * P4.6 (part) / P2.5 / SHELL-02 / ELEC-06 — live close-confirmation coverage.
 *
 * `App.shellSession.test.tsx` already unit-tests the renderer's half of this
 * round trip (pushing dirty state, answering a save-before-close request)
 * against a mocked `window.electronAPI`, and `closeGuard.test.ts` unit-tests
 * `electron/lib/closeGuard.cjs`'s pure decision logic. Neither exercises the
 * REAL main process: does a genuine `BrowserWindow` close request actually
 * get blocked while the document is dirty, does the real
 * `dialog.showMessageBoxSync` choice actually drive Save/Discard/Cancel, and
 * does the window actually stay open or actually close as a result. This
 * spec launches the real Electron app and drives a real `close()` call on
 * the real `BrowserWindow`.
 *
 * A real native message box can't be clicked through Playwright (it's an OS
 * window, not part of the Chromium page), so this "auto-answers" it the
 * standard Playwright-Electron way: `electronApp.evaluate()` runs code
 * inside the app's own main process, where `require('electron').dialog` is
 * the exact same singleton object `main.cjs` itself holds — stubbing
 * `dialog.showMessageBoxSync` there replaces the call main.cjs makes,
 * without ever needing a test-only code path inside main.cjs itself.
 */

const projectRoot = process.cwd()
const fixtureDir = path.join(projectRoot, 'tests', 'e2e', 'fixtures')

test.beforeAll(() => {
  execFileSync(process.execPath, ['tests/e2e/fixtures/generate.mjs'], {
    cwd: projectRoot,
    stdio: 'inherit',
  })
})

/** Copies sample.md into a scratch temp dir so a real Save-before-close
 * round trip writes a throwaway file instead of the git-tracked fixture. */
function copyMarkdownFixtureToTemp(): string {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-e2e-close-'))
  const dest = path.join(tempDir, 'sample.md')
  fs.copyFileSync(path.join(fixtureDir, 'sample.md'), dest)
  return dest
}

async function launchWithFile(filePath: string): Promise<{ app: ElectronApplication; page: Page }> {
  const app = await electron.launch({
    args: ['.', filePath],
    cwd: projectRoot,
    env: { ...process.env, CI: '1', PLAYWRIGHT: '1' },
  })
  const page = await app.firstWindow()
  return { app, page }
}

/** Stubs `dialog.showMessageBoxSync` in the app's real main process to
 * return `choiceIndex` (0=Save, 1=Discard, 2=Cancel — `CLOSE_PROMPT_CHOICE`
 * in electron/lib/closeGuard.cjs) instead of showing a real native dialog. */
async function autoAnswerCloseDialog(app: ElectronApplication, choiceIndex: number): Promise<void> {
  await app.evaluate(({ dialog }, choice) => {
    dialog.showMessageBoxSync = (() => choice) as typeof dialog.showMessageBoxSync
  }, choiceIndex)
}

/** Drives a real `close()` call on the app's real (only) BrowserWindow —
 * the same code path Electron itself invokes for a titlebar-X click. */
async function requestWindowClose(app: ElectronApplication): Promise<void> {
  await app.evaluate(({ BrowserWindow }) => {
    BrowserWindow.getAllWindows()[0]?.close()
  })
}

/** Playwright's own `electronApp.close()` teardown may itself trigger a real
 * `app.quit()` -> `BrowserWindow` 'close' sequence. If an earlier test left
 * the dialog stub answering Cancel (or the document still dirty with no
 * stub at all), that teardown call could hit the same guard and hang
 * instead of actually exiting. Called defensively before every `app.close()`
 * below so cleanup never depends on which scenario a test just exercised.
 */
async function forceCleanExit(app: ElectronApplication): Promise<void> {
  await app
    .evaluate(({ dialog }) => {
      dialog.showMessageBoxSync = (() => 1) as typeof dialog.showMessageBoxSync // DISCARD
    })
    .catch(() => {
      // The app may already be gone (a prior close succeeded) — nothing to clean up.
    })
}

async function makeMarkdownDirty(page: Page): Promise<void> {
  await page.getByRole('button', { name: 'Editor' }).click()
  const textarea = page.locator('.editor-panel__textarea')
  await expect(textarea).toBeVisible({ timeout: 10_000 })
  await textarea.fill('# Edited by the close-confirmation e2e spec\n')
}

test('a clean (non-dirty) document closes immediately with no prompt at all', async () => {
  const tempPath = copyMarkdownFixtureToTemp()
  const { app, page } = await launchWithFile(tempPath)
  try {
    await expect(page.locator('[data-viewer="markdown"]')).toHaveCount(1, { timeout: 15_000 })

    // No dialog stub needed — an un-stubbed showMessageBoxSync call here
    // would hang the test on a real native dialog, proving on its own that
    // decideOnClose correctly short-circuits before ever reaching it.
    const closedPromise = page.waitForEvent('close', { timeout: 15_000 })
    await requestWindowClose(app)
    await closedPromise
  } finally {
    await forceCleanExit(app)
    await app.close().catch(() => {})
    fs.rmSync(path.dirname(tempPath), { recursive: true, force: true })
  }
})

test('Cancel keeps the window open with the edit intact', async () => {
  const tempPath = copyMarkdownFixtureToTemp()
  const { app, page } = await launchWithFile(tempPath)
  try {
    await expect(page.locator('[data-viewer="markdown"]')).toHaveCount(1, { timeout: 15_000 })
    await makeMarkdownDirty(page)

    await autoAnswerCloseDialog(app, 2 /* CANCEL */)
    await requestWindowClose(app)

    // Give main's close handler a moment to run, then prove the window is
    // still genuinely alive and usable — not just "hasn't closed yet".
    await page.waitForTimeout(500)
    await expect(page.locator('.editor-panel__textarea')).toHaveValue(
      '# Edited by the close-confirmation e2e spec\n',
    )
    expect(fs.readFileSync(tempPath, 'utf-8')).not.toContain('Edited by the close-confirmation e2e spec')
  } finally {
    await forceCleanExit(app)
    await app.close().catch(() => {})
    fs.rmSync(path.dirname(tempPath), { recursive: true, force: true })
  }
})

test('Discard closes the window immediately without saving', async () => {
  const tempPath = copyMarkdownFixtureToTemp()
  const { app, page } = await launchWithFile(tempPath)
  try {
    await expect(page.locator('[data-viewer="markdown"]')).toHaveCount(1, { timeout: 15_000 })
    await makeMarkdownDirty(page)

    await autoAnswerCloseDialog(app, 1 /* DISCARD */)
    const closedPromise = page.waitForEvent('close', { timeout: 15_000 })
    await requestWindowClose(app)
    await closedPromise

    expect(fs.readFileSync(tempPath, 'utf-8')).not.toContain('Edited by the close-confirmation e2e spec')
  } finally {
    await app.close().catch(() => {})
    fs.rmSync(path.dirname(tempPath), { recursive: true, force: true })
  }
})

test('Save runs the real save-before-close round trip, writes the file, and then closes', async () => {
  const tempPath = copyMarkdownFixtureToTemp()
  const { app, page } = await launchWithFile(tempPath)
  try {
    await expect(page.locator('[data-viewer="markdown"]')).toHaveCount(1, { timeout: 15_000 })
    await makeMarkdownDirty(page)

    await autoAnswerCloseDialog(app, 0 /* SAVE */)
    const closedPromise = page.waitForEvent('close', { timeout: 15_000 })
    await requestWindowClose(app)
    await closedPromise

    // The real renderer save flow ran end to end (request-save-before-close
    // -> App's saveFile() -> window.electronAPI.saveFile -> the real
    // save-file IPC handler -> disk) before main destroyed the window.
    expect(fs.readFileSync(tempPath, 'utf-8')).toContain('Edited by the close-confirmation e2e spec')
  } finally {
    await app.close().catch(() => {})
    fs.rmSync(path.dirname(tempPath), { recursive: true, force: true })
  }
})

// The close in-flight guard itself (a rapid second close attempt while a
// Save round trip is pending must reuse it, not show a second dialog) is
// covered deterministically at the unit level instead of here —
// `main.ipc.test.ts`'s "close in-flight guard" suite drives the exact same
// `handleWindowCloseRequest` logic with synchronous mocks, with no risk of
// the real async save round trip and a real second CDP-evaluated close call
// racing each other and making this spec flaky.

// QUIT-DRAFT-1 — the bug this covers only shows in the real app: quitting
// used to skip the renderer entirely on Discard (`mainWindow.destroy()`
// called directly, with no IPC round trip at all), so the autosave draft in
// `localStorage` (src/hooks/useAutosave.ts) for the content just discarded
// survived the quit. A later, unrelated crash would then offer to "recover"
// content the user explicitly threw away. `App.shellSession.test.tsx`'s
// "main-process discard-before-close round trip" suite already covers the
// renderer's half against a mocked `window.electronAPI`, and
// `main.ipc.test.ts`'s "Discard round trip" suite covers main's half
// (notify-then-wait-then-destroy, bounded by a timeout) with synchronous
// mocks. Neither proves the real `request-discard-before-close` /
// `discard-before-close-result` IPC pair actually reaches a real renderer
// and a real `clearDraft()` before a real window teardown — this does.
//
// `PLAYWRIGHT=1` gives every launch its own fresh temp `userData` (see
// electron/main.cjs's top-of-file comment), and `localStorage` lives inside
// that profile — so a relaunch after this test's `app.close()` would open a
// brand-new, empty profile and trivially show no restore offer regardless of
// whether this fix works. That would not be a real assertion, so instead
// this reads `localStorage` directly, inside this one launch, in the window
// between "main decided the renderer's discard round trip is done" and the
// real window teardown actually running — the same real gap the fix closes.
//
// Racing a plain `page.evaluate()` read against the real teardown was tried
// first and is NOT used here: once the real window is actually gone,
// `evaluate()` calls issued around that moment can hang indefinitely rather
// than reject promptly (observed directly — this spec's own worker teardown
// hung for a further 60s on top of the test's own timeout). Instead, this
// temporarily replaces the real `BrowserWindow.destroy` and `app.quit` in
// the app's own main process (the same `app.evaluate` technique this file
// already uses to stub `dialog.showMessageBoxSync`) — `destroy` is delayed
// rather than skipped, and `app.quit` is no-op'd so the process itself
// cannot exit out from under the test — leaving a controlled window in which
// the page is still genuinely alive and its `localStorage` safely readable.
// `Object.defineProperty` is used instead of plain assignment: Electron's
// native `BrowserWindow.prototype.destroy` is non-writable, so a plain `win
// .destroy = ...` silently no-ops in sloppy mode (verified — that was this
// test's own first, silently-broken attempt) rather than shadowing it; an
// explicit own-property definition on the instance always wins regardless.
// Cleanup force-kills the process (`taskkill /T /F`, per this project's
// e2e convention) rather than `app.close()`, since a real `app.close()`
// waiting on a graceful quit that `app.quit` no longer performs would
// reintroduce the exact hang this approach avoids.
test('Quit + Discard clears the persisted autosave draft before the window is destroyed', async () => {
  test.setTimeout(30_000)
  const tempPath = copyMarkdownFixtureToTemp()
  const { app, page } = await launchWithFile(tempPath)
  try {
    await expect(page.locator('[data-viewer="markdown"]')).toHaveCount(1, { timeout: 15_000 })
    await makeMarkdownDirty(page)

    // Wait for the REAL ~800ms autosave debounce to actually persist a draft
    // of the EDIT specifically — proves the live autosave wiring, not a
    // hand-seeded localStorage entry. Polling for the edited text (rather
    // than merely "not null") matters here: autosave is enabled as soon as
    // any markdown document is open (`isMarkdownDocument`, not gated on
    // dirty — see the call site in App.tsx), so the freshly-opened file's
    // own unedited content can autosave first if its own debounce elapses
    // during this test's setup/click/fill overhead; polling for "not null"
    // alone could observe that transient unedited write instead of the
    // edit's own (later) settled one.
    await expect
      .poll(() => page.evaluate(() => localStorage.getItem('atlas-draft')), { timeout: 5_000 })
      .toContain('Edited by the close-confirmation e2e spec')

    await app.evaluate(({ app: electronApp, BrowserWindow }) => {
      const win = BrowserWindow.getAllWindows()[0]
      if (!win) return
      const g = globalThis as { __atlasDestroyCalled?: boolean }
      g.__atlasDestroyCalled = false
      const originalDestroy = win.destroy.bind(win)
      Object.defineProperty(win, 'destroy', {
        configurable: true,
        value: () => {
          g.__atlasDestroyCalled = true
          setTimeout(originalDestroy, 1000)
        },
      })
      // Prevents the real quit that 'window-all-closed' would otherwise
      // trigger once the (delayed) destroy above eventually runs for real,
      // so the whole process can't disappear out from under this test.
      electronApp.quit = (() => {}) as typeof electronApp.quit
    })

    await autoAnswerCloseDialog(app, 1 /* DISCARD */)
    await requestWindowClose(app)

    // Main only reaches destroy() once the renderer's discard round trip has
    // actually concluded (ack, or the bounded DISCARD_ACK_TIMEOUT_MS) —
    // waiting for this flag proves the round trip ran, rather than main
    // destroying immediately the way the pre-fix code did.
    await expect
      .poll(
        () => app.evaluate(() => (globalThis as { __atlasDestroyCalled?: boolean }).__atlasDestroyCalled === true),
        { timeout: 5_000 },
      )
      .toBe(true)

    // The real destroy() is still delayed (1s) — the page is genuinely alive
    // right now, so this reads the renderer's real post-discard state.
    const draftAfterDiscard = await page.evaluate(() => localStorage.getItem('atlas-draft'))
    expect(draftAfterDiscard).toBeNull()

    // The discarded edit was still never written to disk.
    expect(fs.readFileSync(tempPath, 'utf-8')).not.toContain('Edited by the close-confirmation e2e spec')
  } finally {
    try {
      execFileSync('taskkill', ['/PID', String(app.process().pid), '/T', '/F'], { stdio: 'ignore' })
    } catch {
      // Already gone.
    }
    fs.rmSync(path.dirname(tempPath), { recursive: true, force: true })
  }
})
