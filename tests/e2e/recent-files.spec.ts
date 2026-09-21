import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execSync } from 'node:child_process'

import { _electron as electron, expect, test, type ElectronApplication } from '@playwright/test'

/**
 * TEST-7 — everything that lives under Electron's `userData` (recent files,
 * window state, and the persisted `recentFilesStore` ground truth behind
 * `recent:request-open`'s security check) was structurally untestable across
 * a relaunch: `electron/main.cjs` mints a brand-new temp profile directory on
 * every single `PLAYWRIGHT=1` launch (so tests never collide with each other
 * or with the owner's real profile, and so a not-yet-torn-down previous
 * instance's singleton lock never blocks the next test). That's the right
 * default — these specs don't change it — but it means no spec has ever
 * driven two launches that share a profile.
 *
 * `ATLAS_E2E_PROFILE_DIR` (added alongside these specs, see main.cjs) is an
 * opt-in override, gated behind the exact same `PLAYWRIGHT === '1' &&
 * !app.isPackaged` condition as the fresh-profile default: set it and two
 * `electron.launch()` calls share a profile; leave it unset (as every other
 * spec does) and behavior is unchanged.
 *
 * Sequencing two launches that share a profile is the actual hazard here
 * (see main.cjs's own header comment): a `taskkill /T` is not synchronous,
 * so if launch 2 starts before launch 1's process has actually exited, its
 * still-open singleton lock file makes launch 2 fail outright with "Lock
 * file can not be created! Error code: 32". `killAndWait` below kills with
 * `taskkill /PID <pid> /T /F` and then actually awaits the child process's
 * own `exit` event (not a fixed sleep standing in for it) before returning,
 * so every relaunch in this file is sequenced correctly by construction.
 */

const projectRoot = process.cwd()

function freshProfileDir(prefix: string): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), prefix))
}

/** Kills the app's real OS process and waits for it to actually exit — see
 * this file's header comment for why a fixed sleep isn't good enough here. */
async function killAndWait(app: ElectronApplication, timeoutMs = 10_000): Promise<void> {
  const proc = app.process()
  if (proc.exitCode !== null || proc.signalCode !== null) return
  const exited = new Promise<void>((resolve) => proc.once('exit', () => resolve()))
  try {
    execSync(`taskkill /PID ${proc.pid} /T /F`, { stdio: 'ignore' })
  } catch {
    try {
      proc.kill()
    } catch {
      // already gone
    }
  }
  await Promise.race([exited, new Promise<void>((resolve) => setTimeout(resolve, timeoutMs))])
}

/** Forces Chromium's DOM-storage backend (localStorage, backing the
 * renderer's own `atlas-recent` list) to write out to disk immediately,
 * instead of on its own internal commit timer — see the call site's comment
 * for why this matters before a `taskkill /F`. */
async function flushStorage(app: ElectronApplication): Promise<void> {
  await app.evaluate(({ BrowserWindow }) => {
    const win = BrowserWindow.getAllWindows()[0]
    win.webContents.session.flushStorageData()
  })
}

function launchWithProfile(profileDir: string, extraArgs: string[] = []): Promise<ElectronApplication> {
  return electron.launch({
    args: ['.', ...extraArgs],
    cwd: projectRoot,
    env: {
      ...process.env,
      CI: '1',
      PLAYWRIGHT: '1',
      ATLAS_E2E_PROFILE_DIR: profileDir,
    },
  })
}

test.describe('recent files and window state survive a relaunch (TEST-7)', () => {
  test('a file opened at launch 1 appears in Recent files at launch 2, and reopens correctly', async () => {
    const profileDir = freshProfileDir('atlas-e2e-recent-')
    const docsDir = freshProfileDir('atlas-e2e-recent-doc-')
    const filePath = path.join(docsDir, 'launch1-note.md')
    fs.writeFileSync(filePath, '# Opened in launch 1\n\nThis should survive a relaunch.')

    // Launch 1 — open the file as an argv path (the same trusted flow a
    // real OS double-click / "Open with Atlas" uses), which both adds it to
    // the renderer's own `atlas-recent` localStorage list (via
    // useFileHandler's `addRecent`) and records it in main's persisted
    // `recentFilesStore` (extractFilePath -> pathAllowlist.trust at argv
    // time, flushed to disk on `whenReady`; see main.cjs).
    const app1 = await launchWithProfile(profileDir, [filePath])
    try {
      const page1 = await app1.firstWindow()
      await expect(page1.locator('.statusbar__file')).toHaveText('launch1-note.md', { timeout: 20_000 })

      // The renderer's "Recent files" list lives in `localStorage`
      // (`atlas-recent`), which Chromium's DOM-storage backend commits to
      // disk asynchronously on its own schedule — a `taskkill /F` a moment
      // later can outrace that and kill the process before the write lands.
      // `flushStorageData()` forces it out now, so this test is exercising
      // "does the value survive a relaunch", not "did we get lucky with the
      // commit timer".
      await flushStorage(app1)
    } finally {
      await killAndWait(app1)
    }

    // Launch 2 — same profile, no file argument: lands on the Welcome
    // screen, which should now list launch 1's file under "Recent files".
    const app2 = await launchWithProfile(profileDir)
    try {
      const page2 = await app2.firstWindow()
      await expect(page2.getByRole('button', { name: 'Load Sample Document' })).toBeVisible({ timeout: 20_000 })

      const recentEntry = page2.locator('.recent__open', { hasText: 'launch1-note.md' })
      await expect(recentEntry).toBeVisible({ timeout: 15_000 })

      await recentEntry.click()

      await expect(page2.locator('.statusbar__file')).toHaveText('launch1-note.md', { timeout: 15_000 })
      await expect(page2.locator('[data-viewer="markdown"]')).toHaveCount(1, { timeout: 15_000 })
    } finally {
      await killAndWait(app2)
    }
  })

  test('window size and position are restored on the next launch', async () => {
    const profileDir = freshProfileDir('atlas-e2e-winstate-')

    const app1 = await launchWithProfile(profileDir)
    let targetBounds: { x: number; y: number; width: number; height: number }
    try {
      const page1 = await app1.firstWindow()
      await expect(page1.getByRole('button', { name: 'Load Sample Document' })).toBeVisible({ timeout: 20_000 })

      // Pick a size/position that's deliberately different from the
      // 1200x800 default and guaranteed to fit the primary display's work
      // area, so this can't pass by coincidence.
      targetBounds = await app1.evaluate(({ BrowserWindow, screen }) => {
        const win = BrowserWindow.getAllWindows()[0]
        const work = screen.getPrimaryDisplay().workArea
        const bounds = {
          x: work.x + 37,
          y: work.y + 29,
          width: Math.max(700, Math.min(760, work.width - 80)),
          height: Math.max(500, Math.min(560, work.height - 80)),
        }
        win.setBounds(bounds)
        return win.getBounds()
      })

      // The listener debounces saves by 300ms (WINDOW_STATE_SAVE_DEBOUNCE_MS)
      // — poll the real persisted file instead of a fixed sleep guessing at
      // that window.
      const stateFile = path.join(profileDir, 'window-state.json')
      await expect
        .poll(
          () => {
            try {
              const saved = JSON.parse(fs.readFileSync(stateFile, 'utf-8'))
              return saved.width === targetBounds.width && saved.height === targetBounds.height
            } catch {
              return false
            }
          },
          { timeout: 10_000 },
        )
        .toBe(true)
    } finally {
      await killAndWait(app1)
    }

    const app2 = await launchWithProfile(profileDir)
    try {
      const page2 = await app2.firstWindow()
      await expect(page2.getByRole('button', { name: 'Load Sample Document' })).toBeVisible({ timeout: 20_000 })

      const restored = await app2.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].getBounds())
      // A small tolerance, not exact equality: Windows/Chromium can round
      // outer-frame bounds by a few device pixels between a live
      // `setBounds()` call and a value re-applied at BrowserWindow
      // construction time (DPI scale-factor conversion), independent of
      // anything `windowState.cjs` does. The default launch is 1200x800 at
      // the origin, hundreds of pixels away from this test's target — this
      // tolerance is nowhere near wide enough to let that default pass.
      const TOLERANCE_PX = 8
      expect(Math.abs(restored.width - targetBounds.width)).toBeLessThanOrEqual(TOLERANCE_PX)
      expect(Math.abs(restored.height - targetBounds.height)).toBeLessThanOrEqual(TOLERANCE_PX)
      expect(Math.abs(restored.x - targetBounds.x)).toBeLessThanOrEqual(TOLERANCE_PX)
      expect(Math.abs(restored.y - targetBounds.y)).toBeLessThanOrEqual(TOLERANCE_PX)
    } finally {
      await killAndWait(app2)
    }
  })

  test('a maximized window is restored maximized on the next launch', async () => {
    const profileDir = freshProfileDir('atlas-e2e-winmax-')

    const app1 = await launchWithProfile(profileDir)
    try {
      const page1 = await app1.firstWindow()
      await expect(page1.getByRole('button', { name: 'Load Sample Document' })).toBeVisible({ timeout: 20_000 })

      await app1.evaluate(({ BrowserWindow }) => {
        BrowserWindow.getAllWindows()[0].maximize()
      })

      const stateFile = path.join(profileDir, 'window-state.json')
      await expect
        .poll(
          () => {
            try {
              const saved = JSON.parse(fs.readFileSync(stateFile, 'utf-8'))
              return saved.isMaximized === true
            } catch {
              return false
            }
          },
          { timeout: 10_000 },
        )
        .toBe(true)
    } finally {
      await killAndWait(app1)
    }

    const app2 = await launchWithProfile(profileDir)
    try {
      const page2 = await app2.firstWindow()
      await expect(page2.getByRole('button', { name: 'Load Sample Document' })).toBeVisible({ timeout: 20_000 })

      await expect
        .poll(() => app2.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].isMaximized()), {
          timeout: 10_000,
        })
        .toBe(true)
    } finally {
      await killAndWait(app2)
    }
  })

  // The actual security invariant (ELEC-02/03 fix) `recentFilesStore.cjs`
  // exists for: `recent:request-open` must accept a path that was genuinely
  // opened through a trusted flow in an EARLIER launch, and must refuse one
  // that never was — even though both paths equally exist on disk, so a
  // bare `fs.existsSync` check (what this replaced) can't tell them apart.
  // This is the one persistence gap the task singles out as not merely a
  // convenience feature: without it, a script that can only write to
  // `localStorage` could claim any file on disk is "recent" and get it
  // promoted into the write-eligible allowlist.
  test('recent:request-open accepts a path recorded in an earlier launch, and refuses one that was never recorded', async () => {
    const profileDir = freshProfileDir('atlas-e2e-recent-sec-')
    const docsDir = freshProfileDir('atlas-e2e-recent-sec-doc-')
    const trackedPath = path.join(docsDir, 'tracked.md')
    const untrackedPath = path.join(docsDir, 'untracked.md')
    fs.writeFileSync(trackedPath, '# Tracked — genuinely opened in launch 1')
    fs.writeFileSync(untrackedPath, '# Untracked — exists on disk, never opened by Atlas')

    // Launch 1 — open ONLY trackedPath (as an argv path, a trusted flow),
    // so only it lands in the persisted recentFilesStore. untrackedPath
    // exists on disk the whole time but Atlas never touches it — proving
    // the refusal below is about provenance, not file existence.
    const app1 = await launchWithProfile(profileDir, [trackedPath])
    try {
      const page1 = await app1.firstWindow()
      await expect(page1.locator('.statusbar__file')).toHaveText('tracked.md', { timeout: 20_000 })
    } finally {
      await killAndWait(app1)
    }

    // Launch 2 — same profile, fresh renderer (no in-memory allowlist state
    // carried over — only what's on disk in recentFilesStore.json does).
    const app2 = await launchWithProfile(profileDir)
    try {
      const page2 = await app2.firstWindow()
      await expect(page2.getByRole('button', { name: 'Load Sample Document' })).toBeVisible({ timeout: 20_000 })

      // `window.electronAPI` isn't declared in the e2e project's tsconfig
      // (that type lives under `src/`, out of this project's `include`) —
      // reach it the same untyped way this spec file's sibling
      // (`saveas-race.spec.ts`) already reaches main-process internals via
      // `app.evaluate`.
      const requestOpenRecent = (p: string) =>
        page2.evaluate((filePath) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          return (window as any).electronAPI.requestOpenRecent(filePath)
        }, p)

      expect(await requestOpenRecent(trackedPath)).toEqual({ ok: true })
      expect(await requestOpenRecent(untrackedPath)).toEqual({ ok: false })
    } finally {
      await killAndWait(app2)
    }
  })
})
