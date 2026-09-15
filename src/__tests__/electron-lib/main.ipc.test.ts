/**
 * electron/main.cjs — IPC handler argument validation, sender-frame checks,
 * and allowlist gating (P2.13 IPC part / SHELL-27, QA-05).
 *
 * The fake-`electron`-module harness (why it's needed, and how it works) is
 * shared with `main.ipc.misc.test.ts` — see `mainIpcTestHarness.ts`'s own
 * header for the full explanation.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { CLOSE_PROMPT_CHOICE } from '../../../electron/lib/closeGuard.cjs'
import { STATE_FILE_NAME, loadWindowState } from '../../../electron/lib/windowState.cjs'
import {
  ALLOWED_EVENT,
  ELECTRON_MODULE_PATH,
  MAIN_CJS_PATH,
  REJECTED_EVENT,
  createElectronMock,
  loadMainCjs,
  nodeRequire,
  type IpcHandler,
} from './mainIpcTestHarness'

describe('electron/main.cjs IPC handlers', () => {
  let tempDir: string
  let mocks: ReturnType<typeof createElectronMock>
  // main.cjs registers `process.on('uncaughtException'/'unhandledRejection', ...)`
  // at module scope (P1.15/ELEC-13) with no matching removeListener — it's
  // designed as a singleton app entry point, not something re-required many
  // times. Requiring it fresh in every test would otherwise accumulate
  // listeners on the real shared `process` object across the whole suite.
  let uncaughtBefore: number
  let unhandledBefore: number

  beforeEach(async () => {
    uncaughtBefore = process.listenerCount('uncaughtException')
    unhandledBefore = process.listenerCount('unhandledRejection')
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-main-ipc-'))
    mocks = createElectronMock(tempDir)
    await loadMainCjs(mocks)
  })

  afterEach(() => {
    delete nodeRequire.cache[ELECTRON_MODULE_PATH]
    delete nodeRequire.cache[MAIN_CJS_PATH]
    fs.rmSync(tempDir, { recursive: true, force: true })

    // Strip off exactly the listeners this test's fresh require added,
    // leaving vitest's own process-level listeners untouched.
    for (const listener of process.listeners('uncaughtException').slice(uncaughtBefore)) {
      process.removeListener('uncaughtException', listener);
    }
    for (const listener of process.listeners('unhandledRejection').slice(unhandledBefore)) {
      process.removeListener('unhandledRejection', listener);
    }
  })

  function handler(channel: string): IpcHandler {
    const fn = mocks.ipcHandlers.get(channel)
    if (!fn) throw new Error(`no handler registered for "${channel}"`)
    return fn
  }

  /** Every `mainWindow.on(event, ...)` listener registered for `event`, in registration order. */
  function windowHandlers(event: string): IpcHandler[] {
    return mocks.windowOnHandlers.get(event) ?? []
  }

  it('creates the window and registers the expected IPC channels', () => {
    expect(mocks.ipcHandlers.has('file:readBinaryByPath')).toBe(true)
    expect(mocks.ipcHandlers.has('save-file')).toBe(true)
    expect(mocks.ipcHandlers.has('save-binary-file')).toBe(true)
    expect(mocks.ipcHandlers.has('shell:reveal-in-folder')).toBe(true)
  })

  describe('dialog:openFileBinary', () => {
    it('rejects a request not from the main frame', async () => {
      const result = (await handler('dialog:openFileBinary')(REJECTED_EVENT)) as { canceled: boolean }
      expect(result.canceled).toBe(true)
      expect(mocks.dialog.showOpenDialog).not.toHaveBeenCalled()
    })

    it('returns canceled when the dialog is dismissed', async () => {
      mocks.dialog.showOpenDialog.mockResolvedValueOnce({ canceled: true, filePaths: [] })
      const result = (await handler('dialog:openFileBinary')(ALLOWED_EVENT)) as { canceled: boolean }
      expect(result.canceled).toBe(true)
    })

    it('reads and allowlists the chosen file, and passes the manifest extension list as the filter', async () => {
      const filePath = path.join(tempDir, 'report.docx')
      fs.writeFileSync(filePath, 'hello world')
      mocks.dialog.showOpenDialog.mockResolvedValueOnce({ canceled: false, filePaths: [filePath] })

      const result = (await handler('dialog:openFileBinary')(ALLOWED_EVENT)) as {
        canceled: boolean
        path: string
        buffer: ArrayBuffer
      }

      expect(result.canceled).toBe(false)
      expect(result.path).toBe(filePath)
      expect(Buffer.from(result.buffer).toString()).toBe('hello world')

      const filterArg = mocks.dialog.showOpenDialog.mock.calls[0]?.[1] as { filters: Array<{ extensions: string[] }> }
      expect(filterArg.filters[0]?.extensions).toContain('docx')
      expect(filterArg.filters[0]?.extensions).toContain('mdown')

      // The file is now allowlisted — file:readBinaryByPath should succeed.
      const readResult = (await handler('file:readBinaryByPath')(ALLOWED_EVENT, filePath)) as { buffer: ArrayBuffer }
      expect(Buffer.from(readResult.buffer).toString()).toBe('hello world')
    })

    it('surfaces a friendly error and does not allowlist a file over the size cap', async () => {
      const filePath = path.join(tempDir, 'huge.docx')
      fs.writeFileSync(filePath, 'x')
      const statSpy = vi.spyOn(fs, 'statSync').mockReturnValue({ size: 500 * 1024 * 1024 } as fs.Stats)
      mocks.dialog.showOpenDialog.mockResolvedValueOnce({ canceled: false, filePaths: [filePath] })

      const result = (await handler('dialog:openFileBinary')(ALLOWED_EVENT)) as { canceled: boolean }

      expect(result.canceled).toBe(true)
      expect(mocks.dialog.showErrorBox).toHaveBeenCalledTimes(1)
      expect(mocks.dialog.showErrorBox.mock.calls[0]?.[1]).toMatch(/too large/i)

      const readAttempt = Promise.resolve(handler('file:readBinaryByPath')(ALLOWED_EVENT, filePath))
      const readResult = await readAttempt.catch((err: unknown) => err)
      expect(readResult).toBeInstanceOf(Error)
      statSpy.mockRestore()
    })
  })

  describe('file:readBinaryByPath', () => {
    it('rejects a request not from the main frame', async () => {
      await expect(handler('file:readBinaryByPath')(REJECTED_EVENT, 'C:/x.docx')).rejects.toThrow(
        /could not be verified/i,
      )
    })

    it('rejects a path that was never allowlisted', async () => {
      const filePath = path.join(tempDir, 'never-opened.docx')
      fs.writeFileSync(filePath, 'x')
      await expect(handler('file:readBinaryByPath')(ALLOWED_EVENT, filePath)).rejects.toThrow(
        /not selected through Atlas/i,
      )
    })

    it('rejects a non-string path', async () => {
      await expect(handler('file:readBinaryByPath')(ALLOWED_EVENT, 42)).rejects.toThrow(/not selected through Atlas/i)
    })
  })

  describe('open-file-by-path', () => {
    it('rejects a request not from the main frame', () => {
      expect(() => handler('open-file-by-path')(REJECTED_EVENT, 'C:/x.md')).toThrow(/could not be verified/i)
    })

    it('rejects a path that was never allowlisted', () => {
      const filePath = path.join(tempDir, 'never-opened.md')
      fs.writeFileSync(filePath, '# hi')
      expect(() => handler('open-file-by-path')(ALLOWED_EVENT, filePath)).toThrow(/not selected through Atlas/i)
    })

    it('reads an allowlisted markdown file', async () => {
      const filePath = path.join(tempDir, 'notes.md')
      fs.writeFileSync(filePath, '# Notes')
      mocks.dialog.showOpenDialog.mockResolvedValueOnce({ canceled: false, filePaths: [filePath] })
      await handler('dialog:openFileBinary')(ALLOWED_EVENT)

      const result = handler('open-file-by-path')(ALLOWED_EVENT, filePath) as { content: string } | null
      expect(result?.content).toBe('# Notes')
    })
  })

  describe('path:register-dropped', () => {
    it('rejects a request not from the main frame', async () => {
      const result = (await handler('path:register-dropped')(REJECTED_EVENT, 'C:/x.md')) as { ok: boolean }
      expect(result.ok).toBe(false)
    })

    it('rejects a path that does not exist on disk', async () => {
      const result = (await handler('path:register-dropped')(ALLOWED_EVENT, path.join(tempDir, 'ghost.md'))) as {
        ok: boolean
      }
      expect(result.ok).toBe(false)
    })

    it('allowlists a real dropped path', async () => {
      const filePath = path.join(tempDir, 'dropped.md')
      fs.writeFileSync(filePath, '# Dropped')
      const result = (await handler('path:register-dropped')(ALLOWED_EVENT, filePath)) as { ok: boolean }
      expect(result.ok).toBe(true)

      const readResult = handler('open-file-by-path')(ALLOWED_EVENT, filePath) as { content: string } | null
      expect(readResult?.content).toBe('# Dropped')
    })
  })

  describe('recent:request-open', () => {
    it('rejects a path never recorded as recent, even if it exists on disk', async () => {
      const filePath = path.join(tempDir, 'exists-but-not-recent.md')
      fs.writeFileSync(filePath, '# hi')
      const result = (await handler('recent:request-open')(ALLOWED_EVENT, filePath)) as { ok: boolean }
      expect(result.ok).toBe(false)
    })

    it('allows a path previously recorded via a trusted open flow (dialog)', async () => {
      const filePath = path.join(tempDir, 'previously-opened.md')
      fs.writeFileSync(filePath, '# hi')
      mocks.dialog.showOpenDialog.mockResolvedValueOnce({ canceled: false, filePaths: [filePath] })
      await handler('dialog:openFileBinary')(ALLOWED_EVENT)

      const result = (await handler('recent:request-open')(ALLOWED_EVENT, filePath)) as { ok: boolean }
      expect(result.ok).toBe(true)
    })
  })

  describe('save-file', () => {
    it('rejects a request not from the main frame', async () => {
      const result = (await handler('save-file')(REJECTED_EVENT, { content: 'x', suggestedName: 'a.md' })) as {
        saved: boolean
        error?: string
      }
      expect(result.saved).toBe(false)
      expect(result.error).toMatch(/could not be verified/i)
    })

    it('rejects a malformed request (non-string content)', async () => {
      const result = (await handler('save-file')(ALLOWED_EVENT, { content: 42 })) as { saved: boolean }
      expect(result.saved).toBe(false)
    })

    it('silently overwrites an allowlisted existingPath without opening the save dialog', async () => {
      const filePath = path.join(tempDir, 'doc.md')
      fs.writeFileSync(filePath, 'old content')
      mocks.dialog.showOpenDialog.mockResolvedValueOnce({ canceled: false, filePaths: [filePath] })
      await handler('dialog:openFileBinary')(ALLOWED_EVENT)

      const result = (await handler('save-file')(ALLOWED_EVENT, {
        content: 'new content',
        suggestedName: 'doc.md',
        existingPath: filePath,
      })) as { saved: boolean; path?: string }

      expect(result.saved).toBe(true)
      expect(mocks.dialog.showSaveDialog).not.toHaveBeenCalled()
      expect(fs.readFileSync(filePath, 'utf-8')).toBe('new content')
    })

    it('falls back to the save dialog for a non-allowlisted existingPath', async () => {
      const filePath = path.join(tempDir, 'via-dialog.md')
      mocks.dialog.showSaveDialog.mockResolvedValueOnce({ canceled: false, filePath })

      const result = (await handler('save-file')(ALLOWED_EVENT, {
        content: 'hello',
        suggestedName: 'via-dialog.md',
        existingPath: path.join(tempDir, 'not-allowlisted.md'),
      })) as { saved: boolean }

      expect(mocks.dialog.showSaveDialog).toHaveBeenCalledTimes(1)
      expect(result.saved).toBe(true)
      expect(fs.readFileSync(filePath, 'utf-8')).toBe('hello')
    })

    it('returns saved:false when the save dialog is canceled', async () => {
      mocks.dialog.showSaveDialog.mockResolvedValueOnce({ canceled: true })
      const result = (await handler('save-file')(ALLOWED_EVENT, { content: 'x', suggestedName: 'a.md' })) as {
        saved: boolean
      }
      expect(result.saved).toBe(false)
    })
  })

  describe('save-binary-file', () => {
    it('rejects a request whose content is not a Uint8Array', async () => {
      const result = (await handler('save-binary-file')(ALLOWED_EVENT, { content: 'not-bytes' })) as {
        saved: boolean
      }
      expect(result.saved).toBe(false)
    })

    it('writes bytes to an allowlisted existingPath', async () => {
      const filePath = path.join(tempDir, 'doc.docx')
      fs.writeFileSync(filePath, 'old')
      mocks.dialog.showOpenDialog.mockResolvedValueOnce({ canceled: false, filePaths: [filePath] })
      await handler('dialog:openFileBinary')(ALLOWED_EVENT)

      const bytes = new Uint8Array([1, 2, 3, 4])
      const result = (await handler('save-binary-file')(ALLOWED_EVENT, {
        content: bytes,
        suggestedName: 'doc.docx',
        existingPath: filePath,
      })) as { saved: boolean }

      expect(result.saved).toBe(true)
      expect([...fs.readFileSync(filePath)]).toEqual([1, 2, 3, 4])
    })
  })

  describe('image:pick', () => {
    it('rejects a request not from the main frame', async () => {
      const result = (await handler('image:pick')(REJECTED_EVENT)) as { cancelled: boolean }
      expect(result.cancelled).toBe(true)
      expect(mocks.dialog.showOpenDialog).not.toHaveBeenCalled()
    })

    it('returns cancelled when the dialog is dismissed', async () => {
      mocks.dialog.showOpenDialog.mockResolvedValueOnce({ canceled: true, filePaths: [] })
      const result = (await handler('image:pick')(ALLOWED_EVENT)) as { cancelled: boolean }
      expect(result.cancelled).toBe(true)
    })

    it('rejects an image over its own (smaller) size cap with a friendly message instead of reading it (ELEC-20)', async () => {
      const filePath = path.join(tempDir, 'huge.png')
      fs.writeFileSync(filePath, 'x')
      const statSpy = vi.spyOn(fs, 'statSync').mockReturnValue({ size: 26 * 1024 * 1024 } as fs.Stats)
      mocks.dialog.showOpenDialog.mockResolvedValueOnce({ canceled: false, filePaths: [filePath] })

      const result = (await handler('image:pick')(ALLOWED_EVENT)) as { cancelled: boolean; error?: string }

      expect(result.cancelled).toBe(true)
      expect(result.error).toMatch(/too large/i)
      statSpy.mockRestore()
    })

    it('returns image bytes and mime type for a valid pick', async () => {
      const filePath = path.join(tempDir, 'pic.png')
      fs.writeFileSync(filePath, Buffer.from([0x89, 0x50, 0x4e, 0x47]))
      mocks.dialog.showOpenDialog.mockResolvedValueOnce({ canceled: false, filePaths: [filePath] })

      const result = (await handler('image:pick')(ALLOWED_EVENT)) as {
        cancelled: boolean
        mime?: string
        bytes?: Uint8Array
      }

      expect(result.cancelled).toBe(false)
      expect(result.mime).toBe('image/png')
      expect(result.bytes && [...result.bytes]).toEqual([0x89, 0x50, 0x4e, 0x47])
    })
  })

  describe('shell:reveal-in-folder', () => {
    it('rejects a request not from the main frame', async () => {
      const result = (await handler('shell:reveal-in-folder')(REJECTED_EVENT, 'C:/x.docx')) as { ok: boolean }
      expect(result.ok).toBe(false)
      expect(mocks.shell.showItemInFolder).not.toHaveBeenCalled()
    })

    it('rejects a path that was never allowlisted', async () => {
      const result = (await handler('shell:reveal-in-folder')(ALLOWED_EVENT, path.join(tempDir, 'nope.bin'))) as {
        ok: boolean
      }
      expect(result.ok).toBe(false)
      expect(mocks.shell.showItemInFolder).not.toHaveBeenCalled()
    })

    it('rejects a non-string path', async () => {
      const result = (await handler('shell:reveal-in-folder')(ALLOWED_EVENT, 123)) as { ok: boolean }
      expect(result.ok).toBe(false)
    })

    it('reveals an allowlisted path', async () => {
      const filePath = path.join(tempDir, 'weird.bin')
      fs.writeFileSync(filePath, 'x')
      mocks.dialog.showOpenDialog.mockResolvedValueOnce({ canceled: false, filePaths: [filePath] })
      await handler('dialog:openFileBinary')(ALLOWED_EVENT)

      const result = (await handler('shell:reveal-in-folder')(ALLOWED_EVENT, filePath)) as { ok: boolean }
      expect(result.ok).toBe(true)
      expect(mocks.shell.showItemInFolder).toHaveBeenCalledWith(filePath)
    })
  })

  // wave-3 shell-polish follow-up — window bounds/maximized state is saved on
  // every resize/move (debounced, not exercised here) and on every close
  // attempt (immediate). These drive the real `mainWindow.on(...)` listeners
  // captured by the mock, not an IPC channel.
  describe('window state persistence (P5.2/ELEC-11)', () => {
    it('persists the current bounds/maximized state to disk on a close attempt', () => {
      const closeListeners = windowHandlers('close')
      expect(closeListeners.length).toBeGreaterThan(0)

      // Every 'close' listener runs on a real close attempt (main.cjs
      // registers persistWindowState AND handleWindowCloseRequest on the
      // same event) — invoke them all, exactly as Electron would.
      for (const listener of closeListeners) {
        listener({ preventDefault: vi.fn() })
      }

      const saved = loadWindowState(tempDir)
      expect(saved).toEqual({ x: 10, y: 20, width: 1200, height: 800, isMaximized: false })
    })

    it('reflects isMaximized:true when the window is maximized at close time', () => {
      mocks.fakeWindow.isMaximized.mockReturnValue(true)

      for (const listener of windowHandlers('close')) {
        listener({ preventDefault: vi.fn() })
      }

      expect(loadWindowState(tempDir)?.isMaximized).toBe(true)
    })

    it('registers resize/move listeners that schedule a persist (debounced, so no file yet)', () => {
      const resizeListeners = windowHandlers('resize')
      const moveListeners = windowHandlers('move')
      expect(resizeListeners.length).toBeGreaterThan(0)
      expect(moveListeners.length).toBeGreaterThan(0)

      for (const listener of resizeListeners) listener(undefined)

      // Debounced (WINDOW_STATE_SAVE_DEBOUNCE_MS) — nothing written yet.
      expect(fs.existsSync(path.join(tempDir, STATE_FILE_NAME))).toBe(false)
    })
  })

  // wave-3 shell-polish follow-up — a repeated close attempt while a
  // Save-before-close round trip is already pending (waiting on the
  // renderer's own save) must reuse that same round trip instead of showing
  // a second stacked confirmation dialog or registering a second
  // `ipcMain.once('save-before-close-result', ...)` listener.
  describe('close in-flight guard (wave-3 shell-polish follow-up)', () => {
    function markDirty() {
      handler('renderer:dirty-state')(ALLOWED_EVENT, true)
    }

    function triggerClose() {
      const event = { preventDefault: vi.fn() }
      // handleWindowCloseRequest is whichever 'close' listener isn't the
      // window-state persist callback — both run on a real close, but only
      // this one can call `event.preventDefault()`, so drive them all like
      // Electron would and let the assertions below observe the effect.
      for (const listener of windowHandlers('close')) {
        listener(event)
      }
      return event
    }

    it('shows exactly one dialog and registers exactly one save-result listener across two rapid close attempts', () => {
      markDirty()
      mocks.dialog.showMessageBoxSync.mockReturnValue(CLOSE_PROMPT_CHOICE.SAVE)

      const first = triggerClose()
      const second = triggerClose()

      expect(first.preventDefault).toHaveBeenCalledTimes(1)
      expect(second.preventDefault).toHaveBeenCalledTimes(1)
      expect(mocks.dialog.showMessageBoxSync).toHaveBeenCalledTimes(1)
      expect(mocks.onceHandlers.get('save-before-close-result')?.length ?? 0).toBe(1)
      expect(mocks.fakeWindow.destroy).not.toHaveBeenCalled()
    })

    it('destroys the window once the pending save reports success, and allows a fresh attempt afterward', () => {
      markDirty()
      mocks.dialog.showMessageBoxSync.mockReturnValue(CLOSE_PROMPT_CHOICE.SAVE)

      triggerClose()
      const [saveResultListener] = mocks.onceHandlers.get('save-before-close-result') ?? []
      expect(saveResultListener).toBeDefined()
      saveResultListener?.(ALLOWED_EVENT, { saved: true })

      expect(mocks.fakeWindow.destroy).toHaveBeenCalledTimes(1)

      // The guard is released — a later close attempt (a fresh window,
      // hypothetically) is free to show its own dialog again.
      mocks.dialog.showMessageBoxSync.mockClear()
      handler('renderer:dirty-state')(ALLOWED_EVENT, false)
      triggerClose()
      expect(mocks.dialog.showMessageBoxSync).not.toHaveBeenCalled()
    })

    it('a failed save leaves the guard released so the user can retry', () => {
      markDirty()
      mocks.dialog.showMessageBoxSync.mockReturnValue(CLOSE_PROMPT_CHOICE.SAVE)

      triggerClose()
      const [saveResultListener] = mocks.onceHandlers.get('save-before-close-result') ?? []
      saveResultListener?.(ALLOWED_EVENT, { saved: false })

      expect(mocks.fakeWindow.destroy).not.toHaveBeenCalled()

      // A second attempt now shows a fresh dialog rather than silently
      // no-oping forever.
      mocks.dialog.showMessageBoxSync.mockClear()
      triggerClose()
      expect(mocks.dialog.showMessageBoxSync).toHaveBeenCalledTimes(1)
    })

    it('Discard destroys the window immediately with no in-flight round trip', () => {
      markDirty()
      mocks.dialog.showMessageBoxSync.mockReturnValue(CLOSE_PROMPT_CHOICE.DISCARD)

      triggerClose()

      expect(mocks.fakeWindow.destroy).toHaveBeenCalledTimes(1)
      expect(mocks.onceHandlers.get('save-before-close-result')?.length ?? 0).toBe(0)
    })

    it('Cancel releases the guard immediately, allowing a fresh close attempt', () => {
      markDirty()
      mocks.dialog.showMessageBoxSync.mockReturnValue(CLOSE_PROMPT_CHOICE.CANCEL)

      triggerClose()
      expect(mocks.fakeWindow.destroy).not.toHaveBeenCalled()

      mocks.dialog.showMessageBoxSync.mockClear()
      triggerClose()
      expect(mocks.dialog.showMessageBoxSync).toHaveBeenCalledTimes(1)
    })

    it('a clean (non-dirty) document closes immediately with no dialog at all', () => {
      handler('renderer:dirty-state')(ALLOWED_EVENT, false)

      const event = triggerClose()

      expect(event.preventDefault).not.toHaveBeenCalled()
      expect(mocks.dialog.showMessageBoxSync).not.toHaveBeenCalled()
    })
  })
})

// wave-3 shell-polish follow-up — restoring saved window bounds happens at
// `createWindow()` time (the `BrowserWindow` constructor call itself), which
// the outer suite's shared `beforeEach` already requires main.cjs through
// before any test body runs. These need the state file seeded BEFORE that
// require, so they manage their own require lifecycle instead.
describe('electron/main.cjs — window state restore on create (P5.2/ELEC-11)', () => {
  let tempDir: string
  let mocks: ReturnType<typeof createElectronMock>
  let uncaughtBefore: number
  let unhandledBefore: number

  beforeEach(() => {
    uncaughtBefore = process.listenerCount('uncaughtException')
    unhandledBefore = process.listenerCount('unhandledRejection')
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-main-ipc-winstate-'))
    mocks = createElectronMock(tempDir)
  })

  afterEach(() => {
    delete nodeRequire.cache[ELECTRON_MODULE_PATH]
    delete nodeRequire.cache[MAIN_CJS_PATH]
    fs.rmSync(tempDir, { recursive: true, force: true })
    for (const listener of process.listeners('uncaughtException').slice(uncaughtBefore)) {
      process.removeListener('uncaughtException', listener)
    }
    for (const listener of process.listeners('unhandledRejection').slice(unhandledBefore)) {
      process.removeListener('unhandledRejection', listener)
    }
  })

  it('restores saved bounds and applies them to the BrowserWindow constructor', async () => {
    fs.writeFileSync(
      path.join(tempDir, STATE_FILE_NAME),
      JSON.stringify({ x: 50, y: 60, width: 1000, height: 700, isMaximized: false }),
      'utf-8',
    )

    await loadMainCjs(mocks)

    const options = mocks.browserWindowCalls[0]
    expect(options).toMatchObject({ x: 50, y: 60, width: 1000, height: 700 })
    expect(mocks.fakeWindow.maximize).not.toHaveBeenCalled()
  })

  it('maximizes the window when the saved state was maximized', async () => {
    fs.writeFileSync(
      path.join(tempDir, STATE_FILE_NAME),
      JSON.stringify({ x: 0, y: 0, width: 1920, height: 1080, isMaximized: true }),
      'utf-8',
    )

    await loadMainCjs(mocks)

    expect(mocks.fakeWindow.maximize).toHaveBeenCalledTimes(1)
  })

  it('falls back to the default size (no x/y) when the saved position is off every current display', async () => {
    fs.writeFileSync(
      path.join(tempDir, STATE_FILE_NAME),
      JSON.stringify({ x: 5000, y: 5000, width: 800, height: 600, isMaximized: false }),
      'utf-8',
    )
    // The mock's default `screen.getAllDisplays()` is a single 1920x1080
    // display at the origin — (5000, 5000) is off it entirely.

    await loadMainCjs(mocks)

    const options = mocks.browserWindowCalls[0]
    expect(options?.x).toBeUndefined()
    expect(options?.y).toBeUndefined()
    expect(options?.width).toBe(1200)
    expect(options?.height).toBe(800)
  })

  it('falls back to the default size on first launch (no saved state file at all)', async () => {
    await loadMainCjs(mocks)

    const options = mocks.browserWindowCalls[0]
    expect(options?.x).toBeUndefined()
    expect(options?.width).toBe(1200)
    expect(options?.height).toBe(800)
  })
})
