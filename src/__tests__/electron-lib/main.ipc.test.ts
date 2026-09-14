/**
 * electron/main.cjs — IPC handler argument validation, sender-frame checks,
 * and allowlist gating (P2.13 IPC part / SHELL-27, QA-05).
 *
 * main.cjs is Atlas's Electron entry point and requires a real Electron
 * runtime to load normally (`require('electron')` returns a path string,
 * not `{app, BrowserWindow, ...}`, when required outside one). Vitest's
 * `vi.mock('electron', ...)` does not intercept this: `.cjs` files' own
 * `require(...)` calls resolve through Node's real module loader rather
 * than vite-node's ESM-aware mock registry (verified — a `vi.mock('electron')`
 * left `require('electron')` inside a `.cjs` sibling file completely
 * unmocked). Instead, this suite injects a fake `electron` module directly
 * into Node's own `require.cache` (via `createRequire`) before requiring a
 * fresh copy of `main.cjs`, so real Node `require()` resolves `'electron'`
 * to the fake exports without ever touching the installed npm package.
 * `main.cjs`'s own `require('./lib/*.cjs')` calls are left real — those are
 * genuine local modules with no Electron dependency, already covered by
 * their own dedicated tests under this same directory.
 */
import { createRequire } from 'node:module'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

type IpcHandler = (event: unknown, ...args: unknown[]) => unknown

const nodeRequire = createRequire(import.meta.url)
const ELECTRON_MODULE_PATH = nodeRequire.resolve('electron')
const MAIN_CJS_PATH = nodeRequire.resolve('../../../electron/main.cjs')

const FAKE_MAIN_FRAME = { id: 'main-frame' }
const OTHER_FRAME = { id: 'other-frame' }
const ALLOWED_EVENT = { senderFrame: FAKE_MAIN_FRAME }
const REJECTED_EVENT = { senderFrame: OTHER_FRAME }

function createElectronMock(tempDir: string) {
  const ipcHandlers = new Map<string, IpcHandler>()

  const fakeWindow = {
    isDestroyed: () => false,
    isMinimized: () => false,
    isVisible: () => true,
    focus: vi.fn(),
    restore: vi.fn(),
    show: vi.fn(),
    reload: vi.fn(),
    setTitleBarOverlay: vi.fn(),
    loadURL: vi.fn(),
    loadFile: vi.fn(),
    on: vi.fn(),
    // Invoke the 'ready-to-show' callback immediately (synchronously) so
    // createWindow()'s real code path — including its own
    // `clearTimeout(showTimeout)` — runs to completion without leaving a
    // real 5s timer dangling past the end of the test.
    once: vi.fn((event: string, cb: () => void) => {
      if (event === 'ready-to-show') cb()
    }),
    webContents: {
      mainFrame: FAKE_MAIN_FRAME,
      session: {
        webRequest: { onHeadersReceived: vi.fn() },
        availableSpellCheckerLanguages: ['en-US'],
        getSpellCheckerLanguages: vi.fn(() => ['en-US']),
        setSpellCheckerLanguages: vi.fn(),
        addWordToSpellCheckerDictionary: vi.fn(),
      },
      on: vi.fn(),
      once: vi.fn(),
      setWindowOpenHandler: vi.fn(),
      getURL: vi.fn(() => 'http://localhost:5173'),
      openDevTools: vi.fn(),
      replaceMisspelling: vi.fn(),
    },
  }

  const dialog = {
    showOpenDialog: vi.fn().mockResolvedValue({ canceled: true, filePaths: [] }),
    showSaveDialog: vi.fn().mockResolvedValue({ canceled: true }),
    showMessageBoxSync: vi.fn(),
    showErrorBox: vi.fn(),
  }

  const shell = {
    openExternal: vi.fn(),
    showItemInFolder: vi.fn(),
  }

  const Menu = {
    buildFromTemplate: vi.fn(() => ({})),
    setApplicationMenu: vi.fn(),
  }

  function BrowserWindowCtor() {
    return fakeWindow
  }
  BrowserWindowCtor.getAllWindows = vi.fn(() => [])

  const app = {
    requestSingleInstanceLock: vi.fn(() => true),
    quit: vi.fn(),
    isPackaged: false,
    isReady: vi.fn(() => true),
    getPath: vi.fn(() => tempDir),
    getLocale: vi.fn(() => 'en-US'),
    getAppPath: vi.fn(() => process.cwd()),
    whenReady: vi.fn(() => Promise.resolve()),
    on: vi.fn(),
  }

  const ipcMain = {
    handle: vi.fn((channel: string, fn: IpcHandler) => {
      ipcHandlers.set(channel, fn)
    }),
    on: vi.fn((channel: string, fn: IpcHandler) => {
      ipcHandlers.set(channel, fn)
    }),
  }

  return { app, BrowserWindow: BrowserWindowCtor, ipcMain, dialog, shell, Menu, ipcHandlers, fakeWindow }
}

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
  let originalArgv: string[]

  beforeEach(async () => {
    uncaughtBefore = process.listenerCount('uncaughtException')
    unhandledBefore = process.listenerCount('unhandledRejection')
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-main-ipc-'))
    mocks = createElectronMock(tempDir)

    // main.cjs's `extractFilePath(process.argv)` runs at module scope to
    // support the real "OS Open With" launch path — but vitest's own argv
    // (a real, existing .mjs script path) can spuriously match its known-
    // extensions heuristic. Stub argv to an innocuous shape (matching a
    // plain "no file argument" launch) for the duration of the require.
    originalArgv = process.argv;
    process.argv = ['node', 'main.cjs'];

    // Inject the fake `electron` module into Node's own require cache so
    // main.cjs's `require('electron')` resolves to it (see file header).
    nodeRequire.cache[ELECTRON_MODULE_PATH] = {
      id: ELECTRON_MODULE_PATH,
      filename: ELECTRON_MODULE_PATH,
      loaded: true,
      exports: mocks,
      children: [],
      paths: [],
    } as unknown as NodeJS.Module

    delete nodeRequire.cache[MAIN_CJS_PATH]
    nodeRequire(MAIN_CJS_PATH)
    process.argv = originalArgv;

    // Flush the `app.whenReady().then(() => { createWindow(); ... })`
    // microtask chain so `mainWindow` (private to main.cjs) is populated
    // before any handler that checks it runs.
    await Promise.resolve()
    await Promise.resolve()
    await Promise.resolve()
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
})
