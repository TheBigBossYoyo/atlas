/**
 * Shared fake-`electron`-module harness for exercising `electron/main.cjs`
 * from Vitest (QA-05). Extracted out of `main.ipc.test.ts` (wave-3
 * shell-polish follow-up) so `main.ipc.misc.test.ts` can cover the
 * remaining IPC channels (get-initial-file, open-file-dialog, spellcheck:*,
 * set-theme, renderer:dirty-state) without duplicating this setup.
 *
 * main.cjs is Atlas's Electron entry point and requires a real Electron
 * runtime to load normally (`require('electron')` returns a path string,
 * not `{app, BrowserWindow, ...}`, when required outside one). Vitest's
 * `vi.mock('electron', ...)` does not intercept this: `.cjs` files' own
 * `require(...)` calls resolve through Node's real module loader rather
 * than vite-node's ESM-aware mock registry (verified — a `vi.mock('electron')`
 * left `require('electron')` inside a `.cjs` sibling file completely
 * unmocked). Instead, this harness injects a fake `electron` module directly
 * into Node's own `require.cache` (via `createRequire`) before requiring a
 * fresh copy of `main.cjs`, so real Node `require()` resolves `'electron'`
 * to the fake exports without ever touching the installed npm package.
 * `main.cjs`'s own `require('./lib/*.cjs')` calls are left real — those are
 * genuine local modules with no Electron dependency, already covered by
 * their own dedicated tests under this same directory.
 */
import { createRequire } from 'node:module'

import { vi } from 'vitest'

export type IpcHandler = (event: unknown, ...args: unknown[]) => unknown

export const nodeRequire = createRequire(import.meta.url)
export const ELECTRON_MODULE_PATH = nodeRequire.resolve('electron')
export const MAIN_CJS_PATH = nodeRequire.resolve('../../../electron/main.cjs')

export const FAKE_MAIN_FRAME = { id: 'main-frame' }
export const OTHER_FRAME = { id: 'other-frame' }
export const ALLOWED_EVENT = { senderFrame: FAKE_MAIN_FRAME }
export const REJECTED_EVENT = { senderFrame: OTHER_FRAME }

export function createElectronMock(tempDir: string) {
  const ipcHandlers = new Map<string, IpcHandler>()
  const onceHandlers = new Map<string, IpcHandler[]>()
  // Every `mainWindow.on(event, fn)` registration, in registration order —
  // main.cjs registers more than one listener for the same event (e.g. both
  // `persistWindowState` and `handleWindowCloseRequest` on 'close'), so
  // tests need every listener for an event, not just the last.
  const windowOnHandlers = new Map<string, IpcHandler[]>()

  const fakeWindow = {
    isDestroyed: () => false,
    isMinimized: () => false,
    isMaximized: vi.fn(() => false),
    isVisible: () => true,
    focus: vi.fn(),
    restore: vi.fn(),
    maximize: vi.fn(),
    show: vi.fn(),
    reload: vi.fn(),
    destroy: vi.fn(),
    setTitleBarOverlay: vi.fn(),
    getNormalBounds: vi.fn(() => ({ x: 10, y: 20, width: 1200, height: 800 })),
    loadURL: vi.fn(),
    loadFile: vi.fn(),
    on: vi.fn((event: string, fn: IpcHandler) => {
      const list = windowOnHandlers.get(event) ?? []
      list.push(fn)
      windowOnHandlers.set(event, list)
    }),
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
      send: vi.fn(),
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

  const browserWindowCalls: Array<Record<string, unknown>> = []
  function BrowserWindowCtor(options?: Record<string, unknown>) {
    browserWindowCalls.push(options ?? {})
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
    // Mirrors real EventEmitter.once semantics closely enough for these
    // tests: every registration for a channel is kept (not just the last),
    // so a test can assert exactly how many `once` listeners a given flow
    // registered — the close in-flight guard's whole point is keeping this
    // at 1 across repeated close attempts instead of accumulating.
    once: vi.fn((channel: string, fn: IpcHandler) => {
      const list = onceHandlers.get(channel) ?? []
      list.push(fn)
      onceHandlers.set(channel, list)
    }),
  }

  const screen = {
    getAllDisplays: vi.fn(() => [{ workArea: { x: 0, y: 0, width: 1920, height: 1080 } }]),
  }

  return {
    app,
    BrowserWindow: BrowserWindowCtor,
    ipcMain,
    dialog,
    shell,
    Menu,
    screen,
    ipcHandlers,
    onceHandlers,
    windowOnHandlers,
    browserWindowCalls,
    fakeWindow,
  }
}

/**
 * Injects the fake `electron` module into Node's own require cache and
 * requires a fresh `main.cjs` against it (see file header for why), then
 * flushes the `app.whenReady().then(() => { createWindow(); ... })`
 * microtask chain so `mainWindow` is populated before the caller touches
 * anything. Stubs `process.argv` for the duration of the require only — see
 * the inline comment below for why.
 *
 * `extraModules` (module path -> fake exports) injects additional local
 * `.cjs` sibling modules the same way — e.g. `main.ipc.test.ts` uses this to
 * fake `electron/lib/printToPdf.cjs` so the `export:printToPdf` handler's
 * own validation/error-mapping gets real unit coverage independent of the
 * rendering pipeline underneath it, without every other suite sharing this
 * harness needing to know that module exists.
 */
export async function loadMainCjs(
  mocks: ReturnType<typeof createElectronMock>,
  extraModules?: Readonly<Record<string, unknown>>,
): Promise<void> {
  // main.cjs's `extractFilePath(process.argv)` runs at module scope to
  // support the real "OS Open With" launch path — but vitest's own argv
  // (a real, existing .mjs script path) can spuriously match its known-
  // extensions heuristic. Stub argv to an innocuous shape (matching a
  // plain "no file argument" launch) for the duration of the require.
  const originalArgv = process.argv
  process.argv = ['node', 'main.cjs']

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

  for (const [modulePath, exports] of Object.entries(extraModules ?? {})) {
    nodeRequire.cache[modulePath] = {
      id: modulePath,
      filename: modulePath,
      loaded: true,
      exports,
      children: [],
      paths: [],
    } as unknown as NodeJS.Module
  }

  delete nodeRequire.cache[MAIN_CJS_PATH]
  nodeRequire(MAIN_CJS_PATH)
  process.argv = originalArgv

  await Promise.resolve()
  await Promise.resolve()
  await Promise.resolve()
}
