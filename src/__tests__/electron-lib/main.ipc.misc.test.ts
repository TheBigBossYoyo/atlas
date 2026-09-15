/**
 * electron/main.cjs — the remaining IPC channels not already covered by
 * `main.ipc.test.ts` (QA-05): `get-initial-file`, `open-file-dialog`,
 * `spellcheck:*`, `set-theme`, and `renderer:dirty-state`'s own basic
 * wiring (its effect on the close guard is covered in depth by
 * `main.ipc.test.ts`'s "close in-flight guard" suite — this only asserts
 * the handler itself validates its argument shape).
 *
 * Shares its fake-`electron` harness with `main.ipc.test.ts` — see
 * `mainIpcTestHarness.ts`'s header for why a real Electron runtime can't be
 * mocked the ordinary `vi.mock('electron', ...)` way here.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

import { afterEach, beforeEach, describe, expect, it } from 'vitest'

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

describe('electron/main.cjs — remaining IPC handlers', () => {
  let tempDir: string
  let mocks: ReturnType<typeof createElectronMock>
  let uncaughtBefore: number
  let unhandledBefore: number

  beforeEach(async () => {
    uncaughtBefore = process.listenerCount('uncaughtException')
    unhandledBefore = process.listenerCount('unhandledRejection')
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-main-ipc-misc-'))
    mocks = createElectronMock(tempDir)
    await loadMainCjs(mocks)
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

  function handler(channel: string): IpcHandler {
    const fn = mocks.ipcHandlers.get(channel)
    if (!fn) throw new Error(`no handler registered for "${channel}"`)
    return fn
  }

  describe('get-initial-file', () => {
    it('returns null when no file was queued at launch', () => {
      const result = handler('get-initial-file')(ALLOWED_EVENT)
      expect(result).toBeNull()
    })
  })

  describe('open-file-dialog', () => {
    it('returns null for a request not from the main frame, without opening a dialog', async () => {
      const result = await handler('open-file-dialog')(REJECTED_EVENT)
      expect(result).toBeNull()
      expect(mocks.dialog.showOpenDialog).not.toHaveBeenCalled()
    })

    it('returns null when the dialog is dismissed', async () => {
      mocks.dialog.showOpenDialog.mockResolvedValueOnce({ canceled: true, filePaths: [] })
      const result = await handler('open-file-dialog')(ALLOWED_EVENT)
      expect(result).toBeNull()
    })

    it('reads and allowlists the chosen markdown file, restricted to markdown/text filters', async () => {
      const filePath = path.join(tempDir, 'notes.md')
      fs.writeFileSync(filePath, '# hello')
      mocks.dialog.showOpenDialog.mockResolvedValueOnce({ canceled: false, filePaths: [filePath] })

      const result = (await handler('open-file-dialog')(ALLOWED_EVENT)) as { content: string; path: string } | null

      expect(result?.path).toBe(filePath)
      expect(result?.content).toContain('hello')

      const filterArg = mocks.dialog.showOpenDialog.mock.calls[0]?.[1] as {
        filters: Array<{ extensions: string[] }>
      }
      expect(filterArg.filters[0]?.extensions).toEqual(['md', 'markdown', 'txt'])

      // The chosen path is now allowlisted via the same trustPath() flow as
      // every other interactive open — file:readBinaryByPath should succeed.
      const readResult = (await handler('file:readBinaryByPath')(ALLOWED_EVENT, filePath)) as {
        buffer: ArrayBuffer
      }
      expect(Buffer.from(readResult.buffer).toString()).toContain('hello')
    })
  })

  describe('spellcheck:add-word', () => {
    it('rejects an empty word without touching the dictionary', () => {
      const result = handler('spellcheck:add-word')({}, '') as { added: boolean }
      expect(result.added).toBe(false)
      expect(mocks.fakeWindow.webContents.session.addWordToSpellCheckerDictionary).not.toHaveBeenCalled()
    })

    it('rejects a non-string word', () => {
      const result = handler('spellcheck:add-word')({}, 42) as { added: boolean }
      expect(result.added).toBe(false)
    })

    it('adds a real word to the dictionary', () => {
      const result = handler('spellcheck:add-word')({}, 'atlas') as { added: boolean }
      expect(result.added).toBe(true)
      expect(mocks.fakeWindow.webContents.session.addWordToSpellCheckerDictionary).toHaveBeenCalledWith('atlas')
    })
  })

  describe('spellcheck:replace-misspelling', () => {
    it('rejects an empty word', () => {
      const result = handler('spellcheck:replace-misspelling')({}, '') as { replaced: boolean }
      expect(result.replaced).toBe(false)
      expect(mocks.fakeWindow.webContents.replaceMisspelling).not.toHaveBeenCalled()
    })

    it('replaces a real correction', () => {
      const result = handler('spellcheck:replace-misspelling')({}, 'atlas') as { replaced: boolean }
      expect(result.replaced).toBe(true)
      expect(mocks.fakeWindow.webContents.replaceMisspelling).toHaveBeenCalledWith('atlas')
    })
  })

  describe('spellcheck:get-languages', () => {
    it('returns the mocked available/enabled language lists', () => {
      const result = handler('spellcheck:get-languages')({}) as { available: string[]; enabled: string[] }
      expect(result.available).toEqual(['en-US'])
      expect(result.enabled).toEqual(['en-US'])
    })
  })

  describe('spellcheck:set-languages', () => {
    it('rejects a non-array payload without calling into the session', () => {
      // createWindow() itself already called setSpellCheckerLanguages once,
      // during its own default-locale spellcheck setup — clear that call so
      // this only asserts what the HANDLER does with a bad payload.
      mocks.fakeWindow.webContents.session.setSpellCheckerLanguages.mockClear()

      const result = handler('spellcheck:set-languages')({}, 'en-US') as { ok: boolean }
      expect(result.ok).toBe(false)
      expect(mocks.fakeWindow.webContents.session.setSpellCheckerLanguages).not.toHaveBeenCalled()
    })

    it('filters the requested languages down to ones the session actually supports', () => {
      const result = handler('spellcheck:set-languages')({}, ['en-US', 'zz-ZZ']) as {
        ok: boolean
        enabled: string[]
      }
      expect(result.ok).toBe(true)
      expect(result.enabled).toEqual(['en-US'])
      expect(mocks.fakeWindow.webContents.session.setSpellCheckerLanguages).toHaveBeenLastCalledWith(['en-US'])
    })
  })

  describe('set-theme', () => {
    it('applies a known theme\'s overlay colors to the title bar', () => {
      handler('set-theme')({}, 'dark')
      expect(mocks.fakeWindow.setTitleBarOverlay).toHaveBeenCalledWith(
        expect.objectContaining({ color: '#161b22', symbolColor: '#e6edf3' }),
      )
    })

    it('falls back to the light overlay colors for an unrecognized theme name', () => {
      handler('set-theme')({}, 'not-a-real-theme')
      expect(mocks.fakeWindow.setTitleBarOverlay).toHaveBeenCalledWith(
        expect.objectContaining({ color: '#f6f8fa', symbolColor: '#1f2328' }),
      )
    })
  })

  describe('renderer:dirty-state', () => {
    // main.cjs registers more than one 'close' listener (persistWindowState
    // AND handleWindowCloseRequest) — invoke every one of them, exactly as a
    // real close attempt would, rather than assuming a registration order.
    function fireCloseAttempt(): { defaultPrevented: boolean } {
      let defaultPrevented = false
      for (const listener of mocks.windowOnHandlers.get('close') ?? []) {
        listener({ preventDefault: () => { defaultPrevented = true } })
      }
      return { defaultPrevented }
    }

    it('ignores a signal not from the main frame — closing while dirty still prompts', () => {
      handler('renderer:dirty-state')(REJECTED_EVENT, true)

      const { defaultPrevented } = fireCloseAttempt()

      // rendererDirty is still false (the untrusted signal was ignored), so
      // the close proceeds without ever prompting.
      expect(defaultPrevented).toBe(false)
      expect(mocks.dialog.showMessageBoxSync).not.toHaveBeenCalled()
    })

    it('a trusted dirty signal blocks the next close attempt behind the Save/Discard/Cancel prompt', () => {
      handler('renderer:dirty-state')(ALLOWED_EVENT, true)

      const { defaultPrevented } = fireCloseAttempt()

      expect(defaultPrevented).toBe(true)
      expect(mocks.dialog.showMessageBoxSync).toHaveBeenCalledTimes(1)
    })
  })
})
