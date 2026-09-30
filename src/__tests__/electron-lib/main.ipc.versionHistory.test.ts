/**
 * VERSIONS-1 — the `history:*` IPC channels.
 *
 * The store itself is tested in `versionHistory.test.ts`; this covers the layer
 * that decides WHO may ask. A version history is a record of the user's own work,
 * so the interesting cases are the ones where a caller should be refused: a
 * request from a subframe, and a request naming a document the user never opened.
 * That second one matters because the renderer supplies the path — without the
 * allowlist check, a script running in a rendered document could read or start
 * writing history for any file on disk it cared to name.
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

describe('electron/main.cjs — history:* channels', () => {
  let tempDir: string
  let docPath: string
  let mocks: ReturnType<typeof createElectronMock>
  let uncaughtBefore: number
  let unhandledBefore: number

  beforeEach(async () => {
    uncaughtBefore = process.listenerCount('uncaughtException')
    unhandledBefore = process.listenerCount('unhandledRejection')
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-main-ipc-history-'))
    docPath = path.join(tempDir, 'thesis.md')
    fs.writeFileSync(docPath, 'on disk', 'utf8')
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

  /** Gets the document onto the allowlist the way a real open would. */
  async function openTheDocument(): Promise<void> {
    mocks.dialog.showOpenDialog.mockResolvedValueOnce({ canceled: false, filePaths: [docPath] })
    await handler('dialog:openFileBinary')(ALLOWED_EVENT)
  }

  const snapshot = (bytes: string, label?: string) =>
    handler('history:snapshot')(ALLOWED_EVENT, {
      path: docPath,
      bytes: new TextEncoder().encode(bytes),
      ...(label === undefined ? {} : { label }),
    })

  describe('security', () => {
    it('refuses every channel when the request is not from the main frame', async () => {
      await openTheDocument()
      const payload = { path: docPath, bytes: new Uint8Array([1]), id: 'a'.repeat(64) }

      for (const channel of ['history:snapshot', 'history:list', 'history:read', 'history:clear']) {
        const result = (await handler(channel)(REJECTED_EVENT, payload)) as { error?: string }
        expect(result.error, channel).toMatch(/could not be verified/i)
      }
    })

    it('refuses a document the user never opened', async () => {
      // The renderer names the path, so this is the check standing between a page
      // script and the history of any file it cares to name.
      const stranger = path.join(tempDir, 'not-opened.md')
      fs.writeFileSync(stranger, 'private', 'utf8')

      const stored = (await handler('history:snapshot')(ALLOWED_EVENT, {
        path: stranger,
        bytes: new Uint8Array([1, 2, 3]),
      })) as { stored: boolean }
      expect(stored.stored).toBe(false)

      const listed = (await handler('history:list')(ALLOWED_EVENT, stranger)) as { versions: unknown[] }
      expect(listed.versions).toEqual([])
    })

    it('refuses bytes that are not a byte array', async () => {
      await openTheDocument()
      for (const bytes of ['a string', 42, null, { length: 3 }, [1, 2, 3]]) {
        const result = (await handler('history:snapshot')(ALLOWED_EVENT, { path: docPath, bytes })) as {
          stored: boolean
        }
        expect(result.stored, String(bytes)).toBe(false)
      }
    })
  })

  describe('recording and reading back', () => {
    it('records a version and lists it', async () => {
      await openTheDocument()
      const result = (await snapshot('first draft')) as { stored: boolean }
      expect(result.stored).toBe(true)

      const listed = (await handler('history:list')(ALLOWED_EVENT, docPath)) as {
        versions: { id: string; bytes: number }[]
      }
      expect(listed.versions).toHaveLength(1)
      expect(listed.versions[0]!.bytes).toBe('first draft'.length)
    })

    it('reads a version back byte for byte', async () => {
      await openTheDocument()
      await snapshot('the exact content')
      const listed = (await handler('history:list')(ALLOWED_EVENT, docPath)) as { versions: { id: string }[] }

      const read = (await handler('history:read')(ALLOWED_EVENT, {
        path: docPath,
        id: listed.versions[0]!.id,
      })) as { bytes?: Uint8Array }

      expect(new TextDecoder().decode(read.bytes)).toBe('the exact content')
    })

    it('records what the renderer HAS, not what is on disk', async () => {
      // The whole point for an unsaved document: the interesting state is the one
      // in front of the user, which by definition is not yet on disk.
      await openTheDocument()
      await snapshot('unsaved edits')
      const listed = (await handler('history:list')(ALLOWED_EVENT, docPath)) as { versions: { id: string }[] }
      const read = (await handler('history:read')(ALLOWED_EVENT, {
        path: docPath,
        id: listed.versions[0]!.id,
      })) as { bytes?: Uint8Array }

      expect(new TextDecoder().decode(read.bytes)).toBe('unsaved edits')
      expect(fs.readFileSync(docPath, 'utf8')).toBe('on disk')
    })

    it('reports an unchanged save rather than pretending it made a version', async () => {
      await openTheDocument()
      await snapshot('same')
      const second = (await snapshot('same')) as { stored: boolean; reason?: string }

      expect(second.stored).toBe(false)
      expect(second.reason).toBe('unchanged')
    })

    it('reports a missing version instead of throwing', async () => {
      await openTheDocument()
      await snapshot('content')
      const read = (await handler('history:read')(ALLOWED_EVENT, { path: docPath, id: 'b'.repeat(64) })) as {
        error?: string
      }
      expect(read.error).toBeTruthy()
    })

    it('keeps a label when one is given', async () => {
      await openTheDocument()
      await snapshot('content', 'before restoring an earlier draft')
      const listed = (await handler('history:list')(ALLOWED_EVENT, docPath)) as {
        versions: { label: string | null }[]
      }
      expect(listed.versions[0]!.label).toBe('before restoring an earlier draft')
    })

    it('records a version automatically on every save, for any format', async () => {
      // The design that makes this feature work without each viewer having to
      // remember: main already holds the finished bytes of every format at the
      // moment it writes them, so history is captured there rather than asked for.
      await openTheDocument()

      // `new Uint8Array(...)`, built with the GLOBAL constructor: `save-binary-file`
      // checks `content instanceof Uint8Array`, and under jsdom that global is
      // jsdom's — which a Node `Buffer` is not an instance of. `main.cjs` runs
      // against the same worker globals, so constructing here the way the
      // renderer would is what makes the check behave as it does over real IPC.
      // Through the save dialog, which is what makes the path WRITE-eligible.
      // Opening a file only grants the read tier (SEC-1), so an `existingPath`
      // save of a merely-opened document correctly falls back to the dialog.
      mocks.dialog.showSaveDialog.mockResolvedValueOnce({ canceled: false, filePath: docPath })
      await handler('save-binary-file')(ALLOWED_EVENT, { content: new Uint8Array(Buffer.from('saved once')) })
      // The second save needs no dialog: the first one trusted the path.
      await handler('save-binary-file')(ALLOWED_EVENT, {
        content: new Uint8Array(Buffer.from('saved twice')),
        existingPath: docPath,
      })

      const listed = (await handler('history:list')(ALLOWED_EVENT, docPath)) as { versions: { id: string }[] }
      expect(listed.versions).toHaveLength(2)

      const first = (await handler('history:read')(ALLOWED_EVENT, {
        path: docPath,
        id: listed.versions[0]!.id,
      })) as { bytes?: Uint8Array }
      expect(new TextDecoder().decode(first.bytes)).toBe('saved once')
    })

    it('does not fail a save when recording the version fails', async () => {
      // A history is a convenience on top of saving. The file is already safely
      // on disk by the time a snapshot is attempted, so a failure there must
      // never be reported to the user as a failed save.
      await openTheDocument()
      const historyRoot = path.join(tempDir, 'history')
      fs.mkdirSync(path.dirname(historyRoot), { recursive: true })
      // A FILE where the history directory needs to be: every write beneath it
      // now fails.
      fs.writeFileSync(historyRoot, 'not a directory', 'utf8')

      mocks.dialog.showSaveDialog.mockResolvedValueOnce({ canceled: false, filePath: docPath })
      const saved = (await handler('save-binary-file')(ALLOWED_EVENT, {
        content: new Uint8Array(Buffer.from('content')),
      })) as { saved: boolean }

      expect(saved.saved).toBe(true)
      expect(fs.readFileSync(docPath, 'utf8')).toBe('content')
    })

    it('clears a document history on request', async () => {
      await openTheDocument()
      await snapshot('one')
      await snapshot('two')

      const cleared = (await handler('history:clear')(ALLOWED_EVENT, docPath)) as { cleared: boolean }
      expect(cleared.cleared).toBe(true)

      const listed = (await handler('history:list')(ALLOWED_EVENT, docPath)) as { versions: unknown[] }
      expect(listed.versions).toEqual([])
    })
  })
})
