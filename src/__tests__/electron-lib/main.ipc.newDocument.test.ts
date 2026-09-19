/**
 * NEW-01 — electron/main.cjs's `document:new` handler (the toolbar's "New"
 * action / Ctrl+N) and the "open a 0-byte file of an editable format"
 * substitution (`substituteBlankTemplateIfEmpty`, applied inside
 * `dialog:openFileBinary` and `file:readBinaryByPath`).
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

const TEMPLATES_DIR = path.resolve(__dirname, '../../../electron/templates')

describe('electron/main.cjs — document:new + empty-file substitution', () => {
  let tempDir: string
  let mocks: ReturnType<typeof createElectronMock>
  let uncaughtBefore: number
  let unhandledBefore: number

  beforeEach(async () => {
    uncaughtBefore = process.listenerCount('uncaughtException')
    unhandledBefore = process.listenerCount('unhandledRejection')
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-main-ipc-newdoc-'))
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

  describe('document:new', () => {
    it('rejects a request not from the main frame', async () => {
      const result = (await handler('document:new')(REJECTED_EVENT, 'docx')) as { created: boolean; error?: string }
      expect(result.created).toBe(false)
      expect(result.error).toMatch(/could not be verified/i)
    })

    it('rejects a format outside the fixed list — never a renderer-supplied path or bytes', async () => {
      const result = (await handler('document:new')(ALLOWED_EVENT, '../../etc/passwd')) as {
        created: boolean
        error?: string
      }
      expect(result.created).toBe(false)
      expect(result.error).toMatch(/unsupported/i)
      expect(mocks.dialog.showSaveDialog).not.toHaveBeenCalled()
    })

    it('rejects inherited property names as formats', async () => {
      for (const name of ['constructor', 'toString', '__proto__']) {
        const result = (await handler('document:new')(ALLOWED_EVENT, name)) as { created: boolean; error?: string }
        expect(result.created).toBe(false)
        expect(result.error).toMatch(/unsupported/i)
      }
      expect(mocks.dialog.showSaveDialog).not.toHaveBeenCalled()
    })

    it('reports not-created when the save dialog is cancelled', async () => {
      mocks.dialog.showSaveDialog.mockResolvedValueOnce({ canceled: true })
      const result = (await handler('document:new')(ALLOWED_EVENT, 'docx')) as { created: boolean }
      expect(result.created).toBe(false)
    })

    it('writes a real (non-empty) blank docx at the chosen path and allowlists it', async () => {
      const filePath = path.join(tempDir, 'Document.docx')
      mocks.dialog.showSaveDialog.mockResolvedValueOnce({ canceled: false, filePath })

      const result = (await handler('document:new')(ALLOWED_EVENT, 'docx')) as { created: boolean; path?: string }
      expect(result.created).toBe(true)
      expect(result.path).toBe(filePath)

      const bytes = fs.readFileSync(filePath)
      expect(bytes.byteLength).toBeGreaterThan(0)
      // The exact bytes committed under electron/templates/ — proves this is
      // a real copy of the generated template, not some other content.
      expect(bytes.equals(fs.readFileSync(path.join(TEMPLATES_DIR, 'blank.docx')))).toBe(true)

      // Freshly created — the next write-back should go straight through
      // without a second save dialog (the allowlist + trusted-history flow
      // every other write path already relies on).
      const readResult = (await handler('file:readBinaryByPath')(ALLOWED_EVENT, filePath)) as { buffer: ArrayBuffer }
      expect(readResult.buffer.byteLength).toBeGreaterThan(0)
    })

    it('writes an empty markdown file (no template — blank markdown IS the blank document)', async () => {
      const filePath = path.join(tempDir, 'Document.md')
      mocks.dialog.showSaveDialog.mockResolvedValueOnce({ canceled: false, filePath })

      const result = (await handler('document:new')(ALLOWED_EVENT, 'markdown')) as { created: boolean }
      expect(result.created).toBe(true)
      expect(fs.readFileSync(filePath, 'utf-8')).toBe('')
    })

    it.each(['xlsx', 'ods', 'pptx', 'odp'] as const)('writes a real blank %s at the chosen path', async (format) => {
      const filePath = path.join(tempDir, `Document.${format}`)
      mocks.dialog.showSaveDialog.mockResolvedValueOnce({ canceled: false, filePath })

      const result = (await handler('document:new')(ALLOWED_EVENT, format)) as { created: boolean }
      expect(result.created).toBe(true)
      expect(fs.readFileSync(filePath).byteLength).toBeGreaterThan(0)
    })
  })

  describe('opening a 0-byte file of an editable format', () => {
    it('dialog:openFileBinary substitutes the blank template for a 0-byte docx', async () => {
      const filePath = path.join(tempDir, 'empty.docx')
      fs.writeFileSync(filePath, Buffer.alloc(0))
      mocks.dialog.showOpenDialog.mockResolvedValueOnce({ canceled: false, filePaths: [filePath] })

      const result = (await handler('dialog:openFileBinary')(ALLOWED_EVENT)) as { buffer: ArrayBuffer }
      expect(result.buffer.byteLength).toBeGreaterThan(0)
      expect(Buffer.from(result.buffer).equals(fs.readFileSync(path.join(TEMPLATES_DIR, 'blank.docx')))).toBe(true)

      // The file ON DISK is still genuinely empty — nothing was written yet;
      // only a normal Save (through the ordinary save-binary-file path)
      // should ever fill it in.
      expect(fs.statSync(filePath).size).toBe(0)
    })

    it('file:readBinaryByPath substitutes the blank template for a 0-byte xlsx', async () => {
      const filePath = path.join(tempDir, 'empty.xlsx')
      fs.writeFileSync(filePath, Buffer.alloc(0))
      mocks.dialog.showOpenDialog.mockResolvedValueOnce({ canceled: false, filePaths: [filePath] })
      await handler('dialog:openFileBinary')(ALLOWED_EVENT) // allowlists filePath (a different 0-byte fixture below reuses the same flow)

      const result = (await handler('file:readBinaryByPath')(ALLOWED_EVENT, filePath)) as { buffer: ArrayBuffer }
      expect(result.buffer.byteLength).toBeGreaterThan(0)
    })

    it('leaves a genuinely 0-byte non-creatable format (pdf) empty — no substitution', async () => {
      const filePath = path.join(tempDir, 'empty.pdf')
      fs.writeFileSync(filePath, Buffer.alloc(0))
      mocks.dialog.showOpenDialog.mockResolvedValueOnce({ canceled: false, filePaths: [filePath] })

      const result = (await handler('dialog:openFileBinary')(ALLOWED_EVENT)) as { buffer: ArrayBuffer }
      expect(result.buffer.byteLength).toBe(0)
    })

    it('leaves a non-empty docx untouched', async () => {
      const filePath = path.join(tempDir, 'real.docx')
      fs.writeFileSync(filePath, 'not actually a zip, but non-empty')
      mocks.dialog.showOpenDialog.mockResolvedValueOnce({ canceled: false, filePaths: [filePath] })

      const result = (await handler('dialog:openFileBinary')(ALLOWED_EVENT)) as { buffer: ArrayBuffer }
      expect(Buffer.from(result.buffer).toString()).toBe('not actually a zip, but non-empty')
    })
  })
})
