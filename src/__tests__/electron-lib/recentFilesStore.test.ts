import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { STORE_FILE_NAME, createRecentFilesStore } from '../../../electron/lib/recentFilesStore.cjs'

let storeDir: string

beforeEach(() => {
  storeDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-recent-files-'))
})

afterEach(() => {
  fs.rmSync(storeDir, { recursive: true, force: true })
})

describe('createRecentFilesStore', () => {
  it('reports a path as known only after it has been recorded', () => {
    const store = createRecentFilesStore(storeDir)
    expect(store.has('C:\\a\\b.docx')).toBe(false)
    store.record('C:\\a\\b.docx')
    expect(store.has('C:\\a\\b.docx')).toBe(true)
  })

  it('does not treat an arbitrary existing path as known just because it was never recorded', () => {
    // This is the exact gap the store closes: `recent:request-open` must not
    // accept a path purely because it exists on disk.
    const target = path.join(storeDir, 'never-opened.txt')
    fs.writeFileSync(target, 'hello')
    const store = createRecentFilesStore(storeDir)
    expect(store.has(target)).toBe(false)
  })

  it('normalizes case and slash direction the same way as the path allowlist', () => {
    const store = createRecentFilesStore(storeDir)
    store.record('C:\\Users\\Test\\Doc.DOCX')
    expect(store.has('c:/users/test/doc.docx')).toBe(true)
  })

  it('persists recorded paths across store instances (survives an app restart)', () => {
    const first = createRecentFilesStore(storeDir)
    first.record('C:\\a\\b.docx')

    const second = createRecentFilesStore(storeDir)
    expect(second.has('C:\\a\\b.docx')).toBe(true)
  })

  it('writes its backing file under the given directory', () => {
    const store = createRecentFilesStore(storeDir)
    store.record('C:\\a\\b.docx')
    expect(fs.existsSync(path.join(storeDir, STORE_FILE_NAME))).toBe(true)
  })

  it('caps stored entries and evicts the oldest first', () => {
    const store = createRecentFilesStore(storeDir)
    for (let i = 0; i < 55; i += 1) {
      store.record(`C:\\docs\\file-${i}.txt`)
    }
    expect(store.has('C:\\docs\\file-0.txt')).toBe(false)
    expect(store.has('C:\\docs\\file-54.txt')).toBe(true)
  })

  it('ignores unresolvable input without throwing', () => {
    const store = createRecentFilesStore(storeDir)
    expect(() => store.record(undefined)).not.toThrow()
    expect(store.has(undefined)).toBe(false)
  })

  it('never throws even when the store directory cannot be written to', () => {
    const store = createRecentFilesStore(path.join(storeDir, 'nested', 'deep', 'dir'))
    expect(() => store.record('C:\\a\\b.docx')).not.toThrow()
  })
})
