import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { FileLockedError, atomicWriteFile } from '../../../electron/lib/atomicWrite.cjs'

let tempDir: string

beforeEach(() => {
  tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-atomic-write-'))
})

afterEach(() => {
  vi.restoreAllMocks()
  fs.rmSync(tempDir, { recursive: true, force: true })
})

describe('atomicWriteFile', () => {
  it('writes string content to a new file', () => {
    const target = path.join(tempDir, 'doc.md')

    const result = atomicWriteFile(target, '# hello')

    expect(result.fallbackUsed).toBe(false)
    expect(fs.readFileSync(target, 'utf-8')).toBe('# hello')
  })

  it('writes binary content to a new file', () => {
    const target = path.join(tempDir, 'doc.docx')
    const bytes = new Uint8Array([1, 2, 3, 4])

    atomicWriteFile(target, bytes)

    expect(new Uint8Array(fs.readFileSync(target))).toEqual(bytes)
  })

  it('leaves no leftover temp file after a successful write', () => {
    const target = path.join(tempDir, 'doc.md')

    atomicWriteFile(target, 'content')

    const entries = fs.readdirSync(tempDir)
    expect(entries).toEqual(['doc.md'])
  })

  it('overwrites the target and produces a rolling .bak of the prior version', () => {
    const target = path.join(tempDir, 'doc.md')
    fs.writeFileSync(target, 'version 1')

    atomicWriteFile(target, 'version 2')

    expect(fs.readFileSync(target, 'utf-8')).toBe('version 2')
    expect(fs.readFileSync(`${target}.bak`, 'utf-8')).toBe('version 1')
  })

  it('does not modify the original file if the temp write fails', () => {
    const target = path.join(tempDir, 'doc.md')
    fs.writeFileSync(target, 'original')

    vi.spyOn(fs, 'openSync').mockImplementationOnce(() => {
      throw new Error('simulated crash before temp write')
    })

    expect(() => atomicWriteFile(target, 'new content')).toThrow()
    expect(fs.readFileSync(target, 'utf-8')).toBe('original')
    expect(fs.existsSync(`${target}.bak`)).toBe(false)
  })

  it('falls back to a direct write when rename fails with EXDEV', () => {
    const target = path.join(tempDir, 'doc.md')
    fs.writeFileSync(target, 'original')

    vi.spyOn(fs, 'renameSync').mockImplementationOnce(() => {
      const err = Object.assign(new Error('cross-device link'), { code: 'EXDEV' })
      throw err
    })

    const result = atomicWriteFile(target, 'new content')

    expect(result.fallbackUsed).toBe(true)
    expect(fs.readFileSync(target, 'utf-8')).toBe('new content')
    expect(fs.readFileSync(`${target}.bak`, 'utf-8')).toBe('original')
  })

  it('surfaces a friendly FileLockedError when the target is locked (EBUSY)', () => {
    const target = path.join(tempDir, 'doc.docx')
    fs.writeFileSync(target, 'original')

    vi.spyOn(fs, 'renameSync').mockImplementationOnce(() => {
      const err = Object.assign(new Error('resource busy or locked'), { code: 'EBUSY' })
      throw err
    })

    expect(() => atomicWriteFile(target, 'new content')).toThrow(FileLockedError)
    // The original file must remain untouched when the rename itself fails.
    expect(fs.readFileSync(target, 'utf-8')).toBe('original')
  })

  it('surfaces a friendly FileLockedError when the target is locked (EPERM)', () => {
    const target = path.join(tempDir, 'doc.docx')

    vi.spyOn(fs, 'openSync').mockImplementationOnce(() => {
      const err = Object.assign(new Error('operation not permitted'), { code: 'EPERM' })
      throw err
    })

    expect(() => atomicWriteFile(target, 'new content')).toThrow(FileLockedError)
  })

  it('cleans up the temp file when the EXDEV fallback write itself fails', () => {
    const target = path.join(tempDir, 'doc.md')
    fs.writeFileSync(target, 'original')

    vi.spyOn(fs, 'renameSync').mockImplementationOnce(() => {
      const err = Object.assign(new Error('cross-device link'), { code: 'EXDEV' })
      throw err
    })

    // First openSync call is for the temp file (must succeed); force the
    // second call (the EXDEV fallback's direct write to targetPath) to fail
    // with a lock error.
    let callCount = 0
    const realOpenSync = fs.openSync.bind(fs)
    vi.spyOn(fs, 'openSync').mockImplementation((...args: Parameters<typeof fs.openSync>) => {
      callCount += 1
      if (callCount === 2) {
        const err = Object.assign(new Error('resource busy or locked'), { code: 'EBUSY' })
        throw err
      }
      return realOpenSync(...args)
    })

    expect(() => atomicWriteFile(target, 'new content')).toThrow(FileLockedError)
    const entries = fs.readdirSync(tempDir).filter((entry) => entry.includes('.tmp-'))
    expect(entries).toEqual([])
  })
})
