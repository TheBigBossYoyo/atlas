import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { FileLockedError, atomicWriteFile, classifyWriteError } from '../../../electron/lib/atomicWrite.cjs'

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

  it('classifies a rename onto an existing directory as EISDIR, not a file lock (X5 fix — Windows raises EPERM for this, same code isLockError treats as "locked")', () => {
    // Empirically verified on Windows: `fs.renameSync(tempFile, existingDir)`
    // throws with `code: 'EPERM'` — indistinguishable from a real file lock
    // by code alone, so without the target-is-a-directory check this used to
    // surface the misleading "this file is open in another program" message
    // instead of `classifyWriteError`'s intended "that's a folder" one.
    const target = path.join(tempDir, 'a-folder')
    fs.mkdirSync(target)

    let thrown: unknown
    try {
      atomicWriteFile(target, 'new content')
    } catch (err) {
      thrown = err
    }

    expect(thrown).not.toBeInstanceOf(FileLockedError)
    expect((thrown as { code?: string } | undefined)?.code).toBe('EISDIR')
    expect(classifyWriteError(thrown)).toBe('That location is a folder, not a file — choose a different name.')
    // The directory itself must be left alone.
    expect(fs.statSync(target).isDirectory()).toBe(true)
  })

  // Windows-only: POSIX rename() checks the DIRECTORY's permissions, so
  // replacing a read-only file there simply succeeds and there is no error to
  // classify. Atlas ships on Windows; CI's Linux job would see no throw.
  it.skipIf(process.platform !== 'win32')('classifies a rename onto a read-only file as EROFS, not a file lock (found by driving the real app)', () => {
    // A file with Windows's read-only attribute set (`attrib +R` /
    // `fs.chmodSync(path, 0o444)`) also makes `fs.renameSync` throw EPERM —
    // the exact same ambiguity as the EISDIR case above, and just as
    // misleading: without this check, a plain read-only file was reported
    // as "open in another program" (sending the user to close a program
    // that was never open) instead of the accurate "read-only" message.
    const target = path.join(tempDir, 'readonly.md')
    fs.writeFileSync(target, 'original')
    fs.chmodSync(target, 0o444)

    let thrown: unknown
    try {
      atomicWriteFile(target, 'new content')
    } catch (err) {
      thrown = err
    } finally {
      fs.chmodSync(target, 0o666)
    }

    expect(thrown).not.toBeInstanceOf(FileLockedError)
    expect((thrown as { code?: string } | undefined)?.code).toBe('EROFS')
    expect(classifyWriteError(thrown)).toBe('That location is read-only — choose a different location.')
    // The read-only file itself must be left alone.
    expect(fs.readFileSync(target, 'utf-8')).toBe('original')
  })

  it('classifies the EXDEV fallback writing onto an existing directory as EISDIR too', () => {
    const target = path.join(tempDir, 'a-folder')
    fs.mkdirSync(target)

    vi.spyOn(fs, 'renameSync').mockImplementationOnce(() => {
      const err = Object.assign(new Error('cross-device link'), { code: 'EXDEV' })
      throw err
    })

    let thrown: unknown
    try {
      atomicWriteFile(target, 'new content')
    } catch (err) {
      thrown = err
    }

    expect(thrown).not.toBeInstanceOf(FileLockedError)
    expect((thrown as { code?: string } | undefined)?.code).toBe('EISDIR')
  })
})

describe('classifyWriteError', () => {
  function withCode(code: string): unknown {
    return Object.assign(new Error('boom'), { code })
  }

  it.each([
    ['EACCES', "Permission denied — you don't have access to save to this location."],
    ['ENOSPC', 'Not enough disk space to save this file.'],
    ['EISDIR', 'That location is a folder, not a file — choose a different name.'],
    ['ENOENT', 'The destination folder no longer exists — choose a different location.'],
    ['EROFS', 'That location is read-only — choose a different location.'],
    ['ENAMETOOLONG', 'That file name or path is too long — choose a shorter one.'],
  ])('maps %s to a friendly message', (code, expected) => {
    expect(classifyWriteError(withCode(code))).toBe(expected)
  })

  it('returns undefined for an unrecognized error code, so the caller can fall back to its own message', () => {
    expect(classifyWriteError(withCode('EWEIRD'))).toBeUndefined()
  })

  it('returns undefined for a non-Error value with no code', () => {
    expect(classifyWriteError('just a string')).toBeUndefined()
    expect(classifyWriteError(null)).toBeUndefined()
    expect(classifyWriteError(undefined)).toBeUndefined()
  })
})
