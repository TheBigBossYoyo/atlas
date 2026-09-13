import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import {
  DEFAULT_MAX_BYTES,
  FileTooLargeError,
  assertFileSizeAllowed,
  assertSizeAllowed,
} from '../../../electron/lib/fileSizeGuard.cjs'

describe('assertSizeAllowed', () => {
  it('does not throw for a size under the cap', () => {
    expect(() => assertSizeAllowed(1024, 2048)).not.toThrow()
  })

  it('does not throw for a size exactly at the cap', () => {
    expect(() => assertSizeAllowed(2048, 2048)).not.toThrow()
  })

  it('throws a friendly FileTooLargeError over the cap', () => {
    expect(() => assertSizeAllowed(3000, 2048)).toThrow(FileTooLargeError)
  })

  it('includes human-readable megabyte sizes in the message', () => {
    try {
      assertSizeAllowed(300 * 1024 * 1024, 200 * 1024 * 1024)
      throw new Error('expected assertSizeAllowed to throw')
    } catch (err) {
      expect(err).toBeInstanceOf(FileTooLargeError)
      expect((err as Error).message).toContain('300.0 MB')
      expect((err as Error).message).toContain('200.0 MB')
    }
  })

  it('uses the default cap when none is provided', () => {
    expect(() => assertSizeAllowed(DEFAULT_MAX_BYTES + 1)).toThrow(FileTooLargeError)
    expect(() => assertSizeAllowed(DEFAULT_MAX_BYTES)).not.toThrow()
  })
})

describe('assertFileSizeAllowed', () => {
  let tempDir: string

  beforeEach(() => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-file-size-guard-'))
  })

  afterEach(() => {
    fs.rmSync(tempDir, { recursive: true, force: true })
  })

  it('does not throw for a small file', () => {
    const filePath = path.join(tempDir, 'small.txt')
    fs.writeFileSync(filePath, 'hello')
    expect(() => assertFileSizeAllowed(filePath, 1024)).not.toThrow()
  })

  it('throws for a file over the given cap', () => {
    const filePath = path.join(tempDir, 'big.txt')
    fs.writeFileSync(filePath, 'x'.repeat(2048))
    expect(() => assertFileSizeAllowed(filePath, 1024)).toThrow(FileTooLargeError)
  })
})
