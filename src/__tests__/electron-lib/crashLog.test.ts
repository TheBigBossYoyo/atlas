import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import {
  MAX_LOG_BYTES,
  appendLogLine,
  formatLogLine,
  getLogFilePath,
  logToFile,
} from '../../../electron/lib/crashLog.cjs'

let logDir: string

beforeEach(() => {
  logDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-crash-log-'))
})

afterEach(() => {
  fs.rmSync(logDir, { recursive: true, force: true })
})

describe('formatLogLine', () => {
  it('includes an ISO timestamp, the level, and the message', () => {
    const line = formatLogLine('ERROR', 'something broke')
    expect(line).toMatch(/^\[\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z\] \[ERROR\] something broke\n$/)
  })

  it('includes the error stack when an Error is passed', () => {
    const line = formatLogLine('ERROR', 'uncaughtException', new Error('boom'))
    expect(line).toContain('uncaughtException')
    expect(line).toContain('Error: boom')
  })
})

describe('appendLogLine / logToFile', () => {
  it('creates the log directory and file on first write', () => {
    logToFile(logDir, 'ERROR', 'first error', new Error('x'))
    const filePath = getLogFilePath(logDir)
    expect(fs.existsSync(filePath)).toBe(true)
    expect(fs.readFileSync(filePath, 'utf-8')).toContain('first error')
  })

  it('appends subsequent lines rather than overwriting', () => {
    logToFile(logDir, 'ERROR', 'first')
    logToFile(logDir, 'ERROR', 'second')
    const content = fs.readFileSync(getLogFilePath(logDir), 'utf-8')
    expect(content).toContain('first')
    expect(content).toContain('second')
  })

  it('rotates (truncates) the file once it would exceed the size cap', () => {
    const filePath = getLogFilePath(logDir)
    fs.mkdirSync(logDir, { recursive: true })
    fs.writeFileSync(filePath, 'x'.repeat(MAX_LOG_BYTES - 10))

    appendLogLine(logDir, formatLogLine('ERROR', 'triggers rotation'))

    const content = fs.readFileSync(filePath, 'utf-8')
    expect(content.length).toBeLessThan(MAX_LOG_BYTES)
    expect(content).toContain('triggers rotation')
    expect(content).not.toContain('xxxxxxxxxx')
  })

  it('never throws even if the log directory cannot be created', () => {
    // A path with a NUL byte is invalid on every platform and forces
    // mkdirSync to throw — appendLogLine must swallow that.
    expect(() => appendLogLine('\0invalid', 'line\n')).not.toThrow()
  })
})
