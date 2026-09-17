/**
 * USR-19 — electron/lib/codeRunner.cjs: runtime resolution, environment
 * scrubbing, output streaming/cap, runtime fallback, stop and timeout.
 * The child process itself is faked; the real thing is exercised end to end
 * by tests/e2e/code-editor.spec.ts.
 */
import { EventEmitter } from 'node:events'
import { PassThrough } from 'node:stream'
import path from 'node:path'

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { childEnv, createCodeRunner, runtimeFor, MAX_OUTPUT_BYTES } from '../../../electron/lib/codeRunner.cjs'

type FakeChild = EventEmitter & {
  stdout: PassThrough
  stderr: PassThrough
  exitCode: number | null
  pid: number
  kill: (signal?: string) => void
}

function createChild(): FakeChild {
  const child = new EventEmitter() as FakeChild
  child.stdout = new PassThrough()
  child.stderr = new PassThrough()
  child.exitCode = null
  child.pid = 4242
  child.kill = vi.fn()
  return child
}

type Event = { channel: string; payload: Record<string, unknown> }

function createRunner(options: { timeoutMs?: number } = {}) {
  const events: Event[] = []
  const children: FakeChild[] = []
  const spawns: Array<{ command: string; args: string[]; options: Record<string, unknown> }> = []
  const spawnImpl = vi.fn((command: string, args: string[], spawnOptions: Record<string, unknown>) => {
    spawns.push({ command, args, options: spawnOptions })
    const child = createChild()
    children.push(child)
    return child
  })
  const runner = createCodeRunner({
    send: (channel: string, payload: Record<string, unknown>) => events.push({ channel, payload }),
    spawnImpl,
    timeoutMs: options.timeoutMs ?? 60_000,
  })
  return { runner, events, children, spawns, spawnImpl }
}

const outputs = (events: Event[]): string =>
  events.filter((event) => event.channel === 'code:run-output').map((event) => String(event.payload.text)).join('')

const exitEvent = (events: Event[]): Record<string, unknown> | undefined =>
  events.find((event) => event.channel === 'code:run-exit')?.payload

beforeEach(() => {
  vi.useFakeTimers()
})

afterEach(() => {
  vi.useRealTimers()
})

describe('runtimeFor', () => {
  it('runs JavaScript and TypeScript with the bundled Node runtime', () => {
    expect(runtimeFor('C:/tmp/a.js')?.candidates[0].command).toBe(process.execPath)
    expect(runtimeFor('C:/tmp/a.mjs')?.candidates[0].env).toEqual({ ELECTRON_RUN_AS_NODE: '1' })
    expect(runtimeFor('C:/tmp/a.ts')?.candidates[0].args).toEqual([
      '--experimental-strip-types',
      '--no-warnings',
      'C:/tmp/a.ts',
    ])
  })

  it('offers the usual Python launchers, and nothing for other files', () => {
    expect(runtimeFor('C:/tmp/a.py')?.candidates.map((candidate) => candidate.command)).toEqual(['py', 'python', 'python3'])
    expect(runtimeFor('C:/tmp/a.docx')).toBeNull()
    expect(runtimeFor('C:/tmp/a.sh')).toBeNull()
  })
})

describe('childEnv', () => {
  it('drops Electron variables from the inherited environment', () => {
    process.env.ELECTRON_TEST_ONLY = 'leaked'
    process.env.ATLAS_TEST_KEEP = 'kept'
    try {
      const env = childEnv({ ELECTRON_RUN_AS_NODE: '1' })
      expect(env.ELECTRON_TEST_ONLY).toBeUndefined()
      expect(env.ATLAS_TEST_KEEP).toBe('kept')
      // Only what the runtime explicitly asks for is added back.
      expect(env.ELECTRON_RUN_AS_NODE).toBe('1')
    } finally {
      delete process.env.ELECTRON_TEST_ONLY
      delete process.env.ATLAS_TEST_KEEP
    }
  })
})

describe('createCodeRunner', () => {
  it('spawns without a shell in the file folder and streams output until exit', async () => {
    const { runner, events, children, spawns } = createRunner()
    const filePath = path.join('C:', 'work', 'script.js')
    expect(runner.start(filePath)).toEqual({ ok: true, runId: 1 })

    expect(spawns[0].options).toMatchObject({ shell: false, cwd: path.dirname(filePath), windowsHide: true })
    const child = children[0]
    child.emit('spawn')
    child.stdout.write('hello\n')
    child.stderr.write('warn\n')
    await vi.advanceTimersByTimeAsync(0)
    child.emit('close', 0)

    expect(outputs(events)).toContain('hello')
    expect(outputs(events)).toContain('warn')
    expect(exitEvent(events)).toMatchObject({ runId: 1, code: 0, timedOut: false, stopped: false })
  })

  it('refuses a second run while one is in flight, and a file it cannot run', () => {
    const { runner } = createRunner()
    expect(runner.start('C:/work/a.js').ok).toBe(true)
    expect(runner.start('C:/work/b.js')).toMatchObject({ ok: false })
    expect(runner.start('C:/work/a.pdf')).toMatchObject({ ok: false })
  })

  it('falls back to the next launcher when one is not installed', async () => {
    const { runner, events, children, spawns } = createRunner()
    runner.start('C:/work/script.py')
    const missing = Object.assign(new Error('spawn py ENOENT'), { code: 'ENOENT' })
    children[0].emit('error', missing)
    await vi.advanceTimersByTimeAsync(0)

    expect(spawns.map((spawn) => spawn.command)).toEqual(['py', 'python'])
    children[1].emit('spawn')
    children[1].emit('close', 0)
    expect(exitEvent(events)).toMatchObject({ code: 0 })
  })

  it('reports a clear error when no launcher is installed at all', async () => {
    const { runner, events, children } = createRunner()
    runner.start('C:/work/script.py')
    for (let i = 0; i < 3; i++) {
      children[i].emit('error', Object.assign(new Error('ENOENT'), { code: 'ENOENT' }))
      await vi.advanceTimersByTimeAsync(0)
    }
    expect(String(exitEvent(events)?.error)).toContain('Python')
  })

  it('caps output instead of streaming an endless program into the renderer', async () => {
    const { runner, events, children } = createRunner()
    runner.start('C:/work/loud.js')
    children[0].emit('spawn')
    children[0].stdout.write('x'.repeat(MAX_OUTPUT_BYTES + 10))
    children[0].stdout.write('more')
    await vi.advanceTimersByTimeAsync(0)

    const text = outputs(events)
    expect(text).toContain('Output limit reached')
    expect(text).not.toContain('more')
  })

  it('kills the program on stop and after the time limit', async () => {
    const { runner, events, children } = createRunner({ timeoutMs: 1000 })
    const started = runner.start('C:/work/slow.js')
    if (!started.ok) throw new Error('the run should have started')
    children[0].emit('spawn')
    expect(runner.stop(started.runId)).toBe(true)
    children[0].emit('close', null)
    expect(exitEvent(events)).toMatchObject({ stopped: true })

    const second = createRunner({ timeoutMs: 1000 })
    second.runner.start('C:/work/slow.js')
    second.children[0].emit('spawn')
    await vi.advanceTimersByTimeAsync(1001)
    second.children[0].emit('close', null)
    expect(exitEvent(second.events)).toMatchObject({ timedOut: true })
  })
})
