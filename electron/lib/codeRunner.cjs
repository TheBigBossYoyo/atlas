'use strict';

/**
 * USR-19 — explicit "Run" for an opened source file.
 *
 * Security model (see the register's USR-19 entry):
 *  - only a file already on the main-process path allowlist can run, and the
 *    file on disk is what runs (never code text sent by the renderer);
 *  - main.cjs asks for confirmation in a native dialog the renderer cannot
 *    answer, once per file per session;
 *  - the program is a separate child process with no shell (no command-line
 *    injection), its working directory is the file's folder, Electron's own
 *    environment variables are stripped, stdin is closed, output is capped and
 *    the whole process tree is killed on Stop, on timeout and on quit.
 */

const path = require('node:path');
const { spawn, execFile } = require('node:child_process');

const DEFAULT_TIMEOUT_MS = 60_000;
const MAX_OUTPUT_BYTES = 1024 * 1024;

/**
 * @typedef {{ command: string, args: string[], env?: Record<string, string> }} Candidate
 * @typedef {{ language: string, candidates: Candidate[] }} Runtime
 */

/** @returns {Runtime | null} */
function runtimeFor(filePath) {
  const ext = path.extname(filePath).slice(1).toLowerCase();
  const node = { command: process.execPath, env: { ELECTRON_RUN_AS_NODE: '1' } };
  switch (ext) {
    case 'js':
    case 'mjs':
    case 'cjs':
      return { language: 'JavaScript (Node.js)', candidates: [{ ...node, args: [filePath] }] };
    case 'ts':
    case 'mts':
    case 'cts':
      return {
        language: 'TypeScript (Node.js type stripping)',
        candidates: [{ ...node, args: ['--experimental-strip-types', '--no-warnings', filePath] }],
      };
    case 'py':
    case 'pyw':
      return {
        language: 'Python',
        candidates: [
          { command: 'py', args: ['-3', filePath] },
          { command: 'python', args: [filePath] },
          { command: 'python3', args: [filePath] },
        ],
      };
    default:
      return null;
  }
}

/** The parent environment without Electron's own switches (a child must not inherit e.g. ELECTRON_RUN_AS_NODE unasked). */
function childEnv(extra) {
  const env = {};
  for (const [key, value] of Object.entries(process.env)) {
    if (!key.toUpperCase().startsWith('ELECTRON_') && value !== undefined) env[key] = value;
  }
  return { ...env, ...(extra || {}) };
}

function killTree(child) {
  if (!child || child.exitCode !== null || child.pid === undefined) return;
  if (process.platform === 'win32') {
    execFile('taskkill', ['/PID', String(child.pid), '/T', '/F'], { windowsHide: true }, () => {});
  } else {
    child.kill('SIGKILL');
  }
}

/**
 * One run at a time. `send(channel, payload)` delivers events to the renderer:
 * `code:run-output` { runId, stream, text } and `code:run-exit` { runId, code, timedOut, stopped, error }.
 */
function createCodeRunner({ send, spawnImpl = spawn, timeoutMs = DEFAULT_TIMEOUT_MS }) {
  let current = null;
  let nextId = 1;

  function start(filePath) {
    const runtime = runtimeFor(filePath);
    if (!runtime) return { ok: false, error: 'Atlas cannot run this kind of file.' };
    if (current) return { ok: false, error: 'A program is already running. Stop it first.' };

    const runId = nextId++;
    const run = { runId, child: null, bytes: 0, truncated: false, timedOut: false, stopped: false, timer: null };
    current = run;

    const emitOutput = (stream, chunk) => {
      if (run.truncated) return;
      const remaining = MAX_OUTPUT_BYTES - run.bytes;
      const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(String(chunk));
      run.bytes += buffer.length;
      if (buffer.length > remaining) {
        run.truncated = true;
        send('code:run-output', { runId, stream, text: buffer.subarray(0, remaining).toString('utf8') });
        send('code:run-output', { runId, stream: 'stderr', text: '\n[Atlas] Output limit reached; further output is hidden.\n' });
        return;
      }
      send('code:run-output', { runId, stream, text: buffer.toString('utf8') });
    };

    const finish = (payload) => {
      if (current !== run) return;
      clearTimeout(run.timer);
      current = null;
      send('code:run-exit', { runId, timedOut: run.timedOut, stopped: run.stopped, ...payload });
    };

    const tryCandidate = (index) => {
      const candidate = runtime.candidates[index];
      if (!candidate) {
        finish({ code: null, error: `${runtime.language} was not found on this computer.` });
        return;
      }
      let child;
      try {
        child = spawnImpl(candidate.command, candidate.args, {
          cwd: path.dirname(filePath),
          env: childEnv(candidate.env),
          shell: false,
          windowsHide: true,
          stdio: ['ignore', 'pipe', 'pipe'],
        });
      } catch (err) {
        finish({ code: null, error: String(err && err.message ? err.message : err) });
        return;
      }
      run.child = child;
      let spawned = false;
      child.once('spawn', () => {
        spawned = true;
      });
      child.stdout?.on('data', (chunk) => emitOutput('stdout', chunk));
      child.stderr?.on('data', (chunk) => emitOutput('stderr', chunk));
      child.once('error', (err) => {
        if (!spawned && err && err.code === 'ENOENT') {
          tryCandidate(index + 1);
          return;
        }
        finish({ code: null, error: String(err && err.message ? err.message : err) });
      });
      child.once('close', (code) => {
        if (spawned) finish({ code });
      });
    };

    run.timer = setTimeout(() => {
      run.timedOut = true;
      killTree(run.child);
    }, timeoutMs);

    send('code:run-output', { runId, stream: 'system', text: `[Atlas] Running ${path.basename(filePath)} with ${runtime.language}\n` });
    tryCandidate(0);
    return { ok: true, runId };
  }

  function stop(runId) {
    if (!current || current.runId !== runId) return false;
    current.stopped = true;
    killTree(current.child);
    return true;
  }

  function dispose() {
    if (current) {
      current.stopped = true;
      killTree(current.child);
    }
  }

  return { start, stop, dispose };
}

module.exports = { createCodeRunner, runtimeFor, childEnv, MAX_OUTPUT_BYTES };
