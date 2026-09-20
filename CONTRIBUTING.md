# Contributing to Atlas

Practical, hard-won rules for working in this repo — most of them exist
because skipping them once already broke something (a hung laptop, a
corrupted worktree, a "flaky" CI failure that was actually a bug). If one of
these stops applying, fix the rule here in the same change that makes it
stop applying — don't leave it to bite the next contributor (human or
agent).

## Environment

- **Windows 11.** PowerShell is the primary shell; Git Bash is also
  available. Bash heredocs mangle backslashes, quotes and `\u` escapes on
  this platform — use an editor/Write-Edit tool for source files instead of
  piping heredocs into them.
- **Node version is pinned** — `.nvmrc` (currently `24.14.0`) and
  `package.json`'s `engines.node` are the single source of truth; CI's
  `setup-node` step reads `.nvmrc` directly. Don't drift from it locally.

## Multiple agents / worktrees working at once

Atlas is regularly developed by several agents in parallel git worktrees,
each targeting a different area (`src/viewers/**`, `src/docx/**`,
`src/i18n/**`, `electron/**`, docs/scripts). A few rules keep that from
colliding:

- **A worktree has no `node_modules` of its own** — it's a junction to the
  main checkout's install:
  ```
  cmd /c mklink /J node_modules <path-to-main-checkout>\node_modules
  ```
  **Never** `rm -rf`/`Remove-Item -Recurse` a path that goes *through* that
  junction — it deletes the shared install out from under every other
  worktree and session using it. To remove the junction itself (e.g. before
  `git worktree remove`), unlink it without recursing into its target:
  ```
  cmd /c rmdir <worktree-path>\node_modules
  ```
- **Never run `npm install`/`npm ci`** from inside a worktree — it's not
  your install to modify, and it fights every other agent using the same
  `node_modules` for CPU/disk.
- **Run targeted tests only**, not the full suite, while other agents are
  active: `npx vitest run <specific paths> --maxWorkers=1` (or `=3` when
  running alone — see below). Never `npm run build`, `electron-builder`, or
  a full Playwright run from a worktree unless you're the one session
  actually validating a release-shaped change.

## Running tests

- **Vitest**: `npx vitest run --maxWorkers=3` for the full suite on a single
  machine (not the default worker count — this machine hard-powers-off
  under full load if Vitest, a Vite/Electron build, and Playwright all run
  concurrently; keep it to one of the three at a time). Prefer
  `--maxWorkers=1` plus a narrow path list when another process might be
  competing for CPU.
- **Playwright**: `workers: 1` is set in `playwright.config.ts` and must
  stay that way — Atlas takes an Electron single-instance lock, so a second
  `_electron.launch()` from a parallel worker just quits immediately rather
  than testing anything.
- **Never run a Vite/Electron build, Vitest, and Playwright at the same
  time.** Pick one. This is the single most common way to lock up the dev
  machine.
- **Build before an Electron probe or e2e run**: `electron/main.cjs`
  decides dev vs. prod by checking whether `dist/index.html` exists on
  disk, not by how the process was launched (see
  `docs/KNOWN_LIMITATIONS.md`'s "Dev tooling" section) — so run
  `npx vite build` first, or set `ATLAS_DEV=1` to force the dev server
  regardless of what's on disk.
- **Running e2e rewrites the fixture corpus.** `tests/e2e/fixtures/generate.mjs`
  regenerates `tests/e2e/fixtures/*` as part of the suite, which dirties
  tracked binary fixtures. **Revert them before committing**:
  ```
  git checkout -- tests/e2e/fixtures/
  ```
  (or `git status` first to confirm nothing else in that directory was
  intentionally changed).

## The bundle-size gate

`npm run build` runs `scripts/check-bundle.mjs` as a `postbuild` step: it
compares every `dist/assets/*.{js,css}` chunk (not just the main entry
bundle) against `.sisyphus/baselines/atlas-phase3-bundle.json` and fails the
build if any chunk, or the total, grew more than 25%. If your change
intentionally grows the bundle (a new dependency, a genuinely bigger
feature), refresh the baseline as part of the same commit:
```
node scripts/capture-bundle-baseline.mjs
```
Don't refresh the baseline to silence a regression you haven't looked at —
the gate exists because bundle growth is otherwise invisible until a user
notices the app got slower to load.

## Type-checking and lint

`npx tsc -b` and `npx eslint src electron tests scripts` must both be clean
before any change lands — this is what CI runs, and it now covers every
source tree in the repo: `src/`, `vite.config.ts`, `tests/e2e/**/*.ts`,
`electron/**/*.cjs`, and `scripts/**/*.mjs` + `tests/e2e/fixtures/*.mjs`
(see `STRICT_MODE_TODO.md`). All five projects are `strict: true`. Never
add `@ts-ignore`, `@ts-nocheck`, `strict: false`, or an `any` widening to
make an error go away — fix the type, or narrow the value, instead.

## A CI-only flake is usually a real bug, not a flaky test

This has happened twice with the same root cause: a Playwright spec passed
locally and failed only on CI's more loaded runner, and the instinct was to
add a longer timeout or a retry. Both times, digging in found an actual
product race instead. The clearest example: the "shortcut-collision" e2e
spec failed on CI twice before anyone looked closely — the PDF find bar's
`Escape` handler only closed the bar once its input had focus, but the
input didn't take focus until ~50ms after the bar opened, so an `Escape`
pressed in that window left the bar open, and its still-focused input then
silently swallowed the *next* keyboard shortcut the test sent (see commit
`314b267`, `fix(pdf): close the find bar on Escape wherever the focus is`).
The fix was in the product, not the test.

**When an e2e spec flakes only under CI's load and not locally, spend a few
minutes looking for a genuine timing race in the feature before reaching
for a longer timeout or a retry.** A timeout bump can be the right call
(rendering work that's legitimately slower on a loaded runner, e.g. PDF
export), but treat that as the fallback, not the first move.

## Commits

Conventional commit format, ending with a blank line then the attribution
line the assistant/session was given for that change (varies by which
Claude model/session made it — check the system reminder or prior commits
on the branch rather than assuming). Create new commits rather than
amending; never skip hooks (`--no-verify`) or force-push `main`.

## Documentation stays truthful

If you land a user-visible or architectural change, update the relevant doc
in the same change: `README.md` (feature/shortcut/stack tables),
`docs/ARCHITECTURE.md` (shared-shell design), `docs/KNOWN_LIMITATIONS.md`
(per-format gaps — add what's now possible, keep what's still a real gap),
and `CHANGELOG.md`. Every claim in these files should be one you verified
against the code, not prose carried forward from an earlier state of the
app — see `docs/KNOWN_LIMITATIONS.md`'s own header for the same rule
applied to itself.
