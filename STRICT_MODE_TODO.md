# Strict-mode ratchet (P4.1)

Status as of 2026-09-19.

## Strict today

TypeScript 6 turns `strict` on by default, and the codebase already compiles clean
under it. `strict: true` is now written out explicitly so a toolchain change can never
loosen it silently:

| Project | Covers | Strict |
|---|---|---|
| `tsconfig.app.json` | all of `src/` (renderer, DOCX engine, viewers, tests) | yes, 0 errors |
| `tsconfig.node.json` | `vite.config.ts` | yes, 0 errors |
| `tsconfig.e2e.json` | `tests/e2e/**/*.ts`, `playwright.config.ts` | yes, 0 errors |
| `tsconfig.electron.json` | `electron/**/*.cjs` (`main.cjs`, `preload.cjs`, `lib/*.cjs`) | yes, `checkJs`, 0 errors (new — P4.1) |

All four are built by `npx tsc -b`, which CI runs on every push.

`electron/` is plain CommonJS JavaScript with JSDoc types (not `.ts`), so its
project additionally sets `allowJs`/`checkJs` with `module`/`moduleResolution:
"node16"` (CommonJS is what actually runs under Electron's main process) instead
of the bundler-mode settings the other three projects use. The ~95 errors this
first run surfaced (preload 34, main 22, codeRunner 22, systemFonts 8,
printToPdf 5, atomicWrite/csp/fileSizeGuard/recentFilesStore 1 each) were fixed
file by file, smallest first, with real JSDoc parameter/return types and
explicit null/undefined narrowing — no `@ts-ignore`, `@ts-nocheck`, or `any`
widening. Two of those fixes were genuine latent-bug corrections, not just type
annotations:
- `main.cjs`'s `set-theme` IPC handler indexed the theme→overlay-color table
  with the renderer-supplied value without checking it was one of the known
  keys; it now validates with a type guard first, consistent with every other
  untrusted-IPC-input handler in the file.
- `main.cjs`'s `ready-to-show` window handler called `mainWindow.show()` with
  no null/destroyed guard, unlike its sibling `did-fail-load`/fallback-show
  handlers, which already defend against the window having closed in the
  async gap before the event fires.

Several `electron/lib/*.cjs` modules already ship a hand-maintained sibling
`*.d.cts` (e.g. `atomicWrite.d.cts`) purely so `.ts` unit tests importing them
get types; those are a separate, narrower surface from what `checkJs` now
verifies against the `.cjs` implementation itself, and were left as-is.

## Remaining

| Directory | Language | Errors under `checkJs` + `strict` | Notes |
|---|---|---|---|
| `scripts/*.mjs` (7 files) | JavaScript | not measured | Build/codegen tooling (extension manifest, templates, etc.); lowest priority. |
| `tests/e2e/fixtures/*.mjs` (1 file) | JavaScript | not measured | E2E fixture helper; lowest priority. |

Rule: a project that compiles clean stays clean. Never add `strict: false`, per-file
`@ts-nocheck`, or loosen a flag to land a change.
