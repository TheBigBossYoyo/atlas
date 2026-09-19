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
| `tsconfig.e2e.json` | `tests/e2e/**/*.ts`, `playwright.config.ts` | yes, 0 errors (new: these specs were not type-checked before) |

All three are built by `npx tsc -b`, which CI runs on every push.

## Remaining

| Directory | Language | Errors under `checkJs` + `strict` | Notes |
|---|---|---|---|
| `electron/` (`main.cjs`, `preload.cjs`, `lib/*.cjs`) | CommonJS JavaScript with JSDoc | 94 (preload 33, main 22, codeRunner 22, systemFonts 8, printToPdf 5, 4 others 1 each) | Not type-checked at all today. Security-sensitive main-process code, so convert one file at a time with its unit tests, starting with the small `lib/` modules. |
| `scripts/*.mjs`, `tests/e2e/fixtures/*.mjs` | JavaScript | not measured | Build/fixture tooling; lowest priority. |

Rule: a project that compiles clean stays clean. Never add `strict: false`, per-file
`@ts-nocheck`, or loosen a flag to land a change.
