# Strict-mode ratchet (P4.1)

Status as of 2026-09-20. **Complete** — every source tree `tsc -b` walks
(app, node config, e2e, electron main process, and now build/codegen
scripts) compiles clean under `strict: true` with zero errors. There is
nothing left on this ratchet.

## Strict today

TypeScript 6 turns `strict` on by default, and the codebase already compiles clean
under it. `strict: true` is now written out explicitly so a toolchain change can never
loosen it silently:

| Project | Covers | Strict |
|---|---|---|
| `tsconfig.app.json` | all of `src/` (renderer, DOCX engine, viewers, tests) | yes, 0 errors |
| `tsconfig.node.json` | `vite.config.ts` | yes, 0 errors |
| `tsconfig.e2e.json` | `tests/e2e/**/*.ts`, `playwright.config.ts` | yes, 0 errors |
| `tsconfig.electron.json` | `electron/**/*.cjs` (`main.cjs`, `preload.cjs`, `lib/*.cjs`) | yes, `checkJs`, 0 errors (P4.1) |
| `tsconfig.scripts.json` | `scripts/**/*.mjs` (8 files), `tests/e2e/fixtures/*.mjs` | yes, `checkJs`, 0 errors (new — P4.1) |

All five are built by `npx tsc -b`, which CI runs on every push.

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

## `scripts/*.mjs` + `tests/e2e/fixtures/*.mjs` (P4.1, closing the ratchet)

`tsconfig.scripts.json` covers `scripts/**/*.mjs` (8 files: the 7 top-level
generator/validator CLIs plus `scripts/lib/*.mjs`) and
`tests/e2e/fixtures/generate.mjs`, and is referenced from `tsconfig.json` so
`npx tsc -b` builds it on every push same as the other four projects.

This is ESM (`"type": "module"`), unlike `electron/`'s CommonJS, so it uses
`module`/`moduleResolution: "bundler"` — the same bundler-mode block
`tsconfig.app.json` uses — rather than `electron.json`'s `node16`. That
choice is forced by `scripts/generate-extension-manifest.mjs`, which imports
`src/formats/extensionManifest.ts` directly by its `.ts` extension (Node
24's native TypeScript type-stripping, no build step); that import is only
legal under bundler resolution with `allowImportingTsExtensions`, and
checking it under `node16` resolution produced spurious cross-project
errors unrelated to any real bug (the file is already covered, correctly,
by `tsconfig.app.json`).

The ~90 errors this first run surfaced were almost entirely
`noImplicitAny` on undeclared function parameters (JSDoc `@param`/`@returns`
added throughout), plus:
- `scripts/lib/officeValidator.mjs` (71 of the ~90 — by far the largest
  single file in this project) needed real shape types for the
  `fast-xml-parser` `preserveOrder` node tree (`OrderedNode`, a recursive
  tag-keyed type) and for the three fixed OPC/ODF control-file schemas it
  parses "collapsed" (`ContentTypesTree`, `RelationshipsTree`,
  `ManifestTree`) — `fast-xml-parser`'s own `.parse()` is typed `any`, so
  those shapes are asserted at the handful of call sites that read known
  fields off it, not threaded through every helper.
- `scripts/lib/corpusArchive.mjs`'s `zipFromFiles` read `files.get(path)`
  for a `path` it had just pulled from `files.keys()` and passed the
  possibly-`undefined` result straight to `JSZip#file`, which strict mode
  correctly rejected; it now skips a path whose content is missing instead
  of assuming it can't happen.
- `scripts/generate-templates.mjs`'s `loadDocxModule` loads
  `src/docx/index.ts` through Vite's SSR module graph, which Vite types as
  `Record<string, any>` (it can load anything); its two real exports
  (`loadDocx`/`saveDocx`) are asserted once, at that one function's return,
  into a local `DocxRoundTripModule` shape instead of threading `any`
  through every caller.
- `scripts/generate-docx-corpus.mjs` had one genuinely dead import
  (`WidthType`, imported from `docx` but never used) that `noUnusedLocals`
  caught and this removed.

No `@ts-ignore`, `@ts-nocheck`, or `any` widening was used; `scripts/lib/officeValidator.mjs`'s
`parseSimpleXml` helper keeps its (correctly) untyped `any` return — that's
what `fast-xml-parser`'s own `.d.ts` declares `XMLParser#parse` to return —
but every caller narrows it to a specific `*Tree` shape via a `@type` cast
before reading any field off it, and `asList` is a generic helper
(`<T>(value: T | T[] | undefined) => T[]`) rather than an `any[]` shim.

All of `officeValidator.mjs`'s and `zipReader.mjs`'s unit tests (78 tests
across `src/office/__tests__/officeValidator.unit.test.ts`,
`src/docx/__tests__/validateOfficeFile.corpus.test.ts`,
`src/docx/__tests__/headerFooterMixed.officeValidation.test.ts`, and the
three viewer-side `officeFileValidation.test.ts` suites) pass unchanged, and
`node scripts/validate-office-file.mjs` still validates every `.docx`/
`.xlsx`/`.pptx`/`.odp`/`.odt` sample fixture with zero violations — the
type annotations added no behavior change.

## Remaining

None. Every project `tsc -b` builds is `strict: true` with 0 errors.

Rule: a project that compiles clean stays clean. Never add `strict: false`, per-file
`@ts-nocheck`, or loosen a flag to land a change.
