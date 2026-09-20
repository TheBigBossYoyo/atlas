import { defineConfig } from 'vitest/config'

// Standalone Vitest config for `scripts/lib/officeValidator.mjs`'s own test
// suite. The project's root `vitest.config.ts` scopes `test.include` to
// `src/**/*.{test,spec}.{ts,tsx}` (this validator's own `.mjs` tests live
// under `scripts/`, outside that glob, and that file isn't ours to edit
// here -- see the worktree's task scope) -- so a plain `npx vitest run
// scripts/...` finds nothing to collect. This config is scoped narrowly to
// `scripts/**` so it never picks up (or affects) anything under `src/`.
//
// Run with:
//   npx vitest run --config scripts/lib/__tests__/vitest.config.mjs --maxWorkers=1 [path...]
export default defineConfig({
  test: {
    environment: 'node',
    include: ['scripts/**/*.{test,spec}.mjs'],
    globals: true,
  },
})
