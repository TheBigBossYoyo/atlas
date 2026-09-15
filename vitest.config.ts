import { defineConfig, type Plugin } from 'vitest/config'
import react from '@vitejs/plugin-react'

const PUBLIC_ASSET_URL_PREFIX = '\0public-asset-url:'

// `src/docx/fonts/families.ts` imports bundled fonts as `/fonts/X.ttf?url`.
// Vite's build resolves those from `public/`, but the vitest module runner
// rejects them ("Denied ID"), which prevented every suite that transitively
// imports the layout engine from loading. Tests only need the URL string, so
// resolve public-root `?url` imports to a module exporting that path.
function publicAssetUrlStub(): Plugin {
  return {
    name: 'atlas:public-asset-url-stub',
    enforce: 'pre',
    resolveId(id) {
      if (id.startsWith('/') && id.endsWith('?url')) {
        return `${PUBLIC_ASSET_URL_PREFIX}${id.slice(0, -'?url'.length)}`
      }
      return null
    },
    load(id) {
      if (id.startsWith(PUBLIC_ASSET_URL_PREFIX)) {
        return `export default ${JSON.stringify(id.slice(PUBLIC_ASSET_URL_PREFIX.length))}`
      }
      return null
    },
  }
}

// Mirrors vite.config.ts (no aliases configured there yet). Keep in sync.
export default defineConfig({
  plugins: [publicAssetUrlStub(), react()],
  test: {
    environment: 'jsdom',
    globals: true,
    setupFiles: ['./src/__tests__/setup.ts'],
    include: ['src/**/*.{test,spec}.{ts,tsx}'],
    css: false,
    // P4.2/QA-03/QA-04 — coverage was previously invisible: no `coverage`
    // script existed, so `npx vitest run --coverage` was the only way to
    // even see a number, and a failing test suite silently suppressed the
    // report entirely (no signal that coverage was untrustworthy that run).
    coverage: {
      provider: 'v8',
      reporter: ['text', 'html'],
      // Show the report even when a test fails, instead of silently
      // skipping it — a failing suite is exactly when you most want to see
      // what did/didn't get exercised.
      reportOnFailure: true,
      include: ['src/**/*.{ts,tsx}'],
      exclude: [
        'src/**/*.{test,spec}.{ts,tsx}',
        'src/__tests__/**',
        'src/**/__tests__/**',
        'src/**/*.d.ts',
        'src/main.tsx',
      ],
      // Ratchet floor — P4.2/QA-04. Measured via `npm run coverage` on
      // 2026-09-15 against wave1+wave2 (main @ 03b2e2a) plus this wave's own
      // changes: statements 76.71%, branches 63.69%, functions 80.97%,
      // lines 78.54% (1844+ tests). Each threshold below is that measured
      // number rounded DOWN to the next whole percent, so today's suite
      // clears the gate with a small margin for incidental variance rather
      // than sitting exactly on the edge. This is a floor, not a target:
      // CI fails a PR that drops below it, but nothing stops coverage from
      // climbing well past these numbers as more tests land — whenever a
      // change pushes the measured number up, lower-bound-round the new
      // number and raise the threshold in the same PR (see
      // docs/KNOWN_LIMITATIONS.md's "Coverage ratchet" note) rather than
      // waiting for a dedicated "raise the thresholds" task.
      thresholds: {
        statements: 76,
        branches: 63,
        functions: 80,
        lines: 78,
      },
    },
  },
})
