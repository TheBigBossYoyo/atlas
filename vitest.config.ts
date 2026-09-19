import fs from 'node:fs'
import path from 'node:path'
import { defineConfig, type Plugin } from 'vitest/config'
import react from '@vitejs/plugin-react'

const PUBLIC_ASSET_URL_PREFIX = '\0public-asset-url:'

// Mirrors vite.config.ts's `resolveNodeModulesAllowRoot` — this worktree's
// `node_modules` is a directory junction into a shared install outside the
// project root, and Vite's module runner resolves the junction to its real
// target before applying `server.fs.allow`, rejecting it ("Denied ID") for
// anything not pre-bundled. Bit us for real once X1 added `?raw` CSS imports
// straight from `node_modules` (`katex/dist/katex.min.css`,
// `highlight.js/styles/github.min.css` in `pdf.ts`) — every suite importing
// `pdf.ts` failed to load until this was added here too.
function resolveNodeModulesAllowRoot(): string | undefined {
  const nodeModulesPath = path.resolve(__dirname, 'node_modules')
  try {
    return fs.realpathSync(nodeModulesPath)
  } catch {
    return undefined
  }
}

const nodeModulesRealPath = resolveNodeModulesAllowRoot()

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
  server: {
    fs: {
      allow: [__dirname, ...(nodeModulesRealPath ? [nodeModulesRealPath] : [])],
    },
  },
  test: {
    environment: 'jsdom',
    globals: true,
    setupFiles: ['./src/__tests__/setup.ts'],
    include: ['src/**/*.{test,spec}.{ts,tsx}'],
    // `css: false` would disable Vitest's CSS pipeline for ordinary
    // side-effect `.css` imports (component styles the tests don't need
    // real values for) — but Vitest's own css-disable plugin matches on
    // bare `.css` extension and ignores the `?raw`/`?inline` query, so a
    // blanket `false` ALSO forces every `?raw` stylesheet import
    // (`pdf.ts`/`slidesPdf.ts`'s harvested `viewer-docx.css`, `katex/dist/
    // katex.min.css`, etc.) to resolve as an empty string in every test —
    // silently, since nothing asserted on that content until X1's PDF
    // export tests needed the REAL CSS text embedded in the exported
    // document. `include` re-enables real processing (Vite's normal `?raw`
    // handling, which returns the file's literal text) for just those,
    // while every plain `.css` import still resolves empty exactly as a
    // blanket `false` would have.
    css: { include: [/\?raw/] },
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
        // Per-file floors — P4.7 (viewer coverage sweep). These three PDF
        // sub-components were called out in the plan's status section as
        // badly under-covered (measured via `npx vitest run src/viewers/pdf
        // --coverage --maxWorkers=1` on 2026-09-19, against this worktree's
        // HEAD @ 8bf6f0f): PdfToolbar.tsx ~29% statements, PdfThumbnailRail
        // ~4%, and the dialog the plan calls "PdfWordDialog.tsx" — which
        // doesn't exist; the actual low-coverage dialog file in
        // `src/viewers/pdf` is `PdfPasswordDialog.tsx` (a password prompt,
        // not a word-lookup dialog) — ~41%. New behavioural RTL tests now
        // bring them to (per `coverage-summary.json`, the authoritative
        // source — the v8 text-reporter's own directory table intermittently
        // drops fully-covered rows, a pre-existing reporter quirk unrelated
        // to this change): PdfToolbar 100/91.66/100/100 (stmts/branch/func/
        // line), PdfThumbnailRail 92.3/70/100/97.14, PdfPasswordDialog
        // 100/100/100/100. Floored to the next whole percent below, same as
        // the global floor above, with per-FILE glob keys (not a global
        // bump) precisely because these numbers were only measured over
        // `src/viewers/pdf` — a global bump based on this one directory
        // could false-fail CI's whole-project `npm run coverage` if other
        // concurrently-landing work shifts the rest of the codebase's
        // coverage down. Never lower — only raise once a future change
        // measurably improves one of these further.
        'src/viewers/pdf/PdfToolbar.tsx': { statements: 99, branches: 90, functions: 99, lines: 99 },
        'src/viewers/pdf/PdfThumbnailRail.tsx': { statements: 91, branches: 69, functions: 99, lines: 96 },
        'src/viewers/pdf/PdfPasswordDialog.tsx': { statements: 99, branches: 99, functions: 99, lines: 99 },
      },
    },
  },
})
