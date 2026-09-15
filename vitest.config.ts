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
    // `css: false` disables Vitest's CSS pipeline for ordinary side-effect
    // `.css` imports (component styles the tests don't need real values
    // for) — but Vitest's own css-disable plugin matches on bare `.css`
    // extension and ignores the `?raw`/`?inline` query, so a blanket
    // `false` was ALSO forcing every `?raw` stylesheet import
    // (`pdf.ts`/`slidesPdf.ts`'s harvested `viewer-docx.css`, `katex/dist/
    // katex.min.css`, etc.) to resolve as an empty string in every test —
    // silently, since nothing asserted on that content until X1's PDF
    // export tests needed the REAL CSS text embedded in the exported
    // document. `include` re-enables real processing (Vite's normal `?raw`
    // handling, which returns the file's literal text) for just those.
    css: { include: [/\?raw/] },
  },
})
