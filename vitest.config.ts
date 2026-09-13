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
  },
})
