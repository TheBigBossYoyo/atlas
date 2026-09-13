import fs from 'node:fs'
import path from 'node:path'
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// This worktree's `node_modules` is a directory junction into a shared
// install living outside the project root (see the repo's worktree
// conventions). Vite's dev server resolves symlinks/junctions before
// applying its `server.fs` allow-list, so without this the junction's real
// target is rejected as "outside of Vite serving allow list" (a plain 403,
// not a CSP violation) for any asset requested by literal URL rather than
// pre-bundled — e.g. pdf.js's worker script. Harmless to resolve eagerly:
// falls back to the project root alone if the path doesn't exist yet.
function resolveNodeModulesAllowRoot(): string | undefined {
  const nodeModulesPath = path.resolve(__dirname, 'node_modules')
  try {
    return fs.realpathSync(nodeModulesPath)
  } catch {
    return undefined
  }
}

const nodeModulesRealPath = resolveNodeModulesAllowRoot()

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  base: './',
  server: {
    fs: {
      allow: [__dirname, ...(nodeModulesRealPath ? [nodeModulesRealPath] : [])],
    },
  },
})
