// Shared chunk-inventory logic behind both `check-bundle.mjs` (the postbuild
// regression gate, P4.3/RUN-15/QA-17) and `capture-bundle-baseline.mjs` (the
// tool that refreshes the baseline it checks against) — kept in one place so
// the two can never silently disagree about what a "chunk" is or how its
// size is measured.

import { gzipSync } from 'node:zlib'
import { promises as fs } from 'node:fs'
import path from 'node:path'

/**
 * Vite/Rollup's default output naming is `[name]-[hash].ext`, where `hash`
 * is an 8-character content hash that changes on every build even when the
 * chunk's logical content doesn't (a different hashing salt, a neighboring
 * chunk changing, etc). Stripping it yields a name that stays stable across
 * rebuilds — e.g. `rtf.js-BKGdIdS4.js` and a later build's
 * `rtf.js-9fQzR2Lm.js` both become `rtf.js.js` — so the baseline can compare
 * "the rtf.js vendor chunk" release over release instead of a filename that
 * would never match twice.
 */
const HASH_SUFFIX = /-[A-Za-z0-9_-]{8}(\.(?:js|css))$/

/** @param {string} fileName */
export function toStableChunkName(fileName) {
  return fileName.replace(HASH_SUFFIX, '$1')
}

/** @param {number} bytes */
export function toKb(bytes) {
  return bytes / 1000
}

/**
 * Reads every `.js`/`.css` file directly under `dist/assets/` and returns
 * one row per stable chunk name, summing bytes for the rare case where two
 * physical files collapse onto the same stable name.
 *
 * @param {string} projectRoot
 * @returns {Promise<Map<string, { name: string, bytes: number, gzipBytes: number, files: string[] }>>}
 */
export async function collectChunks(projectRoot) {
  const assetsDir = path.join(projectRoot, 'dist', 'assets')
  const entries = await fs.readdir(assetsDir, { withFileTypes: true })
  const chunks = new Map()

  for (const entry of entries) {
    if (!entry.isFile() || !/\.(js|css)$/.test(entry.name)) continue

    const buf = await fs.readFile(path.join(assetsDir, entry.name))
    const name = toStableChunkName(entry.name)
    const existing = chunks.get(name)
    const gzipBytes = gzipSync(buf).byteLength

    if (existing) {
      existing.bytes += buf.byteLength
      existing.gzipBytes += gzipBytes
      existing.files.push(entry.name)
    } else {
      chunks.set(name, { name, bytes: buf.byteLength, gzipBytes, files: [entry.name] })
    }
  }

  return chunks
}
