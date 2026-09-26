#!/usr/bin/env node
// Bundle regression gate (P4.3 / RUN-15 / QA-17).
//
// Wired as `postbuild` (package.json) so it runs automatically after every
// `npm run build`. Compares every `dist/assets/*.{js,css}` chunk — not just
// the main entry chunk the original Phase-0 version tracked — against
// `.sisyphus/baselines/atlas-phase3-bundle.json`, and fails the build if any
// single chunk or the total bundle size grew more than `THRESHOLD_PERCENT`.
//
// A chunk name changes its content hash on every build even when nothing
// about it changed logically, so comparisons are keyed on the
// hash-stripped "stable" chunk name (see `lib/bundleChunks.mjs`).
//
// To refresh the baseline after an intentional size change:
//   node scripts/capture-bundle-baseline.mjs

import { promises as fs } from 'node:fs'
import path from 'node:path'

import { collectChunks, toKb } from './lib/bundleChunks.mjs'

const projectRoot = process.cwd()
const baselinePath = path.join(projectRoot, '.sisyphus', 'baselines', 'atlas-phase3-bundle.json')
const THRESHOLD_PERCENT = 25

// Absolute ceilings, added 2026-09-26 after an audit of the tree.
//
// The percentage check above is the sensitive one and stays the primary gate:
// it compares against the STORED baseline (not the previous build), so slow
// creep does accumulate against a fixed reference and eventually trips 25%.
// What it cannot catch is creep across baseline REFRESHES: every intentional
// size change ends with `capture-bundle-baseline.mjs`, which moves the
// reference up, so a long series of individually-defensible +20% steps has no
// stop anywhere. These ceilings are that stop — deliberately loose, since the
// whole point is that a refresh cannot silently move them.
//
// Measured 2026-09-26 (267 chunks): total 12,507.82 KB / gzip 3,989.53 KB, and
// the largest single chunk is `rtf.js` at 2,236.44 KB — a size the project has
// explicitly decided to keep (see docs/KNOWN_LIMITATIONS.md's "the `rtf.js`
// bundle size": it is lazy-loaded, and the size is WMF/EMF rasterization that
// real legacy RTF documents genuinely need). The ceilings sit roughly 30% above
// each of those, so they are an absurdity stop rather than a second budget:
// hitting one means something went wrong in a build, or that a genuine decision
// to ship materially more code needs to be made explicitly here.
const MAX_TOTAL_BYTES = 16 * 1024 * 1024 // 16 MiB
const MAX_TOTAL_GZIP_BYTES = 5 * 1024 * 1024 // 5 MiB
const MAX_CHUNK_BYTES = 3 * 1024 * 1024 // 3 MiB — rtf.js, the known largest, is ~2.2 MiB

/**
 * @param {number} bytes
 * @returns {string}
 */
function formatKb(bytes) {
  return `${toKb(bytes).toFixed(2)} KB`
}

/**
 * @param {number} percent
 * @returns {string}
 */
function formatDelta(percent) {
  if (!Number.isFinite(percent)) return 'n/a'
  const sign = percent >= 0 ? '+' : ''
  return `${sign}${percent.toFixed(2)}%`
}

/**
 * @param {number} current
 * @param {number} baseline
 * @returns {number} Percent change from `baseline` to `current`; `Infinity` when baseline was 0 and current isn't.
 */
function percentDelta(current, baseline) {
  if (baseline === 0) return current === 0 ? 0 : Infinity
  return ((current - baseline) / baseline) * 100
}

async function readBaseline() {
  try {
    return JSON.parse(await fs.readFile(baselinePath, 'utf8'))
  } catch (err) {
    if (err instanceof Error && /** @type {NodeJS.ErrnoException} */ (err).code === 'ENOENT') {
      throw new Error(
        `No bundle baseline found at ${path.relative(projectRoot, baselinePath)}. ` +
          'Run `node scripts/capture-bundle-baseline.mjs` once to create it.',
      )
    }
    throw err
  }
}

async function main() {
  const baseline = await readBaseline()
  const currentChunks = await collectChunks(projectRoot)
  const baselineChunks = baseline.chunks ?? {}

  const rows = []
  const violations = []
  let totalBytes = 0
  let totalGzipBytes = 0

  for (const chunk of currentChunks.values()) {
    totalBytes += chunk.bytes
    totalGzipBytes += chunk.gzipBytes

    // Absolute per-chunk ceiling — checked for every chunk, including one with
    // no baseline entry, so a brand-new oversized chunk cannot slip in as
    // "(new)" and skip the percentage comparison below entirely.
    if (chunk.bytes > MAX_CHUNK_BYTES) {
      violations.push(
        `Chunk "${chunk.name}" is ${formatKb(chunk.bytes)}, over the ${formatKb(MAX_CHUNK_BYTES)} absolute per-chunk ceiling.`,
      )
    }

    const base = baselineChunks[chunk.name]
    if (!base) {
      rows.push({ chunk: chunk.name, baseline: '(new)', current: formatKb(chunk.bytes), delta: 'n/a' })
      continue
    }

    const delta = percentDelta(chunk.bytes, base.bytes)
    rows.push({
      chunk: chunk.name,
      baseline: formatKb(base.bytes),
      current: formatKb(chunk.bytes),
      delta: formatDelta(delta),
    })
    if (delta > THRESHOLD_PERCENT) {
      violations.push(
        `Chunk "${chunk.name}" grew ${formatDelta(delta)} (${formatKb(base.bytes)} -> ${formatKb(chunk.bytes)}), exceeding the ${THRESHOLD_PERCENT}% threshold.`,
      )
    }
  }

  // A baseline chunk that no longer exists (renamed, merged, or removed
  // entirely) is reported for visibility but never fails the gate on its
  // own — disappearing is never the regression this check exists to catch.
  for (const name of Object.keys(baselineChunks)) {
    if (!currentChunks.has(name)) {
      rows.push({ chunk: name, baseline: formatKb(baselineChunks[name].bytes), current: '(removed)', delta: 'n/a' })
    }
  }

  rows.sort((a, b) => a.chunk.localeCompare(b.chunk))
  console.table(rows)

  const totalDelta = percentDelta(totalBytes, baseline.totalBytes)
  const totalGzipDelta = percentDelta(totalGzipBytes, baseline.totalGzipBytes)
  console.log(
    `Total dist/assets: ${formatKb(totalBytes)} (gzip ${formatKb(totalGzipBytes)}) vs baseline ` +
      `${formatKb(baseline.totalBytes)} (gzip ${formatKb(baseline.totalGzipBytes)}) — ` +
      `${formatDelta(totalDelta)} (gzip ${formatDelta(totalGzipDelta)})`,
  )
  if (totalDelta > THRESHOLD_PERCENT || totalGzipDelta > THRESHOLD_PERCENT) {
    violations.push(
      `Total bundle size grew ${formatDelta(totalDelta)} (gzip ${formatDelta(totalGzipDelta)}), exceeding the ${THRESHOLD_PERCENT}% threshold.`,
    )
  }

  if (totalBytes > MAX_TOTAL_BYTES) {
    violations.push(
      `Total dist/assets is ${formatKb(totalBytes)}, over the ${formatKb(MAX_TOTAL_BYTES)} absolute ceiling.`,
    )
  }
  if (totalGzipBytes > MAX_TOTAL_GZIP_BYTES) {
    violations.push(
      `Total gzipped dist/assets is ${formatKb(totalGzipBytes)}, over the ${formatKb(MAX_TOTAL_GZIP_BYTES)} absolute ceiling.`,
    )
  }

  if (violations.length > 0) {
    console.error('')
    for (const message of violations) console.error(message)
    // Two different fixes, so say which one applies: refreshing the baseline
    // clears a percentage violation but does nothing for an absolute ceiling
    // (that is the point of the ceiling), and following the wrong hint just
    // produces a second failing build.
    if (violations.some((message) => message.includes('ceiling'))) {
      console.error(
        '\nAn absolute ceiling cannot be cleared by refreshing the baseline — either bring the size back' +
          '\ndown, or raise the ceiling in scripts/check-bundle.mjs as a deliberate, reviewed decision.',
      )
    }
    if (violations.some((message) => message.includes('threshold'))) {
      console.error(
        '\nIf this growth is intentional, refresh the baseline: node scripts/capture-bundle-baseline.mjs',
      )
    }
    process.exit(1)
  }

  console.log(
    `\nBundle regression check passed (threshold ${THRESHOLD_PERCENT}%, ceilings ` +
      `${formatKb(MAX_TOTAL_BYTES)} total / ${formatKb(MAX_TOTAL_GZIP_BYTES)} gzip / ` +
      `${formatKb(MAX_CHUNK_BYTES)} per chunk, ${currentChunks.size} chunks checked).`,
  )
}

await main()
