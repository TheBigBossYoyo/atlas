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

function formatKb(bytes) {
  return `${toKb(bytes).toFixed(2)} KB`
}

function formatDelta(percent) {
  if (!Number.isFinite(percent)) return 'n/a'
  const sign = percent >= 0 ? '+' : ''
  return `${sign}${percent.toFixed(2)}%`
}

/** @returns {number} Percent change from `baseline` to `current`; `Infinity` when baseline was 0 and current isn't. */
function percentDelta(current, baseline) {
  if (baseline === 0) return current === 0 ? 0 : Infinity
  return ((current - baseline) / baseline) * 100
}

async function readBaseline() {
  try {
    return JSON.parse(await fs.readFile(baselinePath, 'utf8'))
  } catch (err) {
    if (err && err.code === 'ENOENT') {
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

  if (violations.length > 0) {
    console.error('')
    for (const message of violations) console.error(message)
    console.error(
      '\nIf this growth is intentional, refresh the baseline: node scripts/capture-bundle-baseline.mjs',
    )
    process.exit(1)
  }

  console.log(`\nBundle regression check passed (threshold ${THRESHOLD_PERCENT}%, ${currentChunks.size} chunks checked).`)
}

await main()
