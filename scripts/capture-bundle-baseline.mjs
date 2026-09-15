#!/usr/bin/env node
// Captures a fresh bundle-size baseline from the current `dist/` build for
// `check-bundle.mjs` (P4.3/RUN-15/QA-17) to compare future builds against.
//
// Run this deliberately after `npm run build`, and only when a bundle-size
// change is intentional (a new format viewer, a swapped library, etc.) —
// never as a way to make a failing `check-bundle.mjs` pass without looking
// at why it grew.
//
//   node scripts/capture-bundle-baseline.mjs

import { promises as fs } from 'node:fs'
import path from 'node:path'

import { collectChunks, toKb } from './lib/bundleChunks.mjs'

const projectRoot = process.cwd()
const baselinePath = path.join(projectRoot, '.sisyphus', 'baselines', 'atlas-phase3-bundle.json')

async function main() {
  const chunks = await collectChunks(projectRoot)

  /** @type {Record<string, { bytes: number, gzipBytes: number }>} */
  const chunkBaseline = {}
  let totalBytes = 0
  let totalGzipBytes = 0

  for (const chunk of chunks.values()) {
    chunkBaseline[chunk.name] = { bytes: chunk.bytes, gzipBytes: chunk.gzipBytes }
    totalBytes += chunk.bytes
    totalGzipBytes += chunk.gzipBytes
  }

  const baseline = {
    capturedAt: new Date().toISOString().slice(0, 10),
    packageVersion: JSON.parse(await fs.readFile(path.join(projectRoot, 'package.json'), 'utf8')).version,
    chunkCount: chunks.size,
    totalBytes,
    totalGzipBytes,
    chunks: chunkBaseline,
    notes:
      'Phase 3 baseline (post wave1+wave2), captured by scripts/capture-bundle-baseline.mjs. ' +
      'Superseded atlas-phase0.json, which only tracked the single main entry chunk and predated ' +
      'the DOCX/PDF/slides/spreadsheet viewer engines and full shiki grammar set. Chunk names are ' +
      'stripped of their Vite content hash (see bundleChunks.mjs) so they compare release over release.',
  }

  await fs.mkdir(path.dirname(baselinePath), { recursive: true })
  await fs.writeFile(baselinePath, `${JSON.stringify(baseline, null, 2)}\n`)

  console.log(
    `Captured baseline for ${chunks.size} chunks (total ${toKb(totalBytes).toFixed(2)} KB, ` +
      `gzip ${toKb(totalGzipBytes).toFixed(2)} KB) -> ${path.relative(projectRoot, baselinePath)}`,
  )
}

await main()
