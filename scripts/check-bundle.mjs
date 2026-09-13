import { gzipSync } from 'node:zlib'
import { promises as fs } from 'node:fs'
import path from 'node:path'

const projectRoot = process.cwd()
const baselinePath = path.join(projectRoot, '.sisyphus', 'baselines', 'atlas-phase0.json')
const distDir = path.join(projectRoot, 'dist')
const thresholdPercent = 25

function formatKb(value) {
  return `${value.toFixed(2)} KB`
}

function toKb(bytes) {
  return bytes / 1000
}

function formatDelta(deltaPercent) {
  const sign = deltaPercent >= 0 ? '+' : ''
  return `${sign}${deltaPercent.toFixed(2)}%`
}

async function readBaseline() {
  const raw = await fs.readFile(baselinePath, 'utf8')
  return JSON.parse(raw)
}

async function findMainChunk() {
  const indexHtmlPath = path.join(distDir, 'index.html')
  const indexHtml = await fs.readFile(indexHtmlPath, 'utf8')
  const match = indexHtml.match(/<script\s+type="module"\s+crossorigin\s+src="([^\"]+index-[^\"]+\.js)"><\/script>/)

  if (!match) {
    throw new Error(`Could not find main entry chunk in ${indexHtmlPath}`)
  }

  const relativeAssetPath = match[1].replace(/^\//, '')
  const assetPath = path.join(distDir, relativeAssetPath.replace(/^assets\//, 'assets/'))
  const content = await fs.readFile(assetPath)

  return {
    file: path.relative(projectRoot, assetPath).replace(/\\/g, '/'),
    sizeKb: toKb(content.byteLength),
    gzipKb: toKb(gzipSync(content).byteLength),
  }
}

function buildRows(baseline, current) {
  const mainDelta = ((current.sizeKb - baseline.mainChunkKB) / baseline.mainChunkKB) * 100
  const gzipDelta = ((current.gzipKb - baseline.mainChunkGzipKB) / baseline.mainChunkGzipKB) * 100

  return {
    rows: [
      {
        metric: 'main',
        baseline: formatKb(baseline.mainChunkKB),
        current: formatKb(current.sizeKb),
        delta: formatDelta(mainDelta),
      },
      {
        metric: 'gzip',
        baseline: formatKb(baseline.mainChunkGzipKB),
        current: formatKb(current.gzipKb),
        delta: formatDelta(gzipDelta),
      },
    ],
    mainDelta,
    gzipDelta,
  }
}

async function main() {
  const baseline = await readBaseline()
  const current = await findMainChunk()
  const { rows, mainDelta, gzipDelta } = buildRows(baseline, current)

  console.log(`Main chunk: ${current.file}`)
  console.table(rows)

  if (mainDelta > thresholdPercent || gzipDelta > thresholdPercent) {
    console.error(
      `Bundle regression exceeded ${thresholdPercent}% threshold (main ${formatDelta(mainDelta)}, gzip ${formatDelta(gzipDelta)}).`,
    )
    process.exit(1)
  }

  console.log(`Bundle regression check passed (threshold ${thresholdPercent}%).`)
}

await main()
