import { getCachedMetrics, setCachedMetrics } from './cache'
import { resolveFontFamily, type FontVariant } from './families'
import { buildMetrics, type FontMetrics } from './metrics'
import { parseTtf } from './ttf'

export async function loadFontMetrics(family: string, variant: FontVariant): Promise<FontMetrics> {
  const resolvedFamily = resolveFontFamily(family)
  if (!resolvedFamily) {
    throw new Error(`Unsupported font family: ${family}`)
  }

  const fileUrl = resolvedFamily.files[variant]
  const cacheKey = `${resolvedFamily.substituteName}@${variant}@${hashString(fileUrl)}`
  const cachedMetrics = await getCachedMetrics(cacheKey)
  if (cachedMetrics) {
    return cachedMetrics
  }

  const response = await fetch(fileUrl)
  if (!response.ok) {
    throw new Error(`Failed to fetch font file: ${fileUrl}`)
  }

  const buffer = await response.arrayBuffer()
  const tables = parseTtf(buffer)
  const metrics = buildMetrics(tables)
  await setCachedMetrics(cacheKey, metrics)
  return metrics
}

function hashString(value: string): string {
  let hash = 0x811c9dc5
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index)
    hash = Math.imul(hash, 0x01000193)
  }

  return (hash >>> 0).toString(16)
}
